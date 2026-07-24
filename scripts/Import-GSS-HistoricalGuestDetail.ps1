[CmdletBinding()]
param(
    [ValidateSet('Plan', 'Apply', 'Rollback')]
    [string]$Operation = 'Plan',

    [Parameter(Mandatory)]
    [string]$FolderPath,

    [Parameter(Mandatory)]
    [string]$ManifestPath,

    [string[]]$SourcePath = @(),

    [string]$LedgerPath,

    [int]$MutexTimeoutSeconds = 60,

    [ValidateSet('', 'AfterLedgerBaseline', 'AfterFirstPublish', 'AfterFirstRollbackHide')]
    [string]$TestFailurePoint = '',

    [switch]$OutputObject
)

$ErrorActionPreference = 'Stop'

$recoveryScriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $recoveryScriptRoot 'Gss-EmailPackage.ps1')

$script:GssHistoricalRecoveryManifestVersion = 'gss-historical-recovery/v1'
$script:GssHistoricalRecoveryReceiptVersion = 'gss-historical-recovery-receipt/v1'
$script:GssHistoricalResponseIdentityVersion = 'gss-feedback-response-identity/v1'
$script:GssHistoricalResponseSetVersion = 'gss-historical-response-set/v1'
$script:GssHistoricalRecoveryClassification = 'CONTAINS PERSONAL DATA ' + [char]0x2014 + ' RESTRICTED'
$script:GssHistoricalRecoveryArchivePrefix = '03 Uploaded Survey Workbooks/Archive - Previous Uploads/Recovered Historical Detail'
$script:GssHistoricalRecoveryMutexName = 'Global\GSSSurveyWorkbookAutomationTransaction'

function Write-GssHistoricalRecoveryAtomicJson {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$Value
    )

    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    $temporary = Join-Path $directory ".gss-json-$([guid]::NewGuid().ToString('N').Substring(0, 12)).tmp"
    try {
        $json = $Value | ConvertTo-Json -Depth 20
        $encoding = New-Object System.Text.UTF8Encoding($false, $true)
        [System.IO.File]::WriteAllText($temporary, $json + [Environment]::NewLine, $encoding)
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporary -PathType Leaf) {
            Remove-Item -LiteralPath $temporary -Force
        }
    }
}

function Get-GssHistoricalRecoveryProperty {
    param(
        [object]$Value,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $Value) { return $null }
    $property = $Value.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Assert-GssHistoricalRecoveryPropertySet {
    param(
        [Parameter(Mandatory)][object]$Value,
        [Parameter(Mandatory)][string[]]$Required,
        [Parameter(Mandatory)][string[]]$Allowed,
        [Parameter(Mandatory)][string]$Label
    )

    $names = @($Value.PSObject.Properties.Name)
    foreach ($name in $Required) {
        if ($names -notcontains $name) {
            throw "$Label is missing required property '$name'."
        }
    }
    $unexpected = @($names | Where-Object { $_ -notin $Allowed })
    if ($unexpected.Count -gt 0) {
        throw "$Label contains unsupported property '$($unexpected[0])'. The recovery manifest is intentionally PII-free and schema-closed."
    }
}

function Assert-GssHistoricalRecoveryPathUnder {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Label
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    $fullRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    if (-not $fullPath.StartsWith("$fullRoot\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label must remain under '$fullRoot': $fullPath"
    }
    return $fullPath
}

function ConvertTo-GssHistoricalRecoveryDateText {
    param(
        [Parameter(Mandatory)][object]$Value,
        [Parameter(Mandatory)][string]$Label
    )

    $date = [datetime]::MinValue
    $text = [string]$Value
    if (-not [datetime]::TryParseExact(
        $text,
        'yyyy-MM-dd',
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$date
    )) {
        throw "$Label must use yyyy-MM-dd: '$text'."
    }
    return $date.ToString('yyyy-MM-dd')
}

function Get-GssHistoricalResponseSetSha256 {
    param([Parameter(Mandatory)][string[]]$ResponseHashes)

    $hashes = @($ResponseHashes | ForEach-Object { ([string]$_).ToLowerInvariant() } | Sort-Object -Unique)
    if ($hashes.Count -eq 0) {
        throw 'A historical recovery response set cannot be empty.'
    }
    $material = "$($script:GssHistoricalResponseSetVersion)`n" + ($hashes -join "`n")
    return Get-GssStringSha256 $material
}

function Get-GssHistoricalRecoveryExpectedDestination {
    param(
        [Parameter(Mandatory)][string]$FiscalYear,
        [Parameter(Mandatory)][string]$ReportWeek,
        [Parameter(Mandatory)][string]$Sha256
    )

    $weekSlug = $ReportWeek.Replace(' ', '-')
    return "$($script:GssHistoricalRecoveryArchivePrefix)/$FiscalYear/$weekSlug-$Sha256.xlsx"
}

function Read-GssHistoricalRecoveryManifest {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$GssRoot
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Historical recovery manifest is missing: $Path"
    }
    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    try {
        $manifest = Get-Content -Raw -LiteralPath $resolvedPath | ConvertFrom-Json
    }
    catch {
        throw "Historical recovery manifest is not valid JSON: $resolvedPath. $($_.Exception.Message)"
    }

    Assert-GssHistoricalRecoveryPropertySet `
        -Value $manifest `
        -Required @('schema_version', 'fiscal_year', 'sources') `
        -Allowed @('schema_version', 'fiscal_year', 'created_at_utc', 'generated_at_utc', 'sources') `
        -Label 'Historical recovery manifest'

    if ([string]$manifest.schema_version -ne $script:GssHistoricalRecoveryManifestVersion) {
        throw "Unsupported historical recovery manifest version: $($manifest.schema_version)"
    }
    $fiscalYear = [string]$manifest.fiscal_year
    if ($fiscalYear -notmatch '^FY\d{2}$') {
        throw "Historical recovery fiscal_year is invalid: '$fiscalYear'."
    }
    foreach ($timestampName in @('created_at_utc', 'generated_at_utc')) {
        $timestampValue = Get-GssHistoricalRecoveryProperty -Value $manifest -Name $timestampName
        if ($null -eq $timestampValue) { continue }
        if ($timestampValue -is [datetime]) {
            if ($timestampValue.Kind -ne [System.DateTimeKind]::Utc) {
                throw "Historical recovery manifest $timestampName must be an ISO-8601 UTC timestamp."
            }
            continue
        }
        if ($timestampValue -is [datetimeoffset]) {
            if ($timestampValue.Offset -ne [timespan]::Zero) {
                throw "Historical recovery manifest $timestampName must be an ISO-8601 UTC timestamp."
            }
            continue
        }
        $timestampText = [string]$timestampValue
        $timestamp = [datetimeoffset]::MinValue
        if ($timestampText -notmatch 'Z$' -or
            -not [datetimeoffset]::TryParse(
                $timestampText,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::AssumeUniversal,
                [ref]$timestamp
            )) {
            throw "Historical recovery manifest $timestampName must be an ISO-8601 UTC timestamp."
        }
    }

    $sourceObjects = @($manifest.sources)
    if ($sourceObjects.Count -eq 0) {
        throw 'Historical recovery manifest contains no sources.'
    }

    $requiredSourceProperties = @(
        'source_kind',
        'drive_file_id',
        'gmail_message_ids',
        'subject_week',
        'source_report_week',
        'assignment_basis',
        'visit_date_start',
        'visit_date_end',
        'row_count',
        'byte_size',
        'sha256',
        'response_set_sha256',
        'destination_path',
        'validation'
    )
    $normalizedSources = @()
    $sourceIndex = 0
    foreach ($source in $sourceObjects) {
        $sourceIndex++
        $label = "Historical recovery source $sourceIndex"
        Assert-GssHistoricalRecoveryPropertySet `
            -Value $source `
            -Required $requiredSourceProperties `
            -Allowed $requiredSourceProperties `
            -Label $label

        $sourceKind = [string]$source.source_kind
        if ($sourceKind -notmatch '^[a-z0-9][a-z0-9_+-]{0,63}$') {
            throw "$label source_kind is invalid: '$sourceKind'."
        }

        $driveFileId = [string]$source.drive_file_id
        if (-not [string]::IsNullOrWhiteSpace($driveFileId) -and $driveFileId -notmatch '^[A-Za-z0-9_-]{10,256}$') {
            throw "$label drive_file_id is invalid."
        }
        $gmailMessageIds = @($source.gmail_message_ids | ForEach-Object { [string]$_ })
        foreach ($gmailMessageId in $gmailMessageIds) {
            if ($gmailMessageId -notmatch '^[A-Za-z0-9_-]{8,256}$') {
                throw "$label contains an invalid Gmail message ID."
            }
        }
        if ([string]::IsNullOrWhiteSpace($driveFileId) -and $gmailMessageIds.Count -eq 0) {
            throw "$label must include a Drive file ID or at least one Gmail message ID."
        }

        $subjectWeek = [string]$source.subject_week
        if (-not [string]::IsNullOrWhiteSpace($subjectWeek) -and $subjectWeek -notmatch '^FY\d{2} FW(?:[1-9]|[1-4]\d|5[0-3])$') {
            throw "$label subject_week is invalid: '$subjectWeek'."
        }
        $reportWeek = [string]$source.source_report_week
        if ($reportWeek -notmatch '^FY\d{2} FW(?:[1-9]|[1-4]\d|5[0-3])$') {
            throw "$label source_report_week is invalid: '$reportWeek'."
        }
        if (-not $reportWeek.StartsWith("$fiscalYear ", [System.StringComparison]::Ordinal)) {
            throw "$label source_report_week '$reportWeek' is outside manifest fiscal year '$fiscalYear'."
        }

        $assignmentBasis = [string]$source.assignment_basis
        if ($assignmentBasis -notmatch '^[a-z0-9][a-z0-9_-]{0,127}$') {
            throw "$label assignment_basis must be a lowercase provenance slug."
        }

        $visitDateStart = ConvertTo-GssHistoricalRecoveryDateText $source.visit_date_start "$label visit_date_start"
        $visitDateEnd = ConvertTo-GssHistoricalRecoveryDateText $source.visit_date_end "$label visit_date_end"
        if ([datetime]$visitDateEnd -lt [datetime]$visitDateStart) {
            throw "$label visit_date_end precedes visit_date_start."
        }

        try {
            $rowCountNumber = [double]$source.row_count
            $rowCount = [int]$source.row_count
        }
        catch { throw "$label row_count must be an integer." }
        if ($rowCountNumber -ne [math]::Floor($rowCountNumber) -or $rowCount -lt 1) {
            throw "$label row_count must be an integer of at least 1."
        }
        try {
            $byteSizeNumber = [decimal]$source.byte_size
            $byteSize = [long]$source.byte_size
        }
        catch { throw "$label byte_size must be an integer." }
        if ($byteSizeNumber -ne [math]::Floor($byteSizeNumber) -or $byteSize -lt 1) {
            throw "$label byte_size must be an integer of at least 1."
        }

        $sha256 = [string]$source.sha256
        if ($sha256 -notmatch '^[a-f0-9]{64}$') {
            throw "$label sha256 must be a lowercase SHA-256 hash."
        }
        $responseSetSha256 = [string]$source.response_set_sha256
        if ($responseSetSha256 -notmatch '^[a-f0-9]{64}$') {
            throw "$label response_set_sha256 must be a lowercase SHA-256 hash."
        }

        Assert-GssHistoricalRecoveryPropertySet `
            -Value $source.validation `
            -Required @('detail_schema_valid', 'response_identity_version', 'duplicate_response_count') `
            -Allowed @('detail_schema_valid', 'response_identity_version', 'duplicate_response_count') `
            -Label "$label validation"
        if (-not [bool]$source.validation.detail_schema_valid) {
            throw "$label validation must assert detail_schema_valid=true."
        }
        if ([string]$source.validation.response_identity_version -ne $script:GssHistoricalResponseIdentityVersion) {
            throw "$label validation response_identity_version is unsupported: '$($source.validation.response_identity_version)'."
        }
        try {
            $duplicateResponseCountNumber = [double]$source.validation.duplicate_response_count
            $duplicateResponseCount = [int]$source.validation.duplicate_response_count
        }
        catch { throw "$label validation duplicate_response_count must be an integer." }
        if ($duplicateResponseCountNumber -ne [math]::Floor($duplicateResponseCountNumber) -or
            $duplicateResponseCount -lt 0 -or
            $duplicateResponseCount -ge $rowCount) {
            throw "$label validation duplicate_response_count is outside the valid range."
        }

        $expectedDestination = Get-GssHistoricalRecoveryExpectedDestination `
            -FiscalYear $fiscalYear `
            -ReportWeek $reportWeek `
            -Sha256 $sha256
        $destination = ([string]$source.destination_path).Replace('\', '/').Trim('/')
        if (-not $destination.Equals($expectedDestination, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "$label destination_path must be collision-proof and exact: $expectedDestination"
        }
        $destinationFullPath = Resolve-GssDropboxPath -Path $destination -FolderPath $GssRoot

        $normalizedSources += [pscustomobject][ordered]@{
            SourceIndex = $sourceIndex
            SourceKind = $sourceKind
            DriveFileId = $driveFileId
            GmailMessageIds = $gmailMessageIds
            SubjectWeek = $subjectWeek
            SourceReportWeek = $reportWeek
            AssignmentBasis = $assignmentBasis
            VisitDateStart = $visitDateStart
            VisitDateEnd = $visitDateEnd
            RowCount = $rowCount
            ByteSize = $byteSize
            Sha256 = $sha256
            ResponseSetSha256 = $responseSetSha256
            DuplicateResponseCount = $duplicateResponseCount
            DestinationPath = $destination
            DestinationFullPath = $destinationFullPath
        }
    }

    $duplicateFileHashes = @($normalizedSources | Group-Object Sha256 | Where-Object Count -gt 1)
    if ($duplicateFileHashes.Count -gt 0) {
        throw "Historical recovery manifest repeats source SHA-256 '$($duplicateFileHashes[0].Name)'. Consolidate exact duplicate attachments before import."
    }
    $duplicateDestinations = @($normalizedSources | Group-Object { $_.DestinationPath.ToLowerInvariant() } | Where-Object Count -gt 1)
    if ($duplicateDestinations.Count -gt 0) {
        throw "Historical recovery manifest repeats destination '$($duplicateDestinations[0].Group[0].DestinationPath)'."
    }
    foreach ($weekGroup in @($normalizedSources | Group-Object SourceReportWeek | Where-Object Count -gt 1)) {
        $responseSets = @($weekGroup.Group.ResponseSetSha256 | Sort-Object -Unique)
        if ($responseSets.Count -gt 1) {
            throw "Historical recovery manifest has conflicting response sets for source_report_week '$($weekGroup.Name)'. Quarantine the week for chronology review."
        }
        throw "Historical recovery manifest repeats the exact response set for source_report_week '$($weekGroup.Name)'. Consolidate duplicate attachment provenance into one source entry."
    }

    return [pscustomobject][ordered]@{
        Path = $resolvedPath
        Sha256 = Get-GssSha256 $resolvedPath
        FiscalYear = $fiscalYear
        Sources = $normalizedSources
    }
}

function Read-GssHistoricalRecoveryWorkbookEvidence {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$Source,
        [Parameter(Mandatory)][string]$GssRoot,
        [Parameter(Mandatory)][string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label is missing: $Path"
    }
    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    $before = Get-Item -LiteralPath $resolvedPath
    if ([long]$before.Length -ne [long]$Source.ByteSize) {
        throw "$Label byte size mismatch. Expected $($Source.ByteSize); actual $($before.Length)."
    }
    $hashBefore = Get-GssSha256 $resolvedPath
    if ($hashBefore -ne [string]$Source.Sha256) {
        throw "$Label SHA-256 mismatch. Expected $($Source.Sha256); actual $hashBefore."
    }

    $workbook = Read-GssDetailWorkbook -Path $resolvedPath -FolderPath $GssRoot

    $after = Get-Item -LiteralPath $resolvedPath
    $hashAfter = Get-GssSha256 $resolvedPath
    if ([long]$after.Length -ne [long]$before.Length -or $hashAfter -ne $hashBefore) {
        throw "$Label changed while it was being validated."
    }

    $responses = @($workbook.Responses)
    if ($responses.Count -ne [int]$Source.RowCount) {
        throw "$Label row count mismatch. Expected $($Source.RowCount); actual $($responses.Count)."
    }
    $responseHashes = @($responses.ResponseHash | Sort-Object -Unique)
    $duplicateCount = $responses.Count - $responseHashes.Count
    if ($duplicateCount -ne [int]$Source.DuplicateResponseCount) {
        throw "$Label duplicate response count mismatch. Expected $($Source.DuplicateResponseCount); actual $duplicateCount."
    }
    $responseSetSha256 = Get-GssHistoricalResponseSetSha256 $responseHashes
    if ($responseSetSha256 -ne [string]$Source.ResponseSetSha256) {
        throw "$Label response-set SHA-256 mismatch. Expected $($Source.ResponseSetSha256); actual $responseSetSha256."
    }
    $visitDateStart = $workbook.VisitDateStart.ToString('yyyy-MM-dd')
    $visitDateEnd = $workbook.VisitDateEnd.ToString('yyyy-MM-dd')
    if ($visitDateStart -ne [string]$Source.VisitDateStart -or $visitDateEnd -ne [string]$Source.VisitDateEnd) {
        throw "$Label visit-date range mismatch. Expected $($Source.VisitDateStart)..$($Source.VisitDateEnd); actual $visitDateStart..$visitDateEnd."
    }

    $responseEntries = @($responses | Sort-Object ResponseHash -Unique | ForEach-Object {
        [pscustomobject][ordered]@{
            response_hash = [string]$_.ResponseHash
            first_seen_reporting_date = $_.VisitDate.ToString('yyyy-MM-dd')
        }
    })
    return [pscustomobject][ordered]@{
        Path = $resolvedPath
        Sha256 = $hashAfter
        ByteSize = [long]$after.Length
        RowCount = $responses.Count
        UniqueResponseCount = $responseHashes.Count
        DuplicateResponseCount = $duplicateCount
        ResponseSetSha256 = $responseSetSha256
        VisitDateStart = $visitDateStart
        VisitDateEnd = $visitDateEnd
        ResponseEntries = $responseEntries
    }
}

function Resolve-GssHistoricalRecoverySourceEvidence {
    param(
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][string[]]$Paths,
        [Parameter(Mandatory)][string]$GssRoot,
        [Parameter(Mandatory)][string]$RuntimeRoot
    )

    if ($Paths.Count -ne @($Manifest.Sources).Count) {
        throw "Exactly $(@($Manifest.Sources).Count) staged source path(s) are required; received $($Paths.Count)."
    }
    $stagingRoot = Join-Path $RuntimeRoot 'staging'
    $byHash = @{}
    foreach ($pathValue in $Paths) {
        $candidate = if ([System.IO.Path]::IsPathRooted($pathValue)) {
            [System.IO.Path]::GetFullPath($pathValue)
        }
        else {
            Resolve-GssDropboxPath -Path $pathValue -FolderPath $GssRoot
        }
        $resolved = Assert-GssHistoricalRecoveryPathUnder -Path $candidate -Root $stagingRoot -Label 'Historical recovery source'
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "Historical recovery staged source is missing: $resolved"
        }
        $hash = Get-GssSha256 $resolved
        if ($byHash.ContainsKey($hash)) {
            throw "Historical recovery source paths repeat SHA-256 '$hash'."
        }
        $byHash[$hash] = $resolved
    }

    $evidence = @()
    foreach ($source in $Manifest.Sources) {
        $path = $byHash[[string]$source.Sha256]
        if ([string]::IsNullOrWhiteSpace([string]$path)) {
            throw "No staged source matches manifest SHA-256 '$($source.Sha256)'."
        }
        $evidence += Read-GssHistoricalRecoveryWorkbookEvidence `
            -Path $path `
            -Source $source `
            -GssRoot $GssRoot `
            -Label "Historical recovery source $($source.SourceIndex)"
    }
    return $evidence
}

function Get-GssHistoricalRecoveryPlanEntry {
    param([Parameter(Mandatory)][object[]]$Evidence)

    $byHash = @{}
    foreach ($entry in @($Evidence.ResponseEntries)) {
        $hash = [string]$entry.response_hash
        $date = [string]$entry.first_seen_reporting_date
        if ($byHash.ContainsKey($hash) -and [string]$byHash[$hash].first_seen_reporting_date -ne $date) {
            throw "Historical response hash '$hash' has conflicting visit dates."
        }
        $byHash[$hash] = [pscustomobject][ordered]@{
            response_hash = $hash
            first_seen_reporting_date = $date
        }
    }
    return @($byHash.Values | Sort-Object response_hash)
}

function Get-GssHistoricalRecoveryPath {
    param(
        [Parameter(Mandatory)][string]$GssRoot,
        [Parameter(Mandatory)][string]$ManifestSha256,
        [string]$RequestedLedgerPath
    )

    $runtimeRoot = Join-Path $GssRoot '_automation_runs'
    $stateRoot = Join-Path $runtimeRoot 'state'
    $ledger = if ([string]::IsNullOrWhiteSpace($RequestedLedgerPath)) {
        Join-Path $stateRoot 'gss_feedback_first_seen.json'
    }
    elseif ([System.IO.Path]::IsPathRooted($RequestedLedgerPath)) {
        [System.IO.Path]::GetFullPath($RequestedLedgerPath)
    }
    else {
        Resolve-GssDropboxPath -Path $RequestedLedgerPath -FolderPath $GssRoot
    }
    $ledger = Assert-GssHistoricalRecoveryPathUnder -Path $ledger -Root $stateRoot -Label 'Historical recovery ledger'

    $transactionRoot = Join-Path (Join-Path $runtimeRoot 'historical-recovery') $ManifestSha256
    return [pscustomobject][ordered]@{
        RuntimeRoot = $runtimeRoot
        StateRoot = $stateRoot
        LedgerPath = $ledger
        TransactionRoot = $transactionRoot
        ReceiptPath = Join-Path $transactionRoot 'transaction-receipt.json'
        ManifestSnapshotPath = Join-Path $transactionRoot 'recovery-manifest.json'
    }
}

function Copy-GssHistoricalRecoveryExactFile {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$ExpectedSha256,
        [Parameter(Mandatory)][long]$ExpectedByteSize
    )

    $directory = Split-Path -Parent $Destination
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    if (Test-Path -LiteralPath $Destination -PathType Leaf) {
        $existing = Get-Item -LiteralPath $Destination
        if ([long]$existing.Length -ne $ExpectedByteSize -or (Get-GssSha256 $Destination) -ne $ExpectedSha256) {
            throw "Existing transaction file does not match the reviewed source: $Destination"
        }
        return
    }

    $temporary = Join-Path $directory ".gss-copy-$([guid]::NewGuid().ToString('N').Substring(0, 12)).tmp"
    try {
        [System.IO.File]::Copy($Source, $temporary, $false)
        $copy = Get-Item -LiteralPath $temporary
        if ([long]$copy.Length -ne $ExpectedByteSize -or (Get-GssSha256 $temporary) -ne $ExpectedSha256) {
            throw "Copied transaction file failed exact readback: $Destination"
        }
        [System.IO.File]::Move($temporary, $Destination)
    }
    finally {
        if (Test-Path -LiteralPath $temporary -PathType Leaf) {
            Remove-Item -LiteralPath $temporary -Force
        }
    }
}

function Write-GssHistoricalRecoveryReceipt {
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][string]$Path
    )

    $Receipt.updated_at_utc = (Get-Date).ToUniversalTime().ToString('o')
    Write-GssHistoricalRecoveryAtomicJson -Path $Path -Value $Receipt
}

function New-GssHistoricalRecoveryReceipt {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][object[]]$Evidence,
        [Parameter(Mandatory)][object]$Paths
    )

    if (-not $PSCmdlet.ShouldProcess($Paths.TransactionRoot, 'Initialize historical recovery receipt and ledger backup')) {
        throw 'Historical recovery receipt initialization was declined.'
    }
    if (-not (Test-Path -LiteralPath $Paths.TransactionRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $Paths.TransactionRoot -Force | Out-Null
    }
    Copy-GssHistoricalRecoveryExactFile `
        -Source $Manifest.Path `
        -Destination $Paths.ManifestSnapshotPath `
        -ExpectedSha256 $Manifest.Sha256 `
        -ExpectedByteSize (Get-Item -LiteralPath $Manifest.Path).Length

    $ledger = Read-GssFeedbackLedger $Paths.LedgerPath
    $ledgerByHash = @{}
    foreach ($entry in @($ledger.entries)) {
        $ledgerByHash[[string]$entry.response_hash] = $entry
    }
    $plannedEntries = Get-GssHistoricalRecoveryPlanEntry $Evidence
    $plannedInsertions = @($plannedEntries | Where-Object { -not $ledgerByHash.ContainsKey([string]$_.response_hash) })

    $ledgerExisted = Test-Path -LiteralPath $Paths.LedgerPath -PathType Leaf
    $ledgerShaBefore = if ($ledgerExisted) { Get-GssSha256 $Paths.LedgerPath } else { '' }
    $ledgerBackupPath = Join-Path $Paths.TransactionRoot 'ledger-original-attempt-1.json'
    if ($ledgerExisted) {
        Copy-GssHistoricalRecoveryExactFile `
            -Source $Paths.LedgerPath `
            -Destination $ledgerBackupPath `
            -ExpectedSha256 $ledgerShaBefore `
            -ExpectedByteSize (Get-Item -LiteralPath $Paths.LedgerPath).Length
    }

    $fileRecords = @()
    for ($index = 0; $index -lt @($Manifest.Sources).Count; $index++) {
        $source = $Manifest.Sources[$index]
        $sourceEvidence = $Evidence[$index]
        $destinationExists = Test-Path -LiteralPath $source.DestinationFullPath -PathType Leaf
        $partialName = ".$($source.Sha256).recovery-part"
        $fileRecords += [pscustomobject][ordered]@{
            source_index = [int]$source.SourceIndex
            sha256 = [string]$source.Sha256
            byte_size = [long]$source.ByteSize
            row_count = [int]$source.RowCount
            response_set_sha256 = [string]$source.ResponseSetSha256
            destination_path = [string]$source.DestinationPath
            destination_full_path = [string]$source.DestinationFullPath
            partial_path = Join-Path (Split-Path -Parent $source.DestinationFullPath) $partialName
            source_leaf_name = Split-Path -Leaf $sourceEvidence.Path
            response_hashes = @($sourceEvidence.ResponseEntries.response_hash | Sort-Object -Unique)
            prepared = $false
            destination_preexisting = [bool]$destinationExists
            published_by_transaction = $false
        }
    }

    $now = (Get-Date).ToUniversalTime().ToString('o')
    return [pscustomobject][ordered]@{
        schema_version = $script:GssHistoricalRecoveryReceiptVersion
        classification = $script:GssHistoricalRecoveryClassification
        contains_personal_data = $true
        manifest_sha256 = [string]$Manifest.Sha256
        manifest_snapshot_path = [string]$Paths.ManifestSnapshotPath
        transaction_id = "historical-recovery:$($Manifest.Sha256)"
        state = 'Validated'
        attempt = 1
        created_at_utc = $now
        updated_at_utc = $now
        ledger_path = [string]$Paths.LedgerPath
        ledger_existed_before = [bool]$ledgerExisted
        ledger_sha256_before = [string]$ledgerShaBefore
        ledger_backup_path = if ($ledgerExisted) { $ledgerBackupPath } else { '' }
        ledger_sha256_after = ''
        planned_entries = $plannedEntries
        planned_insertions = $plannedInsertions
        inserted_response_hashes = @()
        files = $fileRecords
        published_file_count = 0
        error = ''
    }
}

function Assert-GssHistoricalRecoveryReceipt {
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][object]$Paths
    )

    if ([string]$Receipt.schema_version -ne $script:GssHistoricalRecoveryReceiptVersion) {
        throw "Unsupported historical recovery receipt version: $($Receipt.schema_version)"
    }
    if ([string]$Receipt.manifest_sha256 -ne [string]$Manifest.Sha256) {
        throw 'Historical recovery receipt belongs to a different manifest fingerprint.'
    }
    if ([string]$Receipt.transaction_id -ne "historical-recovery:$($Manifest.Sha256)") {
        throw 'Historical recovery receipt transaction ID is invalid.'
    }
    $receiptManifestSnapshot = [System.IO.Path]::GetFullPath([string]$Receipt.manifest_snapshot_path)
    if (-not $receiptManifestSnapshot.Equals([System.IO.Path]::GetFullPath($Paths.ManifestSnapshotPath), [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Historical recovery receipt manifest snapshot path is outside this transaction.'
    }
    $receiptLedgerPath = [System.IO.Path]::GetFullPath([string]$Receipt.ledger_path)
    if (-not $receiptLedgerPath.Equals([System.IO.Path]::GetFullPath($Paths.LedgerPath), [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Historical recovery receipt ledger path does not match this invocation.'
    }
    if (@($Receipt.files).Count -ne @($Manifest.Sources).Count) {
        throw 'Historical recovery receipt source count does not match the manifest.'
    }
    foreach ($source in $Manifest.Sources) {
        $file = @($Receipt.files | Where-Object { [int]$_.source_index -eq [int]$source.SourceIndex })
        $expectedPartial = Join-Path (Split-Path -Parent $source.DestinationFullPath) ".$($source.Sha256).recovery-part"
        if ($file.Count -ne 1 -or
            [string]$file[0].sha256 -ne [string]$source.Sha256 -or
            [long]$file[0].byte_size -ne [long]$source.ByteSize -or
            [int]$file[0].row_count -ne [int]$source.RowCount -or
            [string]$file[0].response_set_sha256 -ne [string]$source.ResponseSetSha256 -or
            [string]$file[0].destination_path -ne [string]$source.DestinationPath -or
            -not ([System.IO.Path]::GetFullPath([string]$file[0].destination_full_path)).Equals(
                [System.IO.Path]::GetFullPath([string]$source.DestinationFullPath),
                [System.StringComparison]::OrdinalIgnoreCase
            ) -or
            -not ([System.IO.Path]::GetFullPath([string]$file[0].partial_path)).Equals(
                [System.IO.Path]::GetFullPath($expectedPartial),
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
            throw "Historical recovery receipt source $($source.SourceIndex) does not match the manifest."
        }
    }
    if (-not (Test-Path -LiteralPath $Receipt.manifest_snapshot_path -PathType Leaf) -or
        (Get-GssSha256 $Receipt.manifest_snapshot_path) -ne [string]$Manifest.Sha256) {
        throw 'Historical recovery manifest snapshot is missing or changed.'
    }
}

function Reset-GssHistoricalRecoveryReceipt {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][object[]]$Evidence,
        [Parameter(Mandatory)][object]$Paths
    )

    if (-not $PSCmdlet.ShouldProcess($Paths.ReceiptPath, 'Reset a rolled-back historical recovery receipt for a new attempt')) {
        throw 'Historical recovery receipt reset was declined.'
    }
    $Receipt.attempt = [int]$Receipt.attempt + 1
    $Receipt.state = 'Validated'
    $Receipt.error = ''
    $Receipt.ledger_sha256_after = ''
    $Receipt.inserted_response_hashes = @()
    $Receipt.published_file_count = 0

    $ledger = Read-GssFeedbackLedger $Paths.LedgerPath
    $ledgerByHash = @{}
    foreach ($entry in @($ledger.entries)) { $ledgerByHash[[string]$entry.response_hash] = $entry }
    $plannedEntries = Get-GssHistoricalRecoveryPlanEntry $Evidence
    $Receipt.planned_entries = $plannedEntries
    $Receipt.planned_insertions = @($plannedEntries | Where-Object { -not $ledgerByHash.ContainsKey([string]$_.response_hash) })
    $Receipt.ledger_existed_before = Test-Path -LiteralPath $Paths.LedgerPath -PathType Leaf
    $Receipt.ledger_sha256_before = if ($Receipt.ledger_existed_before) { Get-GssSha256 $Paths.LedgerPath } else { '' }
    $backupPath = Join-Path $Paths.TransactionRoot "ledger-original-attempt-$($Receipt.attempt).json"
    $Receipt.ledger_backup_path = if ($Receipt.ledger_existed_before) { $backupPath } else { '' }
    if ($Receipt.ledger_existed_before) {
        Copy-GssHistoricalRecoveryExactFile `
            -Source $Paths.LedgerPath `
            -Destination $backupPath `
            -ExpectedSha256 $Receipt.ledger_sha256_before `
            -ExpectedByteSize (Get-Item -LiteralPath $Paths.LedgerPath).Length
    }

    for ($index = 0; $index -lt @($Receipt.files).Count; $index++) {
        $file = $Receipt.files[$index]
        $sourceEvidence = $Evidence[[int]$file.source_index - 1]
        $file.source_leaf_name = Split-Path -Leaf $sourceEvidence.Path
        $file.response_hashes = @($sourceEvidence.ResponseEntries.response_hash | Sort-Object -Unique)
        $file.prepared = $false
        $file.destination_preexisting = Test-Path -LiteralPath $file.destination_full_path -PathType Leaf
        $file.published_by_transaction = $false
    }
    Write-GssHistoricalRecoveryReceipt -Receipt $Receipt -Path $Paths.ReceiptPath
    return $Receipt
}

function Enter-GssHistoricalRecoveryMutex {
    param([Parameter(Mandatory)][int]$TimeoutSeconds)

    if ($TimeoutSeconds -lt 0) { throw 'MutexTimeoutSeconds cannot be negative.' }
    $mutex = New-Object System.Threading.Mutex($false, $script:GssHistoricalRecoveryMutexName)
    $acquired = $false
    try {
        try {
            $acquired = $mutex.WaitOne([timespan]::FromSeconds($TimeoutSeconds))
        }
        catch [System.Threading.AbandonedMutexException] {
            $acquired = $true
        }
        if (-not $acquired) {
            throw "Timed out waiting for workstation transaction mutex '$($script:GssHistoricalRecoveryMutexName)'."
        }
        return [pscustomobject]@{ Mutex = $mutex; Acquired = $true }
    }
    catch {
        $mutex.Dispose()
        throw
    }
}

function Exit-GssHistoricalRecoveryMutex {
    param([object]$Handle)

    if ($Handle -and $Handle.Mutex) {
        try {
            if ($Handle.Acquired) { $Handle.Mutex.ReleaseMutex() }
        }
        finally {
            $Handle.Mutex.Dispose()
        }
    }
}

function Assert-GssHistoricalRecoveryLedgerCoverage {
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][string]$Ledger
    )

    $state = Read-GssFeedbackLedger $Ledger
    $byHash = @{}
    foreach ($entry in @($state.entries)) { $byHash[[string]$entry.response_hash] = $entry }
    foreach ($planned in @($Receipt.planned_entries)) {
        if (-not $byHash.ContainsKey([string]$planned.response_hash)) {
            throw "Historical recovery ledger is missing response hash '$($planned.response_hash)'."
        }
    }
    return $state
}

function Add-GssHistoricalRecoveryLedgerBaseline {
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][string]$Ledger
    )

    $state = Read-GssFeedbackLedger $Ledger
    $entries = @($state.entries)
    $byHash = @{}
    foreach ($entry in $entries) { $byHash[[string]$entry.response_hash] = $entry }
    foreach ($planned in @($Receipt.planned_entries | Sort-Object response_hash)) {
        $hash = [string]$planned.response_hash
        if ($byHash.ContainsKey($hash)) { continue }
        $newEntry = [pscustomobject][ordered]@{
            response_hash = $hash
            first_seen_reporting_date = [string]$planned.first_seen_reporting_date
            first_seen_package_id = [string]$Receipt.transaction_id
        }
        $entries += $newEntry
        $byHash[$hash] = $newEntry
    }
    $next = [pscustomobject][ordered]@{
        schema_version = $script:GssFeedbackLedgerVersion
        entries = @($entries | Sort-Object response_hash)
    }
    if (-not (Test-Path -LiteralPath $Ledger -PathType Leaf) -or @($next.entries).Count -ne @($state.entries).Count) {
        Write-GssFeedbackLedger -Ledger $next -Path $Ledger
    }

    $verified = Assert-GssHistoricalRecoveryLedgerCoverage -Receipt $Receipt -Ledger $Ledger
    $transactionHashes = @($verified.entries |
        Where-Object { [string]$_.first_seen_package_id -eq [string]$Receipt.transaction_id } |
        Select-Object -ExpandProperty response_hash |
        Sort-Object -Unique)
    $Receipt.inserted_response_hashes = @($Receipt.planned_insertions.response_hash |
        Where-Object { $_ -in $transactionHashes } |
        Sort-Object -Unique)
    $Receipt.ledger_sha256_after = Get-GssSha256 $Ledger
    $Receipt.state = 'LedgerBaselined'
}

function Test-GssHistoricalRecoveryExactArtifact {
    param(
        [Parameter(Mandatory)][object]$File,
        [Parameter(Mandatory)][object]$Source,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$GssRoot,
        [Parameter(Mandatory)][string]$Label
    )

    $evidence = Read-GssHistoricalRecoveryWorkbookEvidence -Path $Path -Source $Source -GssRoot $GssRoot -Label $Label
    $receiptHashes = @($File.response_hashes | ForEach-Object { [string]$_ } | Sort-Object -Unique)
    $actualHashes = @($evidence.ResponseEntries.response_hash | Sort-Object -Unique)
    if (($receiptHashes -join "`n") -ne ($actualHashes -join "`n")) {
        throw "$Label does not match the response hashes recorded in the transaction receipt."
    }
    return $evidence
}

function Assert-GssHistoricalRecoveryReceiptResponsePlan {
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][object[]]$Evidence
    )

    $actualEntries = Get-GssHistoricalRecoveryPlanEntry $Evidence
    $receiptEntries = @($Receipt.planned_entries | Sort-Object response_hash)
    if ($receiptEntries.Count -ne $actualEntries.Count) {
        throw 'Historical recovery receipt response plan count does not match the exact workbooks.'
    }
    for ($index = 0; $index -lt $actualEntries.Count; $index++) {
        if ([string]$receiptEntries[$index].response_hash -ne [string]$actualEntries[$index].response_hash -or
            [string]$receiptEntries[$index].first_seen_reporting_date -ne [string]$actualEntries[$index].first_seen_reporting_date) {
            throw 'Historical recovery receipt response plan does not match the exact workbook identities and visit dates.'
        }
    }
    $plannedHashes = @($actualEntries.response_hash)
    foreach ($hash in @($Receipt.planned_insertions.response_hash) + @($Receipt.inserted_response_hashes)) {
        if ([string]$hash -notin $plannedHashes) {
            throw "Historical recovery receipt contains an out-of-scope ledger hash '$hash'."
        }
    }
}

function Initialize-GssHistoricalRecoveryFile {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][object]$Manifest,
        [object[]]$Evidence,
        [Parameter(Mandatory)][string]$GssRoot,
        [Parameter(Mandatory)][string]$ReceiptPath
    )

    if (-not $PSCmdlet.ShouldProcess($ReceiptPath, 'Prepare exact historical workbooks without XLSX visibility')) {
        throw 'Historical recovery file preparation was declined.'
    }
    $evidenceByHash = @{}
    foreach ($item in @($Evidence)) { $evidenceByHash[[string]$item.Sha256] = $item }

    $validatedEvidence = @()
    foreach ($file in @($Receipt.files | Sort-Object source_index)) {
        $source = $Manifest.Sources[[int]$file.source_index - 1]
        $destination = [string]$file.destination_full_path
        $partial = [string]$file.partial_path
        if (Test-Path -LiteralPath $destination -PathType Leaf) {
            if (Test-Path -LiteralPath $partial -PathType Leaf) {
                throw "Historical recovery has both a final and partial file for source $($file.source_index); manual review is required."
            }
            $validatedEvidence += Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $destination -GssRoot $GssRoot -Label "Historical archive destination $($file.source_index)"
            if (-not [bool]$file.destination_preexisting -and -not (Test-Path -LiteralPath $partial -PathType Leaf)) {
                $file.published_by_transaction = $true
            }
            $file.prepared = $true
            Write-GssHistoricalRecoveryReceipt -Receipt $Receipt -Path $ReceiptPath
            continue
        }
        if (Test-Path -LiteralPath $partial -PathType Leaf) {
            $validatedEvidence += Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $partial -GssRoot $GssRoot -Label "Historical recovery partial $($file.source_index)"
            $file.prepared = $true
            Write-GssHistoricalRecoveryReceipt -Receipt $Receipt -Path $ReceiptPath
            continue
        }

        $sourceEvidence = $evidenceByHash[[string]$file.sha256]
        if ($null -eq $sourceEvidence) {
            throw "Historical recovery source $($file.source_index) must be supplied again because no verified partial or final file exists."
        }
        $destinationDirectory = Split-Path -Parent $destination
        if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
        }
        Copy-GssHistoricalRecoveryExactFile `
            -Source $sourceEvidence.Path `
            -Destination $partial `
            -ExpectedSha256 $file.sha256 `
            -ExpectedByteSize ([long]$file.byte_size)
        $validatedEvidence += Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $partial -GssRoot $GssRoot -Label "Historical recovery partial $($file.source_index)"
        $file.prepared = $true
        Write-GssHistoricalRecoveryReceipt -Receipt $Receipt -Path $ReceiptPath
    }
    Assert-GssHistoricalRecoveryReceiptResponsePlan -Receipt $Receipt -Evidence $validatedEvidence
    $Receipt.state = 'Prepared'
    Write-GssHistoricalRecoveryReceipt -Receipt $Receipt -Path $ReceiptPath
}

function Publish-GssHistoricalRecoveryFile {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][string]$GssRoot,
        [Parameter(Mandatory)][string]$ReceiptPath,
        [string]$FailurePoint
    )

    if (-not $PSCmdlet.ShouldProcess($ReceiptPath, 'Publish historical XLSX files after ledger baseline verification')) {
        throw 'Historical recovery publication was declined.'
    }
    [void](Assert-GssHistoricalRecoveryLedgerCoverage -Receipt $Receipt -Ledger $Receipt.ledger_path)
    $Receipt.state = 'Publishing'
    Write-GssHistoricalRecoveryReceipt -Receipt $Receipt -Path $ReceiptPath

    $publishedThisCall = 0
    foreach ($file in @($Receipt.files | Sort-Object source_index)) {
        $source = $Manifest.Sources[[int]$file.source_index - 1]
        $destination = [string]$file.destination_full_path
        $partial = [string]$file.partial_path
        if (Test-Path -LiteralPath $destination -PathType Leaf) {
            [void](Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $destination -GssRoot $GssRoot -Label "Historical archive destination $($file.source_index)")
            if (-not [bool]$file.destination_preexisting -and -not (Test-Path -LiteralPath $partial -PathType Leaf)) {
                $file.published_by_transaction = $true
            }
            continue
        }
        if (-not (Test-Path -LiteralPath $partial -PathType Leaf)) {
            throw "Verified historical recovery partial is missing for source $($file.source_index): $partial"
        }
        [System.IO.File]::Move($partial, $destination)
        [void](Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $destination -GssRoot $GssRoot -Label "Published historical archive $($file.source_index)")
        $file.published_by_transaction = $true
        $Receipt.published_file_count = @($Receipt.files | Where-Object { [bool]$_.published_by_transaction }).Count
        Write-GssHistoricalRecoveryReceipt -Receipt $Receipt -Path $ReceiptPath
        $publishedThisCall++
        if ($FailurePoint -eq 'AfterFirstPublish' -and $publishedThisCall -eq 1) {
            throw 'Injected historical recovery failure after the first archive publication.'
        }
    }

    [void](Assert-GssHistoricalRecoveryLedgerCoverage -Receipt $Receipt -Ledger $Receipt.ledger_path)
    $Receipt.published_file_count = @($Receipt.files | Where-Object { [bool]$_.published_by_transaction }).Count
    $Receipt.state = 'Committed'
    $Receipt.error = ''
    Write-GssHistoricalRecoveryReceipt -Receipt $Receipt -Path $ReceiptPath
}

function Assert-GssHistoricalRecoveryCommitted {
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][string]$GssRoot
    )

    [void](Assert-GssHistoricalRecoveryLedgerCoverage -Receipt $Receipt -Ledger $Receipt.ledger_path)
    foreach ($file in @($Receipt.files)) {
        $source = $Manifest.Sources[[int]$file.source_index - 1]
        [void](Test-GssHistoricalRecoveryExactArtifact `
            -File $file `
            -Source $source `
            -Path $file.destination_full_path `
            -GssRoot $GssRoot `
            -Label "Committed historical archive $($file.source_index)")
        if (Test-Path -LiteralPath $file.partial_path -PathType Leaf) {
            throw "Committed historical recovery retained an unexpected partial: $($file.partial_path)"
        }
    }
}

function Get-GssHistoricalRecoveryResult {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$RequestedOperation,
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][object]$Paths,
        [object]$Receipt,
        [object[]]$Evidence,
        [bool]$Idempotent = $false
    )

    $uniqueResponseCount = if ($Receipt) {
        @($Receipt.planned_entries).Count
    }
    elseif ($Evidence) {
        @(Get-GssHistoricalRecoveryPlanEntry $Evidence).Count
    }
    else { 0 }
    $rowCount = if ($Receipt) {
        (@($Receipt.files | Measure-Object -Property row_count -Sum).Sum)
    }
    elseif ($Evidence) {
        (@($Evidence | Measure-Object -Property RowCount -Sum).Sum)
    }
    else { 0 }

    $sources = @()
    for ($index = 0; $index -lt @($Manifest.Sources).Count; $index++) {
        $source = $Manifest.Sources[$index]
        $action = if (Test-Path -LiteralPath $source.DestinationFullPath -PathType Leaf) { 'ReuseExact' } else { 'Publish' }
        $sources += [pscustomobject][ordered]@{
            source_index = [int]$source.SourceIndex
            source_kind = [string]$source.SourceKind
            source_report_week = [string]$source.SourceReportWeek
            subject_week = [string]$source.SubjectWeek
            assignment_basis = [string]$source.AssignmentBasis
            row_count = [int]$source.RowCount
            byte_size = [long]$source.ByteSize
            sha256 = [string]$source.Sha256
            response_set_sha256 = [string]$source.ResponseSetSha256
            destination_path = [string]$source.DestinationPath
            action = $action
        }
    }

    return [pscustomobject][ordered]@{
        Status = $Status
        Operation = $RequestedOperation
        ManifestSha256 = [string]$Manifest.Sha256
        TransactionId = "historical-recovery:$($Manifest.Sha256)"
        ReceiptPath = if ($Receipt) { [string]$Paths.ReceiptPath } else { '' }
        LedgerPath = [string]$Paths.LedgerPath
        SourceCount = @($Manifest.Sources).Count
        RowCount = [int]$rowCount
        UniqueResponseCount = [int]$uniqueResponseCount
        LedgerEntriesAdded = if ($Receipt) { @($Receipt.inserted_response_hashes).Count } else { 0 }
        PublishedFileCount = if ($Receipt) { [int]$Receipt.published_file_count } else { 0 }
        Idempotent = [bool]$Idempotent
        Sources = $sources
        Controls = [pscustomobject][ordered]@{
            mutex = $script:GssHistoricalRecoveryMutexName
            ledger_before_archive_visibility = $true
            live_workbook_mutated = $false
            email_package_mutated = $false
            scheduled_task_mutated = $false
            automatic_sending_enabled = $false
        }
    }
}

function Invoke-GssHistoricalRecoveryPlan {
    param(
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][object]$Paths,
        [Parameter(Mandatory)][object[]]$Evidence,
        [Parameter(Mandatory)][string]$GssRoot
    )

    foreach ($source in $Manifest.Sources) {
        if (Test-Path -LiteralPath $source.DestinationFullPath -PathType Leaf) {
            [void](Read-GssHistoricalRecoveryWorkbookEvidence `
                -Path $source.DestinationFullPath `
                -Source $source `
                -GssRoot $GssRoot `
                -Label "Existing historical archive destination $($source.SourceIndex)")
        }
    }
    return Get-GssHistoricalRecoveryResult `
        -Status 'Planned' `
        -RequestedOperation 'Plan' `
        -Manifest $Manifest `
        -Paths $Paths `
        -Evidence $Evidence
}

function Invoke-GssHistoricalRecoveryApply {
    param(
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][object]$Paths,
        [object[]]$Evidence,
        [Parameter(Mandatory)][string]$GssRoot,
        [string]$FailurePoint
    )

    $receipt = $null
    try {
        if (Test-Path -LiteralPath $Paths.ReceiptPath -PathType Leaf) {
            $receipt = Get-Content -Raw -LiteralPath $Paths.ReceiptPath | ConvertFrom-Json
            Assert-GssHistoricalRecoveryReceipt -Receipt $receipt -Manifest $Manifest -Paths $Paths
            if ([string]$receipt.state -eq 'Committed') {
                Assert-GssHistoricalRecoveryCommitted -Receipt $receipt -Manifest $Manifest -GssRoot $GssRoot
                return Get-GssHistoricalRecoveryResult `
                    -Status 'Committed' `
                    -RequestedOperation 'Apply' `
                    -Manifest $Manifest `
                    -Paths $Paths `
                    -Receipt $receipt `
                    -Idempotent $true
            }
            if ([string]$receipt.state -in @('RolledBack', 'RolledBackConservative')) {
                if (@($Evidence).Count -ne @($Manifest.Sources).Count) {
                    throw 'A rolled-back recovery requires the exact staged sources for a new Apply attempt.'
                }
                $receipt = Reset-GssHistoricalRecoveryReceipt `
                    -Receipt $receipt `
                    -Evidence $Evidence `
                    -Paths $Paths
            }
        }
        else {
            if (@($Evidence).Count -ne @($Manifest.Sources).Count) {
                throw 'A fresh historical recovery Apply requires every exact staged source.'
            }
            $receipt = New-GssHistoricalRecoveryReceipt -Manifest $Manifest -Evidence $Evidence -Paths $Paths
            Write-GssHistoricalRecoveryReceipt -Receipt $receipt -Path $Paths.ReceiptPath
        }

        Initialize-GssHistoricalRecoveryFile `
            -Receipt $receipt `
            -Manifest $Manifest `
            -Evidence $Evidence `
            -GssRoot $GssRoot `
            -ReceiptPath $Paths.ReceiptPath

        Add-GssHistoricalRecoveryLedgerBaseline -Receipt $receipt -Ledger $Paths.LedgerPath
        Write-GssHistoricalRecoveryReceipt -Receipt $receipt -Path $Paths.ReceiptPath
        if ($FailurePoint -eq 'AfterLedgerBaseline') {
            throw 'Injected historical recovery failure after ledger baseline.'
        }

        Publish-GssHistoricalRecoveryFile `
            -Receipt $receipt `
            -Manifest $Manifest `
            -GssRoot $GssRoot `
            -ReceiptPath $Paths.ReceiptPath `
            -FailurePoint $FailurePoint

        return Get-GssHistoricalRecoveryResult `
            -Status 'Committed' `
            -RequestedOperation 'Apply' `
            -Manifest $Manifest `
            -Paths $Paths `
            -Receipt $receipt
    }
    catch {
        $caught = $_
        if ($receipt) {
            try {
                if ([string]$receipt.state -ne 'Committed') {
                    $receipt.state = 'NeedsResume'
                    $receipt.error = $caught.Exception.Message
                    Write-GssHistoricalRecoveryReceipt -Receipt $receipt -Path $Paths.ReceiptPath
                }
            }
            catch {
                $caught.Exception.Data['GssHistoricalRecoveryReceiptError'] = $_.Exception.Message
            }
        }
        throw $caught
    }
}

function Invoke-GssHistoricalRecoveryRollback {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][object]$Paths,
        [Parameter(Mandatory)][string]$GssRoot,
        [string]$FailurePoint
    )

    if (-not $PSCmdlet.ShouldProcess($Paths.ReceiptPath, 'Roll back transaction-published historical workbooks and their ledger baselines')) {
        throw 'Historical recovery rollback was declined.'
    }
    if (-not (Test-Path -LiteralPath $Paths.ReceiptPath -PathType Leaf)) {
        throw "Historical recovery receipt is missing; there is no transaction to roll back: $($Paths.ReceiptPath)"
    }
    $receipt = Get-Content -Raw -LiteralPath $Paths.ReceiptPath | ConvertFrom-Json
    Assert-GssHistoricalRecoveryReceipt -Receipt $receipt -Manifest $Manifest -Paths $Paths
    $wasRolledBack = [string]$receipt.state -in @('RolledBack', 'RolledBackConservative')
    $rollbackArtifacts = @()
    $moveToRollback = @()
    $rollbackChanged = $false
    foreach ($file in @($receipt.files | Sort-Object source_index)) {
        $source = $Manifest.Sources[[int]$file.source_index - 1]
        $destination = [string]$file.destination_full_path
        $partial = [string]$file.partial_path
        $rollbackPath = Join-Path (Split-Path -Parent $destination) ".$($file.sha256).rollback-part"
        $destinationExists = Test-Path -LiteralPath $destination -PathType Leaf
        $rollbackExists = Test-Path -LiteralPath $rollbackPath -PathType Leaf

        if ([bool]$file.destination_preexisting) {
            if ($rollbackExists) {
                throw "A rollback partial exists for a destination that predated this transaction: $rollbackPath"
            }
            continue
        }
        if ($destinationExists -and $rollbackExists) {
            [void](Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $destination -GssRoot $GssRoot -Label "Rollback historical archive $($file.source_index)")
            [void](Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $rollbackPath -GssRoot $GssRoot -Label "Existing rollback partial $($file.source_index)")
            throw "Historical recovery found both the final and rollback partial for source $($file.source_index); manual review is required."
        }
        if ($destinationExists) {
            [void](Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $destination -GssRoot $GssRoot -Label "Rollback historical archive $($file.source_index)")
            $file.published_by_transaction = $true
            $moveToRollback += [pscustomobject]@{
                File = $file
                Destination = $destination
                RollbackPath = $rollbackPath
            }
        }
        elseif ($rollbackExists) {
            [void](Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $rollbackPath -GssRoot $GssRoot -Label "Existing rollback partial $($file.source_index)")
            $file.published_by_transaction = $true
        }
        $rollbackArtifacts += [pscustomobject]@{
            File = $file
            RollbackPath = $rollbackPath
        }

        if (Test-Path -LiteralPath $partial -PathType Leaf) {
            [void](Test-GssHistoricalRecoveryExactArtifact -File $file -Source $source -Path $partial -GssRoot $GssRoot -Label "Rollback prepared partial $($file.source_index)")
        }
    }

    $hiddenThisCall = 0
    foreach ($item in $moveToRollback) {
        [System.IO.File]::Move($item.Destination, $item.RollbackPath)
        $rollbackChanged = $true
        $hiddenThisCall++
        $receipt.state = 'RollingBack'
        $receipt.error = ''
        Write-GssHistoricalRecoveryReceipt -Receipt $receipt -Path $Paths.ReceiptPath
        if ($FailurePoint -eq 'AfterFirstRollbackHide' -and $hiddenThisCall -eq 1) {
            $receipt.state = 'NeedsRollback'
            $receipt.error = 'Injected historical recovery failure after the first rollback hide.'
            Write-GssHistoricalRecoveryReceipt -Receipt $receipt -Path $Paths.ReceiptPath
            throw $receipt.error
        }
    }
    $receipt.state = 'RollbackHidden'
    $receipt.error = ''
    Write-GssHistoricalRecoveryReceipt -Receipt $receipt -Path $Paths.ReceiptPath

    $protectedHashes = @($receipt.files |
        Where-Object { [bool]$_.destination_preexisting } |
        ForEach-Object { @($_.response_hashes) } |
        Sort-Object -Unique)
    $insertedHashes = @($receipt.inserted_response_hashes | Sort-Object -Unique)
    $removableHashes = @($insertedHashes | Where-Object { $_ -notin $protectedHashes })
    $ledger = Read-GssFeedbackLedger $Paths.LedgerPath
    $remaining = @($ledger.entries | Where-Object {
        $hash = [string]$_.response_hash
        -not ($hash -in $removableHashes -and [string]$_.first_seen_package_id -eq [string]$receipt.transaction_id)
    })
    if ($remaining.Count -ne @($ledger.entries).Count) {
        $next = [pscustomobject][ordered]@{
            schema_version = $script:GssFeedbackLedgerVersion
            entries = @($remaining | Sort-Object response_hash)
        }
        Write-GssFeedbackLedger -Ledger $next -Path $Paths.LedgerPath
        $rollbackChanged = $true
    }

    foreach ($item in $rollbackArtifacts) {
        if (Test-Path -LiteralPath $item.RollbackPath -PathType Leaf) {
            if ((Get-GssSha256 $item.RollbackPath) -ne [string]$item.File.sha256) {
                throw "Historical recovery rollback partial changed and was retained: $($item.RollbackPath)"
            }
            Remove-Item -LiteralPath $item.RollbackPath -Force
            $rollbackChanged = $true
        }
        $item.File.published_by_transaction = $false
    }
    foreach ($file in @($receipt.files)) {
        if (Test-Path -LiteralPath $file.partial_path -PathType Leaf) {
            if ((Get-GssSha256 $file.partial_path) -ne [string]$file.sha256) {
                throw "Historical recovery partial changed and was retained: $($file.partial_path)"
            }
            Remove-Item -LiteralPath $file.partial_path -Force
            $rollbackChanged = $true
        }
        $file.prepared = $false
        if (-not [bool]$file.destination_preexisting) {
            $destination = [string]$file.destination_full_path
            $rollbackPath = Join-Path (Split-Path -Parent $destination) ".$($file.sha256).rollback-part"
            if ((Test-Path -LiteralPath $destination -PathType Leaf) -or
                (Test-Path -LiteralPath $rollbackPath -PathType Leaf) -or
                (Test-Path -LiteralPath $file.partial_path -PathType Leaf)) {
                throw "Historical recovery rollback could not prove source $($file.source_index) invisible."
            }
        }
    }

    $receipt.published_file_count = 0
    $receipt.ledger_sha256_after = if (Test-Path -LiteralPath $Paths.LedgerPath -PathType Leaf) { Get-GssSha256 $Paths.LedgerPath } else { '' }
    $receipt.state = if (@($insertedHashes | Where-Object { $_ -in $protectedHashes }).Count -gt 0) {
        'RolledBackConservative'
    }
    else { 'RolledBack' }
    $receipt.error = ''
    Write-GssHistoricalRecoveryReceipt -Receipt $receipt -Path $Paths.ReceiptPath

    return Get-GssHistoricalRecoveryResult `
        -Status ([string]$receipt.state) `
        -RequestedOperation 'Rollback' `
        -Manifest $Manifest `
        -Paths $Paths `
        -Receipt $receipt `
        -Idempotent ([bool]($wasRolledBack -and -not $rollbackChanged))
}

$resolvedFolderPath = [System.IO.Path]::GetFullPath($FolderPath).TrimEnd('\', '/')
if (-not (Test-Path -LiteralPath $resolvedFolderPath -PathType Container)) {
    throw "GSS root folder is missing: $resolvedFolderPath"
}
$manifest = Read-GssHistoricalRecoveryManifest -Path $ManifestPath -GssRoot $resolvedFolderPath
$paths = Get-GssHistoricalRecoveryPath `
    -GssRoot $resolvedFolderPath `
    -ManifestSha256 $manifest.Sha256 `
    -RequestedLedgerPath $LedgerPath

$evidence = @()
if (@($SourcePath).Count -gt 0) {
    $evidence = @(Resolve-GssHistoricalRecoverySourceEvidence `
        -Manifest $manifest `
        -Paths $SourcePath `
        -GssRoot $resolvedFolderPath `
        -RuntimeRoot $paths.RuntimeRoot)
}
elseif ($Operation -eq 'Plan') {
    throw 'Historical recovery Plan requires every exact staged source path.'
}

$result = $null
if ($Operation -eq 'Plan') {
    $result = Invoke-GssHistoricalRecoveryPlan `
        -Manifest $manifest `
        -Paths $paths `
        -Evidence $evidence `
        -GssRoot $resolvedFolderPath
}
else {
    $mutexHandle = $null
    try {
        $mutexHandle = Enter-GssHistoricalRecoveryMutex -TimeoutSeconds $MutexTimeoutSeconds
        if ($Operation -eq 'Apply') {
            $result = Invoke-GssHistoricalRecoveryApply `
                -Manifest $manifest `
                -Paths $paths `
                -Evidence $evidence `
                -GssRoot $resolvedFolderPath `
                -FailurePoint $TestFailurePoint
        }
        else {
            $result = Invoke-GssHistoricalRecoveryRollback `
                -Manifest $manifest `
                -Paths $paths `
                -GssRoot $resolvedFolderPath `
                -FailurePoint $TestFailurePoint
        }
    }
    finally {
        Exit-GssHistoricalRecoveryMutex $mutexHandle
    }
}

if ($OutputObject) {
    return $result
}
$result | ConvertTo-Json -Depth 10
