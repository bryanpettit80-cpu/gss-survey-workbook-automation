[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$testScriptRoot = Split-Path -Parent $PSCommandPath
$importScript = Join-Path $testScriptRoot 'Import-GSS-HistoricalGuestDetail.ps1'
. (Join-Path $testScriptRoot 'Gss-EmailPackage.ps1')

function Assert-HistoricalRecoveryTrue {
    param([bool]$Value, [string]$Name)
    if (-not $Value) { throw "Assertion failed for ${Name}." }
}

function Assert-HistoricalRecoveryEqual {
    param([object]$Actual, [object]$Expected, [string]$Name)
    if ($Actual -ne $Expected) {
        throw "Assertion failed for $Name. Expected '$Expected'; actual '$Actual'."
    }
}

function Assert-HistoricalRecoveryThrowsLike {
    param([scriptblock]$Script, [string]$Pattern, [string]$Name)
    try {
        & $Script
        throw "Assertion failed for $Name. Expected an exception."
    }
    catch {
        if ($_.Exception.Message -notlike $Pattern) {
            throw "Assertion failed for $Name. Expected '$Pattern'; actual '$($_.Exception.Message)'."
        }
    }
}

function ConvertTo-HistoricalRecoveryTestExcelColumn {
    param([int]$Number)

    $value = $Number
    $letters = ''
    while ($value -gt 0) {
        $value--
        $letters = [char](65 + ($value % 26)) + $letters
        $value = [math]::Floor($value / 26)
    }
    return $letters
}

function Add-HistoricalRecoveryTestZipText {
    param([object]$Archive, [string]$Name, [string]$Text)

    $entry = $Archive.CreateEntry($Name)
    $stream = $entry.Open()
    try {
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        try { $writer.Write($Text) }
        finally { $writer.Dispose() }
    }
    finally {
        $stream.Dispose()
    }
}

function New-HistoricalRecoveryTestWorkbook {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$Headers,
        [Parameter(Mandatory)][object[]]$Records
    )

    if (-not $PSCmdlet.ShouldProcess($Path, 'Create synthetic historical recovery workbook')) {
        return
    }
    Add-Type -AssemblyName System.IO.Compression
    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    try {
        $archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            Add-HistoricalRecoveryTestZipText `
                -Archive $archive `
                -Name 'xl/workbook.xml' `
                -Text '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sorensen" sheetId="1" r:id="rId1"/></sheets></workbook>'
            Add-HistoricalRecoveryTestZipText `
                -Archive $archive `
                -Name 'xl/_rels/workbook.xml.rels' `
                -Text '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'

            $rowXml = @()
            $allRows = @([pscustomobject]@{ IsHeader = $true; Value = $null }) +
                @($Records | ForEach-Object { [pscustomobject]@{ IsHeader = $false; Value = $_ } })
            for ($rowIndex = 0; $rowIndex -lt $allRows.Count; $rowIndex++) {
                $excelRow = $rowIndex + 1
                $cells = @()
                for ($columnIndex = 0; $columnIndex -lt $Headers.Count; $columnIndex++) {
                    $header = $Headers[$columnIndex]
                    $value = if ($allRows[$rowIndex].IsHeader) {
                        $header
                    }
                    else {
                        $key = Normalize-GssFeedbackHeader $header
                        $property = $allRows[$rowIndex].Value.PSObject.Properties[$key]
                        if ($property) { [string]$property.Value } else { '' }
                    }
                    if ([string]::IsNullOrEmpty($value) -and -not $allRows[$rowIndex].IsHeader) { continue }
                    $reference = (ConvertTo-HistoricalRecoveryTestExcelColumn ($columnIndex + 1)) + $excelRow
                    $escaped = [System.Security.SecurityElement]::Escape([string]$value)
                    $cells += "<c r=`"$reference`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$escaped</t></is></c>"
                }
                $rowXml += "<row r=`"$excelRow`">$($cells -join '')</row>"
            }
            $lastColumn = ConvertTo-HistoricalRecoveryTestExcelColumn $Headers.Count
            $worksheet = '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                "<dimension ref=`"A1:$lastColumn$($allRows.Count)`"/><sheetData>$($rowXml -join '')</sheetData></worksheet>"
            Add-HistoricalRecoveryTestZipText -Archive $archive -Name 'xl/worksheets/sheet1.xml' -Text $worksheet
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Get-HistoricalRecoveryTestResponse {
    param(
        [Parameter(Mandatory)][string]$Restaurant,
        [Parameter(Mandatory)][string]$Date,
        [Parameter(Mandatory)][string]$Time,
        [Parameter(Mandatory)][string]$Text
    )

    return [pscustomobject]@{
        restaurantname = $Restaurant
        reservationdate = $Date
        reservationtime = $Time
        text = $Text
        overall = '5'
        service = '5'
        culinary = '5'
        value = '5'
        paceofmeal = '5'
        recommend = '10'
        managervisit = 'Y'
        steakcookedcorrectly = 'Y'
        alertguestsdonotcontact = ''
        eventbookingprocess = ''
        firstvisit = 'N'
        guestfirstname = 'Synthetic'
        guestlastname = 'Fixture'
    }
}

function Get-HistoricalRecoveryTestResponseSetSha256 {
    param([Parameter(Mandatory)][string[]]$ResponseHashes)

    $hashes = @($ResponseHashes | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $material = "gss-historical-response-set/v1`n" + ($hashes -join "`n")
    return Get-GssStringSha256 $material
}

function Write-HistoricalRecoveryTestJson {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][object]$Value)

    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $encoding = New-Object System.Text.UTF8Encoding($false, $true)
    [System.IO.File]::WriteAllText(
        $Path,
        ($Value | ConvertTo-Json -Depth 20) + [Environment]::NewLine,
        $encoding
    )
}

function New-HistoricalRecoveryTestFixture {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$Name
    )

    $gssRoot = Join-Path $BasePath $Name
    if (-not $PSCmdlet.ShouldProcess($gssRoot, 'Create isolated historical recovery fixture')) {
        return $null
    }
    $staging = Join-Path $gssRoot '_automation_runs\staging'
    New-Item -ItemType Directory -Path $staging -Force | Out-Null

    $headers = @(
        'Restaurant Name',
        'Reservation Date',
        'Reservation Time',
        'Text',
        'Overall',
        'Service',
        'Culinary',
        'Value',
        'Pace of Meal',
        'Recommend',
        'Manager Visit',
        'Steak Cooked Correctly',
        'Alert Guests Do Not Contact',
        'Event Booking Process',
        'First Visit',
        'Guest First Name',
        'Guest Last Name'
    )
    $first = Get-HistoricalRecoveryTestResponse `
        -Restaurant '9354 Synthetic Store' `
        -Date '09/06/2025' `
        -Time '6:01 PM' `
        -Text "$Name first synthetic response"
    $overlap = Get-HistoricalRecoveryTestResponse `
        -Restaurant '9355 Synthetic Store' `
        -Date '09/13/2025' `
        -Time '7:02 PM' `
        -Text "$Name overlap synthetic response"
    $third = Get-HistoricalRecoveryTestResponse `
        -Restaurant '9354 Synthetic Store' `
        -Date '09/20/2025' `
        -Time '8:03 PM' `
        -Text "$Name third synthetic response"

    $sourceOne = Join-Path $staging 'source-one.xlsx'
    $sourceTwo = Join-Path $staging 'source-two.xlsx'
    New-HistoricalRecoveryTestWorkbook -Path $sourceOne -Headers $headers -Records @($first, $overlap)
    New-HistoricalRecoveryTestWorkbook -Path $sourceTwo -Headers $headers -Records @($overlap, $third)

    $sourcePaths = @($sourceOne, $sourceTwo)
    $reportWeeks = @('FY26 FW16', 'FY26 FW17')
    $subjectWeeks = @('FY26 FW14', 'FY26 FW15')
    $manifestSources = @()
    for ($index = 0; $index -lt $sourcePaths.Count; $index++) {
        $sourcePath = $sourcePaths[$index]
        $detail = Read-GssDetailWorkbook -Path $sourcePath -FolderPath $gssRoot
        $hash = Get-GssSha256 $sourcePath
        $responseHashes = @($detail.Responses.ResponseHash | Sort-Object -Unique)
        $manifestSources += [pscustomobject][ordered]@{
            source_kind = 'google_drive_and_gmail'
            drive_file_id = "drivefixture$($index + 1)abcdefghij"
            gmail_message_ids = @("199000000000000$($index + 1)")
            subject_week = $subjectWeeks[$index]
            source_report_week = $reportWeeks[$index]
            assignment_basis = 'paired_rolling_terminal_week'
            visit_date_start = $detail.VisitDateStart.ToString('yyyy-MM-dd')
            visit_date_end = $detail.VisitDateEnd.ToString('yyyy-MM-dd')
            row_count = @($detail.Responses).Count
            byte_size = [long](Get-Item -LiteralPath $sourcePath).Length
            sha256 = $hash
            response_set_sha256 = Get-HistoricalRecoveryTestResponseSetSha256 $responseHashes
            destination_path = "03 Uploaded Survey Workbooks/Archive - Previous Uploads/Recovered Historical Detail/FY26/$($reportWeeks[$index].Replace(' ', '-'))-$hash.xlsx"
            validation = [pscustomobject][ordered]@{
                detail_schema_valid = $true
                response_identity_version = 'gss-feedback-response-identity/v1'
                duplicate_response_count = @($detail.Responses).Count - $responseHashes.Count
            }
        }
    }

    $manifest = [pscustomobject][ordered]@{
        schema_version = 'gss-historical-recovery/v1'
        fiscal_year = 'FY26'
        created_at_utc = '2026-07-23T12:00:00Z'
        sources = $manifestSources
    }
    $manifestPath = Join-Path $staging 'recovery-manifest.json'
    Write-HistoricalRecoveryTestJson -Path $manifestPath -Value $manifest

    return [pscustomobject]@{
        GssRoot = $gssRoot
        Staging = $staging
        SourcePaths = $sourcePaths
        Manifest = $manifest
        ManifestPath = $manifestPath
        ManifestSha256 = Get-GssSha256 $manifestPath
        LedgerPath = Join-Path $gssRoot '_automation_runs\state\gss_feedback_first_seen.json'
        ReceiptPath = Join-Path $gssRoot "_automation_runs\historical-recovery\$(Get-GssSha256 $manifestPath)\transaction-receipt.json"
        ArchiveRoot = Join-Path $gssRoot '03 Uploaded Survey Workbooks\Archive - Previous Uploads\Recovered Historical Detail\FY26'
    }
}

$tokens = $null
$parseErrors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
    $importScript,
    [ref]$tokens,
    [ref]$parseErrors
)
Assert-HistoricalRecoveryEqual @($parseErrors).Count 0 'Historical recovery importer parses cleanly'
$sourceText = Get-Content -Raw -LiteralPath $importScript
Assert-HistoricalRecoveryTrue ($sourceText.Contains('Global\GSSSurveyWorkbookAutomationTransaction')) 'Importer owns the workstation transaction mutex'
Assert-HistoricalRecoveryTrue ($sourceText.Contains('LedgerBaselined')) 'Importer records the ledger-before-publication phase'
foreach ($forbidden in @(
    'GSS Score Trends - Main.xlsx',
    'New-Object -ComObject Excel.Application',
    'New-GssEmailPackage',
    'Send-MailMessage',
    'Set-ScheduledTask',
    'Enable-ScheduledTask'
)) {
    Assert-HistoricalRecoveryTrue (-not $sourceText.Contains($forbidden)) "Importer excludes unrelated live mutation token '$forbidden'"
}

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) "gss-hr-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
try {
    $basic = New-HistoricalRecoveryTestFixture -BasePath $testRoot -Name 'basic'
    $sourceHashesBefore = @($basic.SourcePaths | ForEach-Object { Get-GssSha256 $_ })

    $plan = & $importScript `
        -Operation Plan `
        -FolderPath $basic.GssRoot `
        -ManifestPath $basic.ManifestPath `
        -SourcePath $basic.SourcePaths `
        -OutputObject
    Assert-HistoricalRecoveryEqual $plan.Status 'Planned' 'Plan status'
    Assert-HistoricalRecoveryEqual $plan.SourceCount 2 'Plan source count'
    Assert-HistoricalRecoveryEqual $plan.RowCount 4 'Plan exact row count'
    Assert-HistoricalRecoveryEqual $plan.UniqueResponseCount 3 'Plan deduplicated response count'
    Assert-HistoricalRecoveryTrue (-not (Test-Path -LiteralPath $basic.LedgerPath)) 'Plan does not create the ledger'
    Assert-HistoricalRecoveryTrue (-not (Test-Path -LiteralPath $basic.ArchiveRoot)) 'Plan does not create the archive'
    Assert-HistoricalRecoveryTrue (-not (Test-Path -LiteralPath $basic.ReceiptPath)) 'Plan does not create a receipt'

    $badRowManifest = $basic.Manifest | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $badRowManifest.sources[0].row_count = [int]$badRowManifest.sources[0].row_count + 1
    $badRowManifestPath = Join-Path $basic.Staging 'bad-row-manifest.json'
    Write-HistoricalRecoveryTestJson -Path $badRowManifestPath -Value $badRowManifest
    Assert-HistoricalRecoveryThrowsLike {
        & $importScript `
            -Operation Plan `
            -FolderPath $basic.GssRoot `
            -ManifestPath $badRowManifestPath `
            -SourcePath $basic.SourcePaths `
            -OutputObject
    } '*row count mismatch*' 'Plan rejects an inexact manifest row count'

    $piiFieldManifest = $basic.Manifest | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $piiFieldManifest.sources[0] | Add-Member -NotePropertyName guest_comment -NotePropertyValue 'forbidden'
    $piiFieldManifestPath = Join-Path $basic.Staging 'pii-field-manifest.json'
    Write-HistoricalRecoveryTestJson -Path $piiFieldManifestPath -Value $piiFieldManifest
    Assert-HistoricalRecoveryThrowsLike {
        & $importScript `
            -Operation Plan `
            -FolderPath $basic.GssRoot `
            -ManifestPath $piiFieldManifestPath `
            -SourcePath $basic.SourcePaths `
            -OutputObject
    } '*unsupported property*PII-free*' 'Manifest schema rejects a free-form PII field'

    $conflictingWeekManifest = $basic.Manifest | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $conflictingWeekManifest.sources[1].source_report_week = $conflictingWeekManifest.sources[0].source_report_week
    $conflictingWeekManifest.sources[1].destination_path = "03 Uploaded Survey Workbooks/Archive - Previous Uploads/Recovered Historical Detail/FY26/FY26-FW16-$($conflictingWeekManifest.sources[1].sha256).xlsx"
    $conflictingWeekManifestPath = Join-Path $basic.Staging 'conflicting-week-manifest.json'
    Write-HistoricalRecoveryTestJson -Path $conflictingWeekManifestPath -Value $conflictingWeekManifest
    Assert-HistoricalRecoveryThrowsLike {
        & $importScript `
            -Operation Plan `
            -FolderPath $basic.GssRoot `
            -ManifestPath $conflictingWeekManifestPath `
            -SourcePath $basic.SourcePaths `
            -OutputObject
    } '*conflicting response sets for source_report_week*' 'Manifest refuses two different response sets assigned to one report week'

    $apply = & $importScript `
        -Operation Apply `
        -FolderPath $basic.GssRoot `
        -ManifestPath $basic.ManifestPath `
        -SourcePath $basic.SourcePaths `
        -OutputObject
    Assert-HistoricalRecoveryEqual $apply.Status 'Committed' 'Apply status'
    Assert-HistoricalRecoveryEqual $apply.LedgerEntriesAdded 3 'Apply ledger insertion count'
    Assert-HistoricalRecoveryEqual $apply.PublishedFileCount 2 'Apply publication count'
    Assert-HistoricalRecoveryTrue (-not $apply.Controls.live_workbook_mutated) 'Apply leaves the live workbook unchanged'
    Assert-HistoricalRecoveryTrue (-not $apply.Controls.email_package_mutated) 'Apply leaves email packages unchanged'
    Assert-HistoricalRecoveryTrue (-not $apply.Controls.scheduled_task_mutated) 'Apply leaves scheduled tasks unchanged'
    $ledger = Get-Content -Raw -LiteralPath $basic.LedgerPath | ConvertFrom-Json
    Assert-HistoricalRecoveryEqual @($ledger.entries).Count 3 'Ledger contains the unique recovered responses'
    Assert-HistoricalRecoveryEqual @($ledger.entries | Where-Object {
        [string]$_.first_seen_package_id -eq "historical-recovery:$($basic.ManifestSha256)"
    }).Count 3 'Ledger entries bind to the reviewed manifest fingerprint'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $basic.ArchiveRoot -File -Filter '*.xlsx').Count 2 'Apply publishes exact source files'
    Assert-HistoricalRecoveryTrue (@((Get-Content -Raw -LiteralPath $basic.ReceiptPath | ConvertFrom-Json).state) -contains 'Committed') 'Receipt reaches Committed'
    Assert-HistoricalRecoveryEqual (@($basic.SourcePaths | ForEach-Object { Get-GssSha256 $_ }) -join '|') ($sourceHashesBefore -join '|') 'Apply leaves staged source bytes unchanged'

    $repeat = & $importScript `
        -Operation Apply `
        -FolderPath $basic.GssRoot `
        -ManifestPath $basic.ManifestPath `
        -OutputObject
    Assert-HistoricalRecoveryTrue $repeat.Idempotent 'Committed Apply rerun is idempotent'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $basic.ArchiveRoot -File -Filter '*.xlsx').Count 2 'Idempotent rerun creates no duplicate archive'

    $resume = New-HistoricalRecoveryTestFixture -BasePath $testRoot -Name 'resume'
    Assert-HistoricalRecoveryThrowsLike {
        & $importScript `
            -Operation Apply `
            -FolderPath $resume.GssRoot `
            -ManifestPath $resume.ManifestPath `
            -SourcePath $resume.SourcePaths `
            -TestFailurePoint AfterLedgerBaseline `
            -OutputObject
    } '*Injected historical recovery failure after ledger baseline*' 'Injected post-ledger failure'
    Assert-HistoricalRecoveryEqual @((Get-Content -Raw -LiteralPath $resume.LedgerPath | ConvertFrom-Json).entries).Count 3 'Ledger is conservative after an interrupted Apply'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $resume.ArchiveRoot -File -Filter '*.xlsx' -ErrorAction SilentlyContinue).Count 0 'No XLSX is visible at the post-ledger failure point'
    Assert-HistoricalRecoveryEqual (Get-Content -Raw -LiteralPath $resume.ReceiptPath | ConvertFrom-Json).state 'NeedsResume' 'Interrupted receipt requests resume'
    $resumed = & $importScript `
        -Operation Apply `
        -FolderPath $resume.GssRoot `
        -ManifestPath $resume.ManifestPath `
        -OutputObject
    Assert-HistoricalRecoveryEqual $resumed.Status 'Committed' 'Apply resumes from verified partials without source paths'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $resume.ArchiveRoot -File -Filter '*.xlsx').Count 2 'Resume publishes all files'

    $rollback = New-HistoricalRecoveryTestFixture -BasePath $testRoot -Name 'rollback'
    Assert-HistoricalRecoveryThrowsLike {
        & $importScript `
            -Operation Apply `
            -FolderPath $rollback.GssRoot `
            -ManifestPath $rollback.ManifestPath `
            -SourcePath $rollback.SourcePaths `
            -TestFailurePoint AfterFirstPublish `
            -OutputObject
    } '*Injected historical recovery failure after the first archive publication*' 'Injected partial-publication failure'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $rollback.ArchiveRoot -File -Filter '*.xlsx').Count 1 'Interrupted publication leaves only one visible exact file'
    Assert-HistoricalRecoveryEqual @((Get-Content -Raw -LiteralPath $rollback.LedgerPath | ConvertFrom-Json).entries).Count 3 'Interrupted publication retains the safe baseline'
    $rolledBack = & $importScript `
        -Operation Rollback `
        -FolderPath $rollback.GssRoot `
        -ManifestPath $rollback.ManifestPath `
        -OutputObject
    Assert-HistoricalRecoveryEqual $rolledBack.Status 'RolledBack' 'Rollback status'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $rollback.ArchiveRoot -File -Filter '*.xlsx' -ErrorAction SilentlyContinue).Count 0 'Rollback hides transaction-published workbooks before ledger removal'
    Assert-HistoricalRecoveryEqual @((Get-Content -Raw -LiteralPath $rollback.LedgerPath | ConvertFrom-Json).entries).Count 0 'Rollback removes only transaction-inserted ledger entries'
    $reapplied = & $importScript `
        -Operation Apply `
        -FolderPath $rollback.GssRoot `
        -ManifestPath $rollback.ManifestPath `
        -SourcePath $rollback.SourcePaths `
        -OutputObject
    Assert-HistoricalRecoveryEqual $reapplied.Status 'Committed' 'Apply can restart after rollback'

    $rollbackResume = New-HistoricalRecoveryTestFixture -BasePath $testRoot -Name 'rollback-resume'
    $rollbackResumeApply = & $importScript `
        -Operation Apply `
        -FolderPath $rollbackResume.GssRoot `
        -ManifestPath $rollbackResume.ManifestPath `
        -SourcePath $rollbackResume.SourcePaths `
        -OutputObject
    Assert-HistoricalRecoveryEqual $rollbackResumeApply.Status 'Committed' 'Rollback-resume fixture commits'
    Assert-HistoricalRecoveryThrowsLike {
        & $importScript `
            -Operation Rollback `
            -FolderPath $rollbackResume.GssRoot `
            -ManifestPath $rollbackResume.ManifestPath `
            -TestFailurePoint AfterFirstRollbackHide `
            -OutputObject
    } '*Injected historical recovery failure after the first rollback hide*' 'Injected rollback-hide interruption'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $rollbackResume.ArchiveRoot -File -Filter '*.rollback-part').Count 1 'Interrupted rollback retains one exact hidden rollback part'
    Assert-HistoricalRecoveryEqual (Get-Content -Raw -LiteralPath $rollbackResume.ReceiptPath | ConvertFrom-Json).state 'NeedsRollback' 'Interrupted rollback receipt requests rollback resume'
    $rollbackResumed = & $importScript `
        -Operation Rollback `
        -FolderPath $rollbackResume.GssRoot `
        -ManifestPath $rollbackResume.ManifestPath `
        -OutputObject
    Assert-HistoricalRecoveryEqual $rollbackResumed.Status 'RolledBack' 'Rollback resumes from an existing rollback part'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $rollbackResume.ArchiveRoot -File -Filter '*.rollback-part').Count 0 'Resumed rollback deletes all exact rollback parts'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $rollbackResume.ArchiveRoot -File -Filter '*.xlsx').Count 0 'Resumed rollback leaves no transaction-published XLSX'
    Assert-HistoricalRecoveryEqual @((Get-Content -Raw -LiteralPath $rollbackResume.LedgerPath | ConvertFrom-Json).entries).Count 0 'Resumed rollback removes transaction ledger baselines'
    $rollbackRepeat = & $importScript `
        -Operation Rollback `
        -FolderPath $rollbackResume.GssRoot `
        -ManifestPath $rollbackResume.ManifestPath `
        -OutputObject
    Assert-HistoricalRecoveryTrue $rollbackRepeat.Idempotent 'Completed rollback rerun is idempotent after artifact revalidation'

    $collision = New-HistoricalRecoveryTestFixture -BasePath $testRoot -Name 'collision'
    $collisionDestination = Join-Path $collision.GssRoot ($collision.Manifest.sources[0].destination_path.Replace('/', '\'))
    New-Item -ItemType Directory -Path (Split-Path -Parent $collisionDestination) -Force | Out-Null
    [System.IO.File]::WriteAllText($collisionDestination, 'not the reviewed workbook')
    Assert-HistoricalRecoveryThrowsLike {
        & $importScript `
            -Operation Plan `
            -FolderPath $collision.GssRoot `
            -ManifestPath $collision.ManifestPath `
            -SourcePath $collision.SourcePaths `
            -OutputObject
    } '*byte size mismatch*' 'Plan fails closed on a destination collision'

    $destinationPrototype = '03 Uploaded Survey Workbooks\Archive - Previous Uploads\Recovered Historical Detail\FY26\FY26-FW16-' + ('0' * 64) + '.xlsx'
    $maxPathTargetLength = 248
    $maxPathNameLength = $maxPathTargetLength - $testRoot.Length - 2 - $destinationPrototype.Length
    Assert-HistoricalRecoveryTrue ($maxPathNameLength -ge 8) 'Near-MAX_PATH fixture name has a safe component length'
    $maxPathName = 'm' + ('x' * ($maxPathNameLength - 1))
    $maxPathFixture = New-HistoricalRecoveryTestFixture -BasePath $testRoot -Name $maxPathName
    $maxPathDestination = Join-Path $maxPathFixture.GssRoot ($maxPathFixture.Manifest.sources[0].destination_path.Replace('/', '\'))
    Assert-HistoricalRecoveryTrue ($maxPathDestination.Length -ge 245 -and $maxPathDestination.Length -lt 260) "Near-MAX_PATH destination remains inside the Windows PowerShell 5.1 boundary: $($maxPathDestination.Length)"
    $maxPathApply = & $importScript `
        -Operation Apply `
        -FolderPath $maxPathFixture.GssRoot `
        -ManifestPath $maxPathFixture.ManifestPath `
        -SourcePath $maxPathFixture.SourcePaths `
        -OutputObject
    Assert-HistoricalRecoveryEqual $maxPathApply.Status 'Committed' 'Near-MAX_PATH Apply succeeds with short same-directory temporary names'
    Assert-HistoricalRecoveryTrue (Test-Path -LiteralPath $maxPathDestination -PathType Leaf) 'Near-MAX_PATH archive destination is published'
    Assert-HistoricalRecoveryEqual @(Get-ChildItem -LiteralPath $maxPathFixture.GssRoot -File -Filter '.gss-*.tmp' -Recurse).Count 0 'Near-MAX_PATH transaction leaves no temporary copy files'

    [pscustomobject][ordered]@{
        Status = 'Passed'
        Test = 'GSS historical recovery transaction'
        SyntheticTransactions = 6
        VerifiedControls = @(
            'strict PII-free manifest',
            'exact file, row, response-set, and date checks',
            'same-week response-set conflict refusal',
            'ledger baseline before XLSX visibility',
            'idempotent committed rerun',
            'resume after ledger interruption',
            'rollback after partial publication',
            'resume from an existing rollback part',
            'collision refusal',
            'Windows PowerShell near-MAX_PATH temporary-copy safety',
            'no live workbook, email, or scheduled-task mutation'
        )
    } | ConvertTo-Json -Depth 5
}
finally {
    $resolvedTestRoot = [System.IO.Path]::GetFullPath($testRoot)
    $temporaryRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
    if (-not $resolvedTestRoot.StartsWith($temporaryRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean a synthetic test path outside the temporary root: $resolvedTestRoot"
    }
    if (Test-Path -LiteralPath $resolvedTestRoot -PathType Container) {
        Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
    }
}
