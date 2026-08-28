[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Analyze-GSS-Run.ps1')

function Assert-Equal {
    param([object]$Actual, [object]$Expected, [string]$Name)
    if ($Actual -ne $Expected) { throw "Assertion failed for $Name. Expected '$Expected'; actual '$Actual'." }
}

function Assert-True {
    param([bool]$Value, [string]$Name)
    if (-not $Value) { throw "Assertion failed for $Name." }
}

function Assert-Near {
    param([double]$Actual, [double]$Expected, [double]$Tolerance, [string]$Name)
    if ([math]::Abs($Actual - $Expected) -gt $Tolerance) {
        throw "Assertion failed for $Name. Expected '$Expected' +/- '$Tolerance'; actual '$Actual'."
    }
}

function Assert-ThrowsLike {
    param([scriptblock]$Script, [string]$Pattern, [string]$Name)
    try { & $Script; throw "Assertion failed for $Name. Expected an exception." }
    catch {
        if ($_.Exception.Message -notlike $Pattern) {
            throw "Assertion failed for $Name. Expected '$Pattern'; actual '$($_.Exception.Message)'."
        }
    }
}

$tokens = $null
$parseErrors = $null
$sourceAst = [System.Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $scriptRoot 'Gss-EmailPackage.ps1'),
    [ref]$tokens,
    [ref]$parseErrors
)
Assert-Equal @($parseErrors).Count 0 'Email package source parses cleanly'

$retryFunctionAst = $sourceAst.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'Invoke-GssFileSystemRetry'
}, $true)
$retryCatchAsts = @($retryFunctionAst.Body.FindAll({
    param($node) $node -is [System.Management.Automation.Language.CatchClauseAst]
}, $true))
$retryRethrows = @($retryCatchAsts[0].Body.FindAll({
    param($node) $node -is [System.Management.Automation.Language.ThrowStatementAst]
}, $true))
Assert-Equal $retryRethrows.Count 1 'Final retry rethrows from inside the operation catch'
Assert-True ($null -eq $retryRethrows[0].Pipeline) 'Final retry uses bare throw'

$packageFunctionAst = $sourceAst.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'New-GssEmailPackage'
}, $true)
$packageTerminalThrowCatches = @($packageFunctionAst.Body.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.CatchClauseAst] -and
        $node.Body.Statements.Count -gt 0 -and
        $node.Body.Statements[-1] -is [System.Management.Automation.Language.ThrowStatementAst]
}, $true))
Assert-Equal $packageTerminalThrowCatches.Count 1 'Package creation has one terminal rethrow catch'
$packageRethrow = $packageTerminalThrowCatches[0].Body.Statements[-1]
Assert-True ($null -eq $packageRethrow.Pipeline) 'Package creation uses bare throw after cleanup'
$packageSource = $packageFunctionAst.Extent.Text
$mutexWaitIndex = $packageSource.IndexOf('$ownsPackageMutex = $packageMutex.WaitOne(0)', [System.StringComparison]::Ordinal)
$runValidationIndex = $packageSource.IndexOf("if (`$RunLog.Mode -ne 'ApplyToMainWorkbook')", [System.StringComparison]::Ordinal)
$mutexReleaseIndex = $packageSource.LastIndexOf('[void]$packageMutex.ReleaseMutex()', [System.StringComparison]::Ordinal)
Assert-True ($sourceAst.Extent.Text.Contains("`$script:GssTransactionMutexName = 'Global\GSSSurveyWorkbookAutomationTransaction'")) 'Email publisher uses the shared workstation transaction mutex name'
Assert-True ($mutexWaitIndex -ge 0 -and $mutexWaitIndex -lt $runValidationIndex) 'Email publisher acquires the transaction mutex before inspecting or mutating package state'
Assert-True ($mutexReleaseIndex -gt $runValidationIndex) 'Email publisher releases its mutex acquisition in the outer finally path'
Assert-Equal ([regex]::Matches($packageSource, [regex]::Escape('$packageMutex.WaitOne(0)')).Count) 1 'Email publisher acquires the mutex exactly once'
Assert-Equal ([regex]::Matches($packageSource, [regex]::Escape('$packageMutex.ReleaseMutex()')).Count) 1 'Email publisher releases the mutex exactly once'
Assert-True ($packageSource.Contains('catch [System.Threading.AbandonedMutexException]')) 'Email publisher treats an abandoned mutex as acquired ownership'
$firstEvidenceValidationIndex = $packageSource.IndexOf('$ledgerRollbackState = Get-GssFeedbackLedgerRollbackState', [System.StringComparison]::Ordinal)
$ledgerReadIndex = $packageSource.IndexOf('$ledger = Read-GssFeedbackLedger', [System.StringComparison]::Ordinal)
$finalSourceRehashIndex = $packageSource.LastIndexOf('foreach ($source in $sourceDescriptors)', [System.StringComparison]::Ordinal)
$finalEvidenceValidationIndex = $packageSource.LastIndexOf('$null = Assert-GssExpectedPackageInputEvidence', [System.StringComparison]::Ordinal)
$ledgerWriteIndex = $packageSource.LastIndexOf('Write-GssFeedbackLedger -Ledger $feedback.NextLedger', [System.StringComparison]::Ordinal)
Assert-True ($firstEvidenceValidationIndex -ge 0 -and $firstEvidenceValidationIndex -lt $ledgerReadIndex) 'Expected package inputs are validated before ledger parsing or analysis'
Assert-True ($finalEvidenceValidationIndex -gt $finalSourceRehashIndex -and $finalEvidenceValidationIndex -lt $ledgerWriteIndex) 'Expected package inputs are revalidated after source rehash and immediately before ledger publication'
Assert-Equal ([regex]::Matches($packageSource, 'Assert-GssExpectedPackageInputEvidence').Count) 3 'Publisher has initial, existing-package, and pre-promotion evidence checks'

$inventoryFunctionAst = $sourceAst.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'Get-GssDetailInventory'
}, $true)
$inventorySource = $inventoryFunctionAst.Extent.Text
$preParseVerificationIndex = $inventorySource.IndexOf("Assert-GssHistoricalRecoveryFileIntegrity -Descriptor `$recoveryDescriptor -Phase 'pre-parse'", [System.StringComparison]::Ordinal)
$workbookParseIndex = $inventorySource.IndexOf('Read-GssDetailWorkbook -Path $file.FullName -FolderPath $FolderPath', [System.StringComparison]::Ordinal)
$postParseVerificationIndex = $inventorySource.IndexOf("Assert-GssHistoricalRecoveryFileIntegrity -Descriptor `$recoveryDescriptor -Phase 'post-parse'", [System.StringComparison]::Ordinal)
Assert-True ($preParseVerificationIndex -ge 0 -and $preParseVerificationIndex -lt $workbookParseIndex) 'Recovered workbook bytes are verified before parsing'
Assert-True ($postParseVerificationIndex -gt $workbookParseIndex) 'Recovered workbook bytes are verified after parsing'

$colonQuoteFixture = [pscustomobject]@{
    sanitized_text = 'Synthetic summary:"placeholder".'
}

function Get-TestDirectoryInventory {
    param([Parameter(Mandatory)][string]$RootPath)

    $resolvedRoot = (Resolve-Path -LiteralPath $RootPath).Path.TrimEnd('\')
    return (@(
        Get-ChildItem -LiteralPath $resolvedRoot -Force -Recurse |
            Sort-Object FullName |
            ForEach-Object {
                $relativePath = $_.FullName.Substring($resolvedRoot.Length) -replace '^[\\/]+', ''
                if ($_.PSIsContainer) {
                    "directory|$relativePath"
                }
                else {
                    "file|$relativePath|$($_.Length)|$(Get-GssSha256 $_.FullName)"
                }
            }
    ) -join "`n")
}
$serializedColonQuoteFixture = $colonQuoteFixture | ConvertTo-Json -Compress
$oneLetterColonQuoteFixture = [pscustomobject][ordered]@{
    value = 'A:"quoted"'
}
$serializedOneLetterColonQuoteFixture = $oneLetterColonQuoteFixture | ConvertTo-Json -Compress
Assert-True ($serializedColonQuoteFixture -match '(?i)(?:[A-Z]:[\\/]|\\\\[^\\])') 'Regression fixture reproduces the legacy serialized-JSON false positive'
Assert-Equal $serializedOneLetterColonQuoteFixture '{"value":"A:\"quoted\""}' 'One-letter colon-quote fixture has the exact serialized JSON shape'
Assert-True ($serializedOneLetterColonQuoteFixture -match '(?i)(?:[A-Z]:[\\/]|\\\\[^\\])') 'One-letter fixture reproduces the legacy serialized-JSON false positive'
Assert-GssPortableContentHasNoMachineSpecificPath -StructuredValues @($colonQuoteFixture, $oneLetterColonQuoteFixture)
Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @(
    $serializedColonQuoteFixture,
    $serializedOneLetterColonQuoteFixture,
    'https://example.invalid/synthetic',
    'https://example.invalid/C:/url-segment?next=//server/share/file'
)
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @(
        'Reference: https://example.invalid/report,C:\Private\report.xlsx'
    )
} '*machine-specific path*' 'URL stripping preserves an adjacent drive path for leakage detection'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @(
        'Reference: https://example.invalid/report;D:/Private/report.xlsx'
    )
} '*machine-specific path*' 'URL stripping preserves a semicolon-adjacent drive path for leakage detection'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @(
        'Reference: https://example.invalid/report)E:\Private\report.xlsx'
    )
} '*machine-specific path*' 'URL stripping preserves a parenthesis-adjacent drive path for leakage detection'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @(
        'Reference: https://example.invalid/report\\fileserver\restricted\report.xlsx'
    )
} '*machine-specific path*' 'URL stripping preserves an adjacent UNC path for leakage detection'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -StructuredValues @(
        [pscustomobject]@{ source = 'C:\Private\report.xlsx' }
    )
} '*machine-specific path*' 'Structured portable content rejects an actual drive path'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -StructuredValues @(
        [pscustomobject]@{ source = '\\fileserver\restricted\report.xlsx' }
    )
} '*machine-specific path*' 'Structured portable content rejects an actual UNC path'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @('Local source: D:/Private/report.xlsx')
} '*machine-specific path*' 'Plain portable text rejects an actual drive path'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @('Local source: C:folder\report.xlsx')
} '*machine-specific path*' 'Plain portable text rejects a drive-relative path'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @('Local source: C:folder/report.xlsx')
} '*machine-specific path*' 'Plain portable text rejects a slash-style drive-relative path'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @('Network source: //fileserver/restricted/report.xlsx')
} '*machine-specific path*' 'Plain portable text rejects a forward-slash network path'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @(
        ([pscustomobject]@{ source = 'C:folder\report.xlsx' } | ConvertTo-Json -Compress)
    )
} '*machine-specific path*' 'Serialized JSON is decoded and rejects an actual drive-relative value'
Assert-ThrowsLike {
    Assert-GssPortableContentHasNoMachineSpecificPath -TextValues @(
        ([pscustomobject]@{ source = '//fileserver/restricted/report.xlsx' } | ConvertTo-Json -Compress)
    )
} '*machine-specific path*' 'Serialized JSON is decoded and rejects an actual forward-slash network value'

function Copy-TestJsonObject {
    param([Parameter(Mandatory)][object]$Value)
    return ($Value | ConvertTo-Json -Depth 20 | ConvertFrom-Json)
}

function ConvertTo-TestExcelColumn {
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

function Add-TestZipText {
    param([object]$Archive, [string]$Name, [string]$Text)
    $entry = $Archive.CreateEntry($Name)
    $stream = $entry.Open()
    try {
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        try { $writer.Write($Text) }
        finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function New-TestDetailWorkbook {
    param([string]$Path, [string[]]$Headers, [object[]]$Records, [int]$BlankFormattedRow = 0)

    Add-Type -AssemblyName System.IO.Compression
    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    try {
        $archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            Add-TestZipText $archive 'xl/workbook.xml' '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sorensen" sheetId="1" r:id="rId1"/></sheets></workbook>'
            Add-TestZipText $archive 'xl/_rels/workbook.xml.rels' '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
            $rowXml = @()
            $allRows = @([pscustomobject]@{ IsHeader = $true; Value = $null }) + @($Records | ForEach-Object { [pscustomobject]@{ IsHeader = $false; Value = $_ } })
            for ($rowIndex = 0; $rowIndex -lt $allRows.Count; $rowIndex++) {
                $excelRow = $rowIndex + 1
                $cells = @()
                for ($columnIndex = 0; $columnIndex -lt $Headers.Count; $columnIndex++) {
                    $header = $Headers[$columnIndex]
                    $value = if ($allRows[$rowIndex].IsHeader) { $header } else {
                        $key = Normalize-GssFeedbackHeader $header
                        $property = $allRows[$rowIndex].Value.PSObject.Properties[$key]
                        if ($property) { [string]$property.Value } else { '' }
                    }
                    if ([string]::IsNullOrEmpty($value) -and -not $allRows[$rowIndex].IsHeader) { continue }
                    $reference = (ConvertTo-TestExcelColumn ($columnIndex + 1)) + $excelRow
                    $escaped = [System.Security.SecurityElement]::Escape([string]$value)
                    $cells += "<c r=`"$reference`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$escaped</t></is></c>"
                }
                $rowXml += "<row r=`"$excelRow`">$($cells -join '')</row>"
            }
            if ($BlankFormattedRow -gt 0) { $rowXml += "<row r=`"$BlankFormattedRow`"></row>" }
            $lastColumn = ConvertTo-TestExcelColumn $Headers.Count
            $lastRow = if ($BlankFormattedRow -gt 0) { $BlankFormattedRow } else { $allRows.Count }
            $worksheet = '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                "<dimension ref=`"A1:$lastColumn$lastRow`"/><sheetData>$($rowXml -join '')</sheetData></worksheet>"
            Add-TestZipText $archive 'xl/worksheets/sheet1.xml' $worksheet
        }
        finally { $archive.Dispose() }
    }
    finally { $stream.Dispose() }
}

function New-TestResponse {
    param([string]$Restaurant, [string]$Date, [string]$Time, [string]$Text, [string]$First, [string]$Last, [string]$Dnc = '')
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
        alertguestsdonotcontact = $Dnc
        eventbookingprocess = ''
        firstvisit = 'N'
        guestfirstname = $First
        guestlastname = $Last
        sorensen = ''
        sorensenweeklycomments = ''
    }
}

function Get-TestCommenterFixture {
    param(
        [int]$CommenterCount,
        [int]$CommenterOverallEventCount,
        [int]$PopulationCount,
        [int]$PopulationOverallEventCount,
        [datetime]$VisitDate,
        [int]$MissingOverallCount = 0,
        [string]$RestaurantId = '9354'
    )

    if ($CommenterOverallEventCount + $MissingOverallCount -gt $CommenterCount) {
        throw 'Synthetic commenter-lens fixture counts are inconsistent.'
    }
    $responses = @()
    for ($index = 0; $index -lt $CommenterCount; $index++) {
        $overall = if ($index -lt $CommenterOverallEventCount) {
            '1'
        }
        elseif ($index -ge ($CommenterCount - $MissingOverallCount)) {
            ''
        }
        else {
            '5'
        }
        $responses += [pscustomobject]@{
            RestaurantId = $RestaurantId
            VisitDate = $VisitDate.Date
            Answers = [pscustomobject]@{
                overall = $overall
                service = '5'
                culinary = '5'
                value = '5'
                paceofmeal = '5'
                recommend = '10'
            }
        }
    }

    $metricDetail = @()
    foreach ($metricPolicy in @($script:GssAnalysisPolicy.commenter_lens.metrics | Where-Object { $null -ne $_.population_metric })) {
        $populationEventCount = if ([string]$metricPolicy.id -eq 'low_overall') {
            $PopulationOverallEventCount
        }
        else {
            0
        }
        $metricDetail += [pscustomobject]@{
            RestaurantId = $RestaurantId
            Metric = [string]$metricPolicy.population_metric
            RawMetric = [string]$metricPolicy.population_metric
            Current = 100.0 * $populationEventCount / $PopulationCount
            CurrentCount = $PopulationCount
            IsCandidate = $false
        }
    }
    return [pscustomobject]@{
        Inventory = [pscustomobject]@{ UniqueResponses = $responses }
        MetricDetail = $metricDetail
    }
}

function Write-TestBytes {
    param([string]$Path, [string]$Value)
    New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force | Out-Null
    [System.IO.File]::WriteAllBytes($Path, [System.Text.Encoding]::UTF8.GetBytes($Value))
}

function New-LoggedEvidence {
    param([string]$Role, [string]$Path, [string]$FolderPath)
    return [pscustomobject]@{
        Role = $Role
        RelativePath = ConvertTo-GssDropboxRelativePath -Path $Path -FolderPath $FolderPath
        ByteSize = [long](Get-Item -LiteralPath $Path).Length
        Sha256 = Get-GssSha256 $Path
    }
}

function Initialize-TestRecoveryInventoryFixture {
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string[]]$Headers
    )

    $fixtureRoot = Join-Path $BasePath 'recovery-inventory'
    $detailRoot = Join-Path $fixtureRoot '03 Uploaded Survey Workbooks'
    $ordinaryArchive = Join-Path $detailRoot 'Archive - Previous Uploads'
    $current = New-TestResponse '9354 Richmond' '07/18/2026' '6:00 PM' 'Current weekly response.' 'Current' 'Guest'
    $ordinary = New-TestResponse '9355 Virginia Beach' '07/04/2026' '7:00 PM' 'Ordinary archive response.' 'Archive' 'Guest'
    New-TestDetailWorkbook -Path (Join-Path $detailRoot 'Sorensen Current.xlsx') -Headers $Headers -Records @($current)
    New-TestDetailWorkbook -Path (Join-Path $ordinaryArchive 'Sorensen Archive.xlsx') -Headers $Headers -Records @($ordinary)

    $correctedWeeks = @(
        [pscustomobject]@{ Week = 'FY26 FW16'; SubjectWeek = 'FY26 FW14'; Date = '09/13/2025' },
        [pscustomobject]@{ Week = 'FY26 FW17'; SubjectWeek = 'FY26 FW15'; Date = '09/20/2025' },
        [pscustomobject]@{ Week = 'FY26 FW26'; SubjectWeek = 'FY26 FW25'; Date = '11/22/2025' },
        [pscustomobject]@{ Week = 'FY26 FW30'; SubjectWeek = ''; Date = '12/20/2025' }
    )
    $manifestSources = @()
    $recoveredPaths = @()
    for ($index = 0; $index -lt $correctedWeeks.Count; $index++) {
        $week = $correctedWeeks[$index]
        $provisionalPath = Join-Path $BasePath "recovered-$index.xlsx"
        $response = New-TestResponse `
            '9354 Richmond' `
            $week.Date `
            '5:00 PM' `
            "Recovered response for $($week.Week)." `
            'Recovered' `
            "Guest$index"
        New-TestDetailWorkbook -Path $provisionalPath -Headers $Headers -Records @($response)
        $sha256 = Get-GssSha256 $provisionalPath
        $portableDestination = "$($script:GssHistoricalRecoveryArchivePrefix)/FY26/$($week.Week.Replace(' ', '-'))-$sha256.xlsx"
        $destinationPath = Join-Path $fixtureRoot $portableDestination.Replace('/', '\')
        New-Item -ItemType Directory -Path (Split-Path -Parent $destinationPath) -Force | Out-Null
        Move-Item -LiteralPath $provisionalPath -Destination $destinationPath
        $parsed = Read-GssDetailWorkbook -Path $destinationPath -FolderPath $fixtureRoot
        $manifestSources += [pscustomobject][ordered]@{
            source_kind = 'synthetic_test'
            drive_file_id = "drivefixture$($index + 1)abcdefghij"
            gmail_message_ids = @("199000000000000$($index + 1)")
            subject_week = $week.SubjectWeek
            source_report_week = $week.Week
            assignment_basis = 'approved_corrected_week_fixture'
            visit_date_start = $parsed.VisitDateStart.ToString('yyyy-MM-dd')
            visit_date_end = $parsed.VisitDateEnd.ToString('yyyy-MM-dd')
            row_count = @($parsed.Responses).Count
            byte_size = [long](Get-Item -LiteralPath $destinationPath).Length
            sha256 = $sha256
            response_set_sha256 = Get-GssHistoricalRecoveryResponseSetSha256 -Workbook $parsed
            destination_path = $portableDestination
            validation = [pscustomobject][ordered]@{
                detail_schema_valid = $true
                response_identity_version = 'gss-feedback-response-identity/v1'
                duplicate_response_count = 0
            }
        }
        $recoveredPaths += $destinationPath
    }

    $manifest = [pscustomobject][ordered]@{
        schema_version = $script:GssHistoricalRecoveryManifestVersion
        fiscal_year = 'FY26'
        created_at_utc = '2026-07-24T12:00:00Z'
        sources = $manifestSources
    }
    $manifestStagingPath = Join-Path $BasePath 'recovery-manifest-staging.json'
    Write-GssUtf8NoBomFile -Path $manifestStagingPath -Value ($manifest | ConvertTo-Json -Depth 20)
    $manifestSha256 = Get-GssSha256 $manifestStagingPath
    $transactionRoot = Join-Path $fixtureRoot "_automation_runs\historical-recovery\$manifestSha256"
    New-Item -ItemType Directory -Path $transactionRoot -Force | Out-Null
    $manifestPath = Join-Path $transactionRoot 'recovery-manifest.json'
    Move-Item -LiteralPath $manifestStagingPath -Destination $manifestPath

    $receiptFiles = @()
    for ($index = 0; $index -lt $manifestSources.Count; $index++) {
        $source = $manifestSources[$index]
        $receiptFiles += [pscustomobject][ordered]@{
            source_index = $index + 1
            sha256 = [string]$source.sha256
            byte_size = [long]$source.byte_size
            row_count = [int]$source.row_count
            response_set_sha256 = [string]$source.response_set_sha256
            destination_path = [string]$source.destination_path
            destination_full_path = $recoveredPaths[$index]
            partial_path = Join-Path (Split-Path -Parent $recoveredPaths[$index]) ".$($source.sha256).recovery-part"
            source_leaf_name = Split-Path -Leaf $recoveredPaths[$index]
            response_hashes = @()
            prepared = $true
            destination_preexisting = $false
            published_by_transaction = $true
        }
    }
    $receipt = [pscustomobject][ordered]@{
        schema_version = $script:GssHistoricalRecoveryReceiptVersion
        classification = $script:GssRestrictedClassification
        contains_personal_data = $true
        manifest_sha256 = $manifestSha256
        manifest_snapshot_path = $manifestPath
        transaction_id = "historical-recovery:$manifestSha256"
        state = 'Committed'
        published_file_count = $receiptFiles.Count
        files = $receiptFiles
    }
    $receiptPath = Join-Path $transactionRoot 'transaction-receipt.json'
    Write-GssUtf8NoBomFile -Path $receiptPath -Value ($receipt | ConvertTo-Json -Depth 20)

    return [pscustomobject]@{
        Folder = $fixtureRoot
        ManifestPath = $manifestPath
        ReceiptPath = $receiptPath
        RecoveredPaths = $recoveredPaths
        CorrectedWeeks = @($correctedWeeks.Week)
    }
}

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('gep_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$folder = Join-Path $temporaryRoot 'GSS Surveys'
try {
    $detailFolder = Join-Path $folder '03 Uploaded Survey Workbooks'
    $archiveFolder = Join-Path $detailFolder 'Archive - Previous Uploads'
    $headers18 = @('Restaurant Name', 'Reservation Date', 'Reservation Time', 'Text', 'Overall', 'Service', 'Culinary', 'Value', 'Pace of Meal', 'Recommend', 'Manager Visit', 'Steak Cooked Correctly', 'Event Booking Process', 'First Visit', 'Guest First Name', 'Guest Last Name', 'Sorensen', 'Sorensen Weekly Comments')
    $headers19 = @('Text', 'Restaurant Name', 'Reservation Time', 'Reservation Date', 'Service', 'Overall', 'Culinary', 'Value', 'Pace of Meal', 'Recommend', 'Manager Visit', 'Steak Cooked Correctly', 'Alert Guests DO NOT CONTACT', 'Event Booking Process', 'First Visit', 'Guest Last Name', 'Guest First Name', 'Sorensen Weekly Comments', 'Sorensen')
    $unsafeBidi = [char]0x202E
    $unsafeC0 = [char]0x0001
    $contactableText = 'Synthetic summary:"All set." A:"Quoted." Casey Testperson praised the service. Contact casey@example.invalid or 212-555-0199 and 555-1212; https://example.invalid and bare.example.invalid/path, Reservation ABC123, Resy ZX9876. Unsafe ' + $unsafeBidi + 'bidi control.'
    $contactable = New-TestResponse '9354 Richmond' '07/10/2026' '6:00 PM' $contactableText 'Casey' 'Testperson'
    $dnc = New-TestResponse '9354 Richmond' '07/11/2026' '7:00 PM' 'Robin Sample said the service was excellent and attentive.' 'Robin' 'Sample' 'NC'
    $old = New-TestResponse '9355 Virginia Beach' '07/01/2026' '5:00 PM' 'The food was great.' 'Taylor' 'Archive'
    $currentDetail = Join-Path $detailFolder 'Sorensen Current.xlsx'
    $archiveDetail = Join-Path $archiveFolder 'Sorensen Archive.xlsx'
    New-TestDetailWorkbook -Path $currentDetail -Headers $headers19 -Records @($contactable, $dnc) -BlankFormattedRow 10000
    New-TestDetailWorkbook -Path $archiveDetail -Headers $headers18 -Records @($contactable, $old)

    $parsed19 = Read-GssDetailWorkbook -Path $currentDetail -FolderPath $folder
    $parsed18 = Read-GssDetailWorkbook -Path $archiveDetail -FolderPath $folder
    Assert-Equal $parsed19.HeaderCount 19 'Reordered 19-column schema'
    Assert-Equal $parsed19.Responses.Count 2 'Blank formatted rows are ignored'
    Assert-Equal $parsed18.HeaderCount 18 '18-column schema without optional DNC'
    Assert-Equal $parsed18.Responses.Count 2 '18-column response parsing'

    $invalidDetail = Join-Path $temporaryRoot 'invalid.xlsx'
    $invalid = New-TestResponse '9354 Richmond' '07/10/2026' '6:00 PM' 'Invalid value fixture.' 'Alex' 'Invalid'
    $invalid.service = '9'
    New-TestDetailWorkbook -Path $invalidDetail -Headers $headers18 -Records @($invalid)
    Assert-ThrowsLike { Read-GssDetailWorkbook -Path $invalidDetail -FolderPath $temporaryRoot } '*Invalid service answer*' 'Invalid detail value blocks readiness'

    $corruptDetail = Join-Path $temporaryRoot 'corrupt.xlsx'
    [System.IO.File]::WriteAllText($corruptDetail, 'not an Open XML workbook')
    Assert-ThrowsLike { Read-GssDetailWorkbook -Path $corruptDetail -FolderPath $temporaryRoot } '*corrupt or unsupported*' 'Corrupt detail workbook blocks readiness'

    $controlProbe = Protect-GssFeedbackText -Text "A${unsafeC0}B${unsafeBidi}C" -KnownNames @()
    Assert-Equal $controlProbe.RedactionCount 2 'C0 and bidi controls are counted as redactions'
    Assert-True (-not [regex]::IsMatch($controlProbe.Text, (Get-GssUnsafeControlPattern))) 'C0 and bidi controls are removed before AI processing'
    $phonePattern = @(Get-GssPiiRedactionRules | Where-Object Label -eq 'phone')[0].Pattern
    foreach ($phoneProbe in @('212-555-0199', '555-1212', '(212) 555-0199', '212.555.0199', '212.5550199', '+1 212.5550199', '2125550199', '+44 20 7946 0958', '(212)5550199', '+1 (212)5550199', '212-5551212')) {
        Assert-True ([regex]::IsMatch($phoneProbe, $phonePattern)) "Phone pattern recognizes $phoneProbe"
        $protectedPhone = Protect-GssFeedbackText -Text "Call $phoneProbe today." -KnownNames @()
        Assert-Equal $protectedPhone.Text 'Call [REDACTED PHONE] today.' "Phone is redacted: $phoneProbe"
        Assert-Equal $protectedPhone.RedactionCount 1 "Phone redaction is counted: $phoneProbe"
        Assert-True ([bool]$protectedPhone.PiiScanPassed) "Post-redaction PII scan passes: $phoneProbe"
        Assert-Equal @($protectedPhone.RemainingPiiTypes).Count 0 "No phone PII remains: $phoneProbe"
    }
    foreach ($nonPhoneProbe in @('78.0952380952381', '78.0952381', '7.12345678', '1.123456', '100.0000000', '1234.12345678', '2026-07-12', 'metric-9354-service')) {
        Assert-True (-not [regex]::IsMatch($nonPhoneProbe, $phonePattern)) "Phone pattern rejects non-phone value $nonPhoneProbe"
        $protectedNonPhone = Protect-GssFeedbackText -Text $nonPhoneProbe -KnownNames @()
        Assert-Equal $protectedNonPhone.Text $nonPhoneProbe "Non-phone value is preserved: $nonPhoneProbe"
        Assert-Equal $protectedNonPhone.RedactionCount 0 "Non-phone value is not counted as a redaction: $nonPhoneProbe"
        Assert-True ([bool]$protectedNonPhone.PiiScanPassed) "Non-phone value passes the PII scan: $nonPhoneProbe"
    }
    $retryProbe = [pscustomobject]@{ Count = 0 }
    Invoke-GssFileSystemRetry -MaxAttempts 3 -DelayMilliseconds 1 -Operation {
        $retryProbe.Count++
        if ($retryProbe.Count -lt 3) { throw [System.IO.IOException]::new('Synthetic transient file lock') }
    }
    Assert-Equal $retryProbe.Count 3 'Transient filesystem operation is retried to success'
    Assert-ThrowsLike {
        Invoke-GssFileSystemRetry -MaxAttempts 2 -DelayMilliseconds 1 -Operation { throw 'Synthetic permanent failure' }
    } '*Synthetic permanent failure*' 'Filesystem retry preserves the final operation error'
    $retryContextProbe = [pscustomobject]@{ InnerError = $null }
    $retryContextError = $null
    try {
        Invoke-GssFileSystemRetry -MaxAttempts 1 -DelayMilliseconds 0 -Operation {
            try {
                throw [System.InvalidOperationException]::new('Synthetic retry context failure')
            }
            catch {
                $retryContextProbe.InnerError = $_
                throw
            }
        }
    }
    catch {
        $retryContextError = $_
    }
    Assert-True ($null -ne $retryContextError) 'Final retry error is rethrown'
    Assert-True ([object]::ReferenceEquals($retryContextProbe.InnerError.Exception, $retryContextError.Exception)) 'Final retry preserves the original exception instance'
    Assert-Equal $retryContextError.FullyQualifiedErrorId $retryContextProbe.InnerError.FullyQualifiedErrorId 'Final retry preserves the original error ID'
    Assert-Equal $retryContextError.InvocationInfo.PositionMessage $retryContextProbe.InnerError.InvocationInfo.PositionMessage 'Final retry preserves the original error position'
    Assert-Equal $retryContextError.ScriptStackTrace $retryContextProbe.InnerError.ScriptStackTrace 'Final retry preserves the original script stack'

    $promotionOutbox = Join-Path $temporaryRoot 'promotion-retry'
    $samePromotionPath = Join-Path $promotionOutbox 'same-path-package'
    New-Item -ItemType Directory -Path $samePromotionPath -Force | Out-Null
    'same-path-owner' | Set-Content -LiteralPath (Join-Path $samePromotionPath 'owner.txt') -Encoding ASCII
    Assert-ThrowsLike {
        Publish-GssStagedEmailPackage `
            -StagingPath $samePromotionPath `
            -PackagePath $samePromotionPath `
            -MaxAttempts 1 `
            -DelayMilliseconds 0 `
            -ValidationOperation { throw 'Validation must not run for identical paths.' }
    } '*must be different*' 'Identical staging and package paths are rejected before cleanup'
    Assert-True (Test-Path -LiteralPath $samePromotionPath -PathType Container) 'Identical-path rejection preserves the source directory'
    Assert-Equal (Get-Content -Raw -LiteralPath (Join-Path $samePromotionPath 'owner.txt')).Trim() 'same-path-owner' 'Identical-path rejection preserves source contents'

    $concurrentStaging = Join-Path $promotionOutbox '.staging-concurrent'
    $concurrentPackage = Join-Path $promotionOutbox 'concurrent-package'
    New-Item -ItemType Directory -Path $concurrentStaging -Force | Out-Null
    New-Item -ItemType Directory -Path $concurrentPackage -Force | Out-Null
    'current-invocation' | Set-Content -LiteralPath (Join-Path $concurrentStaging 'owner.txt') -Encoding ASCII
    'other-invocation' | Set-Content -LiteralPath (Join-Path $concurrentPackage 'owner.txt') -Encoding ASCII
    $concurrentValidationProbe = [pscustomobject]@{ Count = 0 }
    $concurrentPromotionError = $null
    try {
        Publish-GssStagedEmailPackage `
            -StagingPath $concurrentStaging `
            -PackagePath $concurrentPackage `
            -MaxAttempts 1 `
            -DelayMilliseconds 0 `
            -ValidationOperation { $concurrentValidationProbe.Count++ }
    }
    catch {
        $concurrentPromotionError = $_
    }
    Assert-True ($null -ne $concurrentPromotionError) 'Concurrent destination makes promotion fail'
    Assert-Equal $concurrentValidationProbe.Count 0 'Concurrent destination is never validated as this invocation'
    Assert-True (-not (Test-Path -LiteralPath $concurrentStaging)) 'Only this invocation staging directory is cleaned after the promotion race'
    Assert-True (Test-Path -LiteralPath $concurrentPackage -PathType Container) 'Concurrent package is preserved'
    Assert-Equal (Get-Content -Raw -LiteralPath (Join-Path $concurrentPackage 'owner.txt')).Trim() 'other-invocation' 'Concurrent package contents are preserved'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $concurrentPackage '.staging-concurrent'))) 'Staging directory is not nested in a concurrent package'

    $failedStaging = Join-Path $promotionOutbox '.staging-package'
    $failedPackage = Join-Path $promotionOutbox 'package'
    New-Item -ItemType Directory -Path $failedStaging -Force | Out-Null
    'ready' | Set-Content -LiteralPath (Join-Path $failedStaging 'READY') -Encoding ASCII
    $validationProbe = [pscustomobject]@{ Count = 0 }
    Assert-ThrowsLike {
        Publish-GssStagedEmailPackage `
            -StagingPath $failedStaging `
            -PackagePath $failedPackage `
            -MaxAttempts 3 `
            -DelayMilliseconds 1 `
            -ValidationOperation {
                $validationProbe.Count++
                if (-not (Test-Path -LiteralPath (Join-Path $failedPackage 'READY') -PathType Leaf)) {
                    throw 'Synthetic package was not promoted before validation.'
                }
                throw [System.IO.IOException]::new('Synthetic promoted-package validation failure')
            }
    } '*Synthetic promoted-package validation failure*' 'Post-promotion validation failure is surfaced'
    Assert-Equal $validationProbe.Count 3 'Post-promotion validation uses bounded retries'
    Assert-True (-not (Test-Path -LiteralPath $failedStaging)) 'Failed staging directory is absent after promotion'
    Assert-True (-not (Test-Path -LiteralPath $failedPackage)) 'Package promoted by the failing invocation is cleaned up'

    New-Item -ItemType Directory -Path $failedStaging -Force | Out-Null
    'ready' | Set-Content -LiteralPath (Join-Path $failedStaging 'READY') -Encoding ASCII
    $validationResult = Publish-GssStagedEmailPackage `
        -StagingPath $failedStaging `
        -PackagePath $failedPackage `
        -MaxAttempts 3 `
        -DelayMilliseconds 1 `
        -ValidationOperation {
            if (-not (Test-Path -LiteralPath (Join-Path $failedPackage 'READY') -PathType Leaf)) {
                throw 'Synthetic rebuilt package validation failed.'
            }
            'validated'
        }
    Assert-Equal $validationResult 'validated' 'Cleaned package path can be rebuilt and validated'
    Assert-True (Test-Path -LiteralPath $failedPackage -PathType Container) 'Rebuilt package is promoted'

    $inventory = Get-GssDetailInventory -FolderPath $folder -ReportingDate ([datetime]'2026-07-12')
    Assert-Equal $inventory.CurrentWorkbook.PortablePath '03 Uploaded Survey Workbooks/Sorensen Current.xlsx' 'Current detail selection by visit date'
    Assert-Equal $inventory.UniqueResponses.Count 3 'Overlapping exports deduplicate responses'
    Assert-Equal $inventory.DuplicateResponseCount 1 'Duplicate response count'
    $knownGuestNames = @(Get-GssKnownGuestNames $inventory.AllResponseInstances)
    Assert-True ($knownGuestNames -contains 'Casey' -and $knownGuestNames -contains 'Robin') 'Distinct same-length guest names remain in the redaction set'
    Assert-ThrowsLike {
        Get-GssDetailInventory -FolderPath $folder -ReportingDate ([datetime]'2026-07-13')
    } '*reporting date must be a Sunday*' 'Detail inventory rejects a non-Sunday reporting date'

    $recoveryFixture = Initialize-TestRecoveryInventoryFixture -BasePath $temporaryRoot -Headers $headers19
    $recoveredInventory = Get-GssDetailInventory `
        -FolderPath $recoveryFixture.Folder `
        -ReportingDate ([datetime]'2026-07-19')
    Assert-Equal @($recoveredInventory.Workbooks).Count 6 'Recovery inventory retains normal current/archive files and four recovered files'
    $recoveredPortablePaths = @(
        $recoveredInventory.Workbooks.PortablePath |
            Where-Object { $_ -like "$($script:GssHistoricalRecoveryArchivePrefix)/*" }
    )
    Assert-Equal $recoveredPortablePaths.Count 4 'Four corrected recovered weeks are accepted'
    foreach ($correctedWeek in $recoveryFixture.CorrectedWeeks) {
        Assert-True (
            @($recoveredPortablePaths | Where-Object {
                $_ -like "*/$($correctedWeek.Replace(' ', '-'))-*.xlsx"
            }).Count -eq 1
        ) "Corrected recovery assignment is represented exactly once: $correctedWeek"
    }

    $tamperedPath = $recoveryFixture.RecoveredPaths[0]
    $tamperedOriginalBytes = [System.IO.File]::ReadAllBytes($tamperedPath)
    [System.IO.File]::WriteAllText($tamperedPath, 'tampered recovered workbook')
    Assert-ThrowsLike {
        Get-GssDetailInventory -FolderPath $recoveryFixture.Folder -ReportingDate ([datetime]'2026-07-19')
    } '*hash or size mismatch*' 'Recovered inventory rejects substituted workbook bytes'
    [System.IO.File]::WriteAllBytes($tamperedPath, $tamperedOriginalBytes)

    $missingPath = $recoveryFixture.RecoveredPaths[1]
    $hiddenMissingPath = "$missingPath.missing"
    Move-Item -LiteralPath $missingPath -Destination $hiddenMissingPath
    Assert-ThrowsLike {
        Get-GssDetailInventory -FolderPath $recoveryFixture.Folder -ReportingDate ([datetime]'2026-07-19')
    } '*is missing*' 'Recovered inventory rejects a missing committed workbook'
    Move-Item -LiteralPath $hiddenMissingPath -Destination $missingPath

    $allMissingMoves = @()
    foreach ($recoveredPath in $recoveryFixture.RecoveredPaths) {
        $hiddenPath = "$recoveredPath.missing"
        Move-Item -LiteralPath $recoveredPath -Destination $hiddenPath
        $allMissingMoves += [pscustomobject]@{ Original = $recoveredPath; Hidden = $hiddenPath }
    }
    Assert-ThrowsLike {
        Get-GssDetailInventory -FolderPath $recoveryFixture.Folder -ReportingDate ([datetime]'2026-07-19')
    } '*is missing*' 'Committed evidence rejects deletion of the entire recovered XLSX set'
    foreach ($move in $allMissingMoves) {
        Move-Item -LiteralPath $move.Hidden -Destination $move.Original
    }

    $extraPath = Join-Path (Split-Path -Parent $recoveryFixture.RecoveredPaths[2]) 'FY26-FW31-unmanifested.xlsx'
    Copy-Item -LiteralPath $recoveryFixture.RecoveredPaths[2] -Destination $extraPath
    Assert-ThrowsLike {
        Get-GssDetailInventory -FolderPath $recoveryFixture.Folder -ReportingDate ([datetime]'2026-07-19')
    } '*not covered by a committed manifest*' 'Recovered inventory rejects an extra XLSX'
    Remove-Item -LiteralPath $extraPath -Force

    $receiptOriginalBytes = [System.IO.File]::ReadAllBytes($recoveryFixture.ReceiptPath)
    $nonCommittedReceipt = Get-Content -Raw -LiteralPath $recoveryFixture.ReceiptPath | ConvertFrom-Json
    $nonCommittedReceipt.state = 'Validated'
    Write-GssUtf8NoBomFile -Path $recoveryFixture.ReceiptPath -Value ($nonCommittedReceipt | ConvertTo-Json -Depth 20)
    Assert-ThrowsLike {
        Get-GssDetailInventory -FolderPath $recoveryFixture.Folder -ReportingDate ([datetime]'2026-07-19')
    } '*no committed historical-recovery receipt covers it*' 'Recovered inventory rejects noncommitted receipt evidence'
    [System.IO.File]::WriteAllBytes($recoveryFixture.ReceiptPath, $receiptOriginalBytes)

    $manifestOriginalBytes = [System.IO.File]::ReadAllBytes($recoveryFixture.ManifestPath)
    $manifestText = [System.IO.File]::ReadAllText($recoveryFixture.ManifestPath)
    [System.IO.File]::WriteAllText(
        $recoveryFixture.ManifestPath,
        $manifestText + [Environment]::NewLine,
        (New-Object System.Text.UTF8Encoding($false, $true))
    )
    Assert-ThrowsLike {
        Get-GssDetailInventory -FolderPath $recoveryFixture.Folder -ReportingDate ([datetime]'2026-07-19')
    } '*manifest snapshot hash mismatch*' 'Recovered inventory rejects changed manifest bytes'
    [System.IO.File]::WriteAllBytes($recoveryFixture.ManifestPath, $manifestOriginalBytes)

    $pdf = Join-Path $folder '04 Email Comparison PDFs\GSS Email Comparison 071226.pdf'
    $rolling = Join-Path $folder '02 Weekly Rolling Source Workbooks\Sorensen Rolling.xlsx'
    $priorRolling = Join-Path $folder '02 Weekly Rolling Source Workbooks\Sorensen Prior Year.xlsx'
    $liveWorkbook = Join-Path $folder '01 Main Workbook\GSS Score Trends - Main.xlsx'
    $logPath = Join-Path $folder '_automation_runs\logs\gss_update_fixture.json'
    Write-TestBytes $pdf '%PDF synthetic comparison'
    Write-TestBytes $rolling 'synthetic rolling workbook'
    Write-TestBytes $priorRolling 'synthetic prior rolling workbook'
    Write-TestBytes $liveWorkbook 'synthetic live workbook'
    Write-TestBytes $logPath '{}'
    $runLog = [pscustomobject]@{
        Mode = 'ApplyToMainWorkbook'
        CurrentWeekEnding = '2026-07-12'
        PriorYearWeekEnding = '2025-07-13'
        EmailComparisonPdf = $pdf
        CurrentSourceWorkbook = $rolling
        PriorYearSourceWorkbook = $priorRolling
        TargetWorkbook = $liveWorkbook
        FileEvidence = @(
            (New-LoggedEvidence 'comparison_pdf' $pdf $folder),
            (New-LoggedEvidence 'rolling_workbook' $rolling $folder),
            (New-LoggedEvidence 'prior_year_rolling_workbook' $priorRolling $folder),
            (New-LoggedEvidence 'live_workbook' $liveWorkbook $folder)
        )
    }
    $finding = [pscustomobject]@{
        EvidenceId = 'metric-9354-service'
        RestaurantId = '9354'
        Entity = '9354 Richmond'
        Metric = 'Service'
        RawMetric = 'Service'
        Category = 'Service'
        LowerIsBetter = $false
        Current = 78.0952380952381
        CurrentCount = 125
        ChangeVsPreviousRollingWindow = 2.0
        YoYImprovement = 5.0
        VsAllFranchisees = 1.0
        VsSorensenTotal = 1.0
        PersistentMovement = $false
        IsCandidate = $true
        CandidateDirection = 'Improvement'
        CandidateMagnitude = 5.0
        Corroboration = @()
        BaseActionItem = $true
    }
    $populationMetricDetail = @()
    $populationEventCounts = @{
        low_overall = 13
        low_service = 20
        low_culinary = 15
        low_value = 10
        low_pace = 5
    }
    foreach ($restaurantId in @('9354', '9355')) {
        $populationCount = if ($restaurantId -eq '9354') { 125 } else { 100 }
        foreach ($metricPolicy in @($script:GssAnalysisPolicy.commenter_lens.metrics | Where-Object { $null -ne $_.population_metric })) {
            $eventCount = if ($restaurantId -eq '9354') { [int]$populationEventCounts[[string]$metricPolicy.id] } else { 10 }
            $populationMetricDetail += [pscustomobject]@{
                RestaurantId = $restaurantId
                Metric = [string]$metricPolicy.population_metric
                RawMetric = [string]$metricPolicy.population_metric
                Current = 100.0 * $eventCount / $populationCount
                CurrentCount = $populationCount
                IsCandidate = $false
            }
        }
    }

    $fixtureReportingDate = [datetime]'2026-07-12'
    $currentLensFixture = Get-TestCommenterFixture `
        -CommenterCount 309 `
        -CommenterOverallEventCount 55 `
        -PopulationCount 534 `
        -PopulationOverallEventCount 69 `
        -VisitDate $fixtureReportingDate
    $currentLens = Get-GssCommenterLens `
        -Inventory $currentLensFixture.Inventory `
        -MetricDetail $currentLensFixture.MetricDetail `
        -ReportingDate $fixtureReportingDate
    Assert-GssCommenterLensContract -CommenterLens $currentLens
    $currentRestaurantLens = @($currentLens.restaurants | Where-Object restaurant_id -eq '9354')[0]
    $currentOverallLens = @($currentRestaurantLens.metrics | Where-Object metric_id -eq 'low_overall')[0]
    $currentRecommendLens = @($currentRestaurantLens.metrics | Where-Object metric_id -eq 'recommend_detractor')[0]
    Assert-Equal $currentLens.status 'Ready' 'Commenter-only lens is reviewable when its aggregate definitions are available'
    Assert-Equal $currentLens.scope_label 'Among guests who provided comments' 'Commenter lens carries the bounded scope label'
    Assert-True (-not [bool]$currentLens.source_design.population_raw_rows_available) 'Commenter lens explicitly records that population raw rows are unavailable'
    Assert-True ($null -eq $currentRestaurantLens.comment_coverage_pct) 'Unverified partition alignment suppresses cross-source comment coverage'
    Assert-Equal $currentRestaurantLens.comment_coverage_status 'SuppressedUnverifiedPartitionAlignment' 'Unverified partition alignment carries an explicit coverage status'
    Assert-Equal ([int]$currentOverallLens.commenter_event_count) 55 'Current commenter event count'
    Assert-Near ([double]$currentOverallLens.commenter_event_rate_pct) (100.0 * 55 / 309) 0.00000001 'Current commenter event rate'
    Assert-Near ([double]$currentOverallLens.population_event_rate_pct) (100.0 * 69 / 534) 0.00000001 'Current population aggregate event rate'
    Assert-True ($null -eq $currentOverallLens.commenter_minus_population_percentage_points) 'Unverified partition alignment suppresses the commenter-minus-population gap'
    Assert-True ($null -eq $currentOverallLens.reconstructed_population_event_count) 'Unverified partition alignment suppresses cross-source event-count reconstruction'
    Assert-Equal $currentOverallLens.comparison_status 'SuppressedUnverifiedPartitionAlignment' 'Current cross-source comparison is suppressed without verified partition alignment'
    Assert-Equal $currentOverallLens.derived_non_comment_status 'SuppressedUnverifiedPartitionAlignment' 'Non-comment subtraction is suppressed without exact partition alignment'
    Assert-True ($null -eq $currentOverallLens.derived_non_comment_response_count) 'Unverified alignment suppresses derived non-comment denominator'
    Assert-True ($null -eq $currentOverallLens.derived_non_comment_event_count) 'Unverified alignment suppresses derived non-comment event count'
    Assert-True ($null -eq $currentOverallLens.derived_non_comment_event_rate_pct) 'Unverified alignment suppresses derived non-comment event rate'
    Assert-True (-not [bool]$currentLens.reporting_window.exact_partition_alignment_verified) 'Default visit-date inventory is explicitly not exact-partition aligned'
    Assert-True ($null -eq $currentRecommendLens.population_metric) 'Recommend remains commenter-only without an exact population metric'
    Assert-True ($null -eq $currentRecommendLens.population_event_rate_pct) 'Recommend population rate is unavailable'
    Assert-True ($null -eq $currentRecommendLens.commenter_minus_population_percentage_points) 'Recommend population gap is unavailable'
    Assert-Equal $currentRecommendLens.comparison_status 'NotAvailableNoExactPopulationMetric' 'Missing Recommend population metric is not treated as a data-quality failure'

    $unexpectedTopLevelLens = Copy-TestJsonObject $currentLens
    $unexpectedTopLevelLens | Add-Member -NotePropertyName raw_rows -NotePropertyValue @([pscustomobject]@{ guest_email = 'not-allowed@example.invalid' })
    Assert-ThrowsLike {
        Assert-GssCommenterLensContract -CommenterLens $unexpectedTopLevelLens
    } "*unsupported property 'raw_rows'*" 'Commenter schema rejects arbitrary top-level row data'

    $unexpectedRestaurantLens = Copy-TestJsonObject $currentLens
    $unexpectedRestaurantLens.restaurants[0] | Add-Member -NotePropertyName guest_email_copy -NotePropertyValue 'not-allowed@example.invalid'
    Assert-ThrowsLike {
        Assert-GssCommenterLensContract -CommenterLens $unexpectedRestaurantLens
    } "*unsupported property 'guest_email_copy'*" 'Commenter schema rejects arbitrary restaurant-level keys'

    $unexpectedMetricLens = Copy-TestJsonObject $currentLens
    $unexpectedMetricLens.restaurants[0].metrics[0] | Add-Member -NotePropertyName source_row_copy -NotePropertyValue 42
    Assert-ThrowsLike {
        Assert-GssCommenterLensContract -CommenterLens $unexpectedMetricLens
    } "*unsupported property 'source_row_copy'*" 'Commenter schema rejects arbitrary metric-level keys'

    $structuredLimitationLens = Copy-TestJsonObject $currentLens
    $structuredLimitationLens.limitations = @([pscustomobject]@{ guest_email = 'not-allowed@example.invalid' })
    Assert-ThrowsLike {
        Assert-GssCommenterLensContract -CommenterLens $structuredLimitationLens
    } '*every limitation must be a scalar string*' 'Commenter schema rejects structured limitation payloads'

    $structuredIssueLens = Copy-TestJsonObject $currentLens
    $structuredIssueLens.restaurants[0].data_quality_issues = @([pscustomobject]@{ source_row = 42 })
    Assert-ThrowsLike {
        Assert-GssCommenterLensContract -CommenterLens $structuredIssueLens
    } '*data-quality issues must be scalar strings*' 'Commenter schema rejects structured data-quality payloads'

    $validCommenterCsv = ConvertTo-GssCommenterLensCsv -CommenterLens $currentLens
    Assert-GssCommenterLensCsvContract -CsvText $validCommenterCsv
    $unexpectedCommenterCsvHeader = 'raw_row,' + $validCommenterCsv
    Assert-ThrowsLike {
        Assert-GssCommenterLensCsvContract -CsvText $unexpectedCommenterCsvHeader
    } '*header column count is invalid*' 'Commenter CSV rejects arbitrary columns'

    $exactPartitionLens = Get-GssCommenterLens `
        -Inventory $currentLensFixture.Inventory `
        -MetricDetail $currentLensFixture.MetricDetail `
        -ReportingDate $fixtureReportingDate `
        -ExactPartitionAlignmentVerified
    Assert-GssCommenterLensContract -CommenterLens $exactPartitionLens
    $exactOverallLens = @($exactPartitionLens.restaurants[0].metrics | Where-Object metric_id -eq 'low_overall')[0]
    Assert-True ([bool]$exactPartitionLens.reporting_window.exact_partition_alignment_verified) 'Exact partition alignment requires an explicit caller assertion'
    Assert-Near ([double]$exactPartitionLens.restaurants[0].comment_coverage_pct) (100.0 * 309 / 534) 0.00000001 'Verified exact partition permits comment coverage'
    Assert-Equal $exactPartitionLens.restaurants[0].comment_coverage_status 'Ready' 'Verified exact partition marks coverage ready'
    Assert-Near ([double]$exactOverallLens.commenter_minus_population_percentage_points) 4.87800444 0.00000001 'Verified exact partition permits the commenter-minus-population gap'
    Assert-Equal ([int]$exactOverallLens.reconstructed_population_event_count) 69 'Verified exact partition permits population event-count reconstruction'
    Assert-Equal $exactOverallLens.comparison_status 'Ready' 'Verified exact partition permits the aggregate comparison'
    Assert-Equal $exactOverallLens.derived_non_comment_status 'Ready' 'Verified exact partition permits derived non-comment arithmetic'
    Assert-Equal ([int]$exactOverallLens.derived_non_comment_response_count) 225 'Verified exact partition non-comment denominator'
    Assert-Equal ([int]$exactOverallLens.derived_non_comment_event_count) 14 'Verified exact partition non-comment event count'
    Assert-Near ([double]$exactOverallLens.derived_non_comment_event_rate_pct) (100.0 * 14 / 225) 0.00000001 'Verified exact partition non-comment event rate'

    $previousLensFixture = Get-TestCommenterFixture `
        -CommenterCount 333 `
        -CommenterOverallEventCount 57 `
        -PopulationCount 621 `
        -PopulationOverallEventCount 76 `
        -VisitDate $fixtureReportingDate
    $previousLens = Get-GssCommenterLens `
        -Inventory $previousLensFixture.Inventory `
        -MetricDetail $previousLensFixture.MetricDetail `
        -ReportingDate $fixtureReportingDate
    $previousOverallLens = @($previousLens.restaurants[0].metrics | Where-Object metric_id -eq 'low_overall')[0]
    Assert-Near ([double]$previousOverallLens.commenter_event_rate_pct) (100.0 * 57 / 333) 0.00000001 'Previous commenter event rate'
    Assert-Near ([double]$previousOverallLens.population_event_rate_pct) (100.0 * 76 / 621) 0.00000001 'Previous population aggregate event rate'
    Assert-True ($null -eq $previousOverallLens.commenter_minus_population_percentage_points) 'Previous cross-source gap is suppressed without verified partition alignment'
    Assert-True ($null -eq $previousOverallLens.reconstructed_population_event_count) 'Previous event-count reconstruction is suppressed without verified partition alignment'

    $missingScoreFixture = Get-TestCommenterFixture `
        -CommenterCount 3 `
        -CommenterOverallEventCount 1 `
        -PopulationCount 10 `
        -PopulationOverallEventCount 2 `
        -VisitDate $fixtureReportingDate `
        -MissingOverallCount 1
    $missingScoreLens = Get-GssCommenterLens `
        -Inventory $missingScoreFixture.Inventory `
        -MetricDetail $missingScoreFixture.MetricDetail `
        -ReportingDate $fixtureReportingDate `
        -ExactPartitionAlignmentVerified
    Assert-GssCommenterLensContract -CommenterLens $missingScoreLens
    $missingOverallLens = @($missingScoreLens.restaurants[0].metrics | Where-Object metric_id -eq 'low_overall')[0]
    Assert-Equal ([int]$missingOverallLens.commenter_scored_response_count) 2 'Missing-score commenter denominator uses nonmissing scores'
    Assert-Equal ([int]$missingOverallLens.commenter_missing_score_count) 1 'Missing commenter score is explicit'
    Assert-Near ([double]$missingOverallLens.commenter_event_rate_pct) 50.0 0.00000001 'Missing-score commenter rate uses the nonmissing denominator'
    Assert-Near ([double]$missingOverallLens.population_event_rate_pct) 20.0 0.00000001 'Population rate retains the aggregate denominator'
    Assert-Equal $missingOverallLens.comparison_status 'SuppressedMissingCommenterScores' 'Missing commenter score suppresses the cross-source gap'
    Assert-True ($null -eq $missingOverallLens.commenter_minus_population_percentage_points) 'Missing commenter score leaves no cross-source gap'
    Assert-True ($null -eq $missingOverallLens.material_gap) 'Missing commenter score leaves no material-gap flag'
    Assert-Equal $missingOverallLens.derived_non_comment_status 'SuppressedMissingCommenterScores' 'Missing commenter score suppresses non-comment derivation'
    Assert-True ($null -eq $missingOverallLens.derived_non_comment_event_count) 'Missing commenter score leaves no derived event count'

    $boundaryFixture = Get-TestCommenterFixture `
        -CommenterCount 4 `
        -CommenterOverallEventCount 0 `
        -PopulationCount 10 `
        -PopulationOverallEventCount 0 `
        -VisitDate $fixtureReportingDate
    $windowStart = $fixtureReportingDate.AddDays(-90)
    $boundaryFixture.Inventory.UniqueResponses[0].VisitDate = $windowStart
    $boundaryFixture.Inventory.UniqueResponses[1].VisitDate = $windowStart.AddDays(-1)
    $boundaryFixture.Inventory.UniqueResponses[2].VisitDate = $fixtureReportingDate
    $boundaryFixture.Inventory.UniqueResponses[3].VisitDate = $fixtureReportingDate.AddDays(1)
    $boundaryLens = Get-GssCommenterLens `
        -Inventory $boundaryFixture.Inventory `
        -MetricDetail $boundaryFixture.MetricDetail `
        -ReportingDate $fixtureReportingDate
    Assert-Equal ([int]$boundaryLens.restaurants[0].commenter_response_count) 2 'Inclusive 13-week visit-date window includes both boundaries only'

    $definitionMismatchFixture = Get-TestCommenterFixture `
        -CommenterCount 2 `
        -CommenterOverallEventCount 1 `
        -PopulationCount 10 `
        -PopulationOverallEventCount 2 `
        -VisitDate $fixtureReportingDate
    $definitionMismatchFixture.MetricDetail[0].Metric = 'Different aggregate'
    $definitionMismatchFixture.MetricDetail[0].RawMetric = 'Different aggregate'
    $definitionMismatchLens = Get-GssCommenterLens `
        -Inventory $definitionMismatchFixture.Inventory `
        -MetricDetail $definitionMismatchFixture.MetricDetail `
        -ReportingDate $fixtureReportingDate `
        -ExactPartitionAlignmentVerified
    $definitionMismatchOverall = @($definitionMismatchLens.restaurants[0].metrics | Where-Object metric_id -eq 'low_overall')[0]
    Assert-Equal $definitionMismatchLens.status 'DataQualityReview' 'Population metric-definition mismatch requires data-quality review'
    Assert-Equal $definitionMismatchOverall.comparison_status 'SuppressedPopulationMetricDefinitionMismatch' 'Definition mismatch suppresses the population comparison'
    Assert-True ($null -eq $definitionMismatchOverall.commenter_minus_population_percentage_points) 'Definition mismatch leaves no population gap'

    $overCoverageFixture = Get-TestCommenterFixture `
        -CommenterCount 2 `
        -CommenterOverallEventCount 1 `
        -PopulationCount 1 `
        -PopulationOverallEventCount 1 `
        -VisitDate $fixtureReportingDate
    $overCoverageLens = Get-GssCommenterLens `
        -Inventory $overCoverageFixture.Inventory `
        -MetricDetail $overCoverageFixture.MetricDetail `
        -ReportingDate $fixtureReportingDate `
        -ExactPartitionAlignmentVerified
    Assert-Equal $overCoverageLens.status 'DataQualityReview' 'Commenter count above population count requires data-quality review'
    Assert-True (@($overCoverageLens.restaurants[0].data_quality_issues) -contains 'commenter_count_exceeds_population_count') 'Over-coverage reason is explicit'
    Assert-True ($null -eq $overCoverageLens.restaurants[0].comment_coverage_pct) 'Invalid over-coverage does not publish a percentage'
    Assert-Equal $overCoverageLens.restaurants[0].metrics[0].comparison_status 'SuppressedInvalidCoverage' 'Invalid over-coverage suppresses comparison'

    $deduplicatedLens = Get-GssCommenterLens `
        -Inventory $inventory `
        -MetricDetail $populationMetricDetail `
        -ReportingDate $fixtureReportingDate
    Assert-Equal (@($deduplicatedLens.restaurants | Measure-Object -Property commenter_response_count -Sum).Sum) 3 'Commenter lens consumes the deduplicated unique-response inventory'
    $serializedCurrentLens = $currentLens | ConvertTo-Json -Depth 20 -Compress
    foreach ($forbiddenRowField in @('response_hash', 'guest_first_name', 'guest_last_name', 'source_path', 'source_row', 'sanitized_text')) {
        Assert-True ($serializedCurrentLens -notmatch ('(?i)"' + [regex]::Escape($forbiddenRowField) + '"\s*:')) "Commenter lens excludes row-level field $forbiddenRowField"
    }
    Assert-True (-not [bool]$currentLens.claims.population_prevalence_claimed) 'Commenter lens makes no population prevalence claim'
    Assert-True (-not [bool]$currentLens.claims.statistical_significance_claimed) 'Commenter lens makes no significance claim'
    Assert-True (-not [bool]$currentLens.claims.causal_driver_claimed) 'Commenter lens makes no causal claim'
    Assert-True (-not [bool]$currentLens.claims.individual_prediction_produced) 'Commenter lens makes no individual prediction'
    Assert-True (
        @($currentLens.limitations | Where-Object {
            $_ -match '(?i)stable vendor response IDs' -and $_ -match '(?i)low residual risk'
        }).Count -eq 1
    ) 'Commenter lens discloses the low-residual stable vendor response-ID limitation'
    Assert-ThrowsLike {
        Get-GssCommenterLens `
            -Inventory $currentLensFixture.Inventory `
            -MetricDetail $currentLensFixture.MetricDetail `
            -ReportingDate ([datetime]'2026-07-13')
    } '*reporting date must be a Sunday*' 'Commenter lens rejects a non-Sunday reporting date'

    $analysis = [pscustomobject]@{
        WorkbookStatus = 'Ready'
        AnalysisStatus = 'Review'
        EmailReadiness = 'Ready'
        LogPath = $logPath
        MetricDetail = @($finding) + $populationMetricDetail
        RestaurantFindings = @(Select-GssRestaurantFindings @($finding))
    }
    $ledgerPath = Join-Path $folder '_automation_runs\state\gss_feedback_first_seen.json'
    $packageInputDescriptors = @(
        (Get-GssSourceDescriptor -Role 'comparison_pdf' -Path $pdf -FolderPath $folder),
        (Get-GssSourceDescriptor -Role 'rolling_workbook' -Path $rolling -FolderPath $folder),
        (Get-GssSourceDescriptor -Role 'prior_year_rolling_workbook' -Path $priorRolling -FolderPath $folder),
        (Get-GssSourceDescriptor -Role 'live_workbook' -Path $liveWorkbook -FolderPath $folder),
        (Get-GssSourceDescriptor -Role 'detail_workbook' -Path $inventory.CurrentWorkbook.Path -FolderPath $folder),
        (Get-GssSourceDescriptor -Role 'run_log' -Path $logPath -FolderPath $folder)
    )
    $packageInputDescriptors += @($inventory.Workbooks |
        Where-Object { $_.PortablePath -ne $inventory.CurrentWorkbook.PortablePath } |
        Sort-Object PortablePath |
        ForEach-Object {
            Get-GssSourceDescriptor -Role 'detail_archive_workbook' -Path $_.Path -FolderPath $folder
        })
    $expectedPackageInputEvidence = Get-GssCurrentPackageInputEvidence `
        -SourceDescriptors $packageInputDescriptors `
        -LedgerPath $ledgerPath
    Assert-Equal @($expectedPackageInputEvidence.Inputs).Count 7 'Expected package-input evidence captures every package source while the ledger is absent'
    Assert-Equal (
        Get-GssPackageInputEvidenceSha256 -Inputs $expectedPackageInputEvidence.Inputs
    ) $expectedPackageInputEvidence.SourceSetSha256 'Expected package-input evidence carries its exact canonical source-set hash'
    $validatedPackageInputEvidence = Assert-GssExpectedPackageInputEvidence `
        -ExpectedEvidence $expectedPackageInputEvidence `
        -SourceDescriptors $packageInputDescriptors `
        -LedgerPath $ledgerPath
    Assert-Equal $validatedPackageInputEvidence.SourceSetSha256 $expectedPackageInputEvidence.SourceSetSha256 'Exact package-input evidence validates'

    $invalidSourceSetEvidence = Copy-TestJsonObject $expectedPackageInputEvidence
    $invalidSourceSetEvidence.SourceSetSha256 = ('0' * 64)
    Assert-ThrowsLike {
        Assert-GssExpectedPackageInputEvidence `
            -ExpectedEvidence $invalidSourceSetEvidence `
            -SourceDescriptors $packageInputDescriptors `
            -LedgerPath $ledgerPath
    } '*SourceSetSha256 does not match its Inputs*' 'Expected package-input evidence rejects a forged aggregate hash'

    $missingInputEvidence = [pscustomobject]@{
        Inputs = @($expectedPackageInputEvidence.Inputs | Select-Object -Skip 1)
        SourceSetSha256 = Get-GssPackageInputEvidenceSha256 -Inputs @($expectedPackageInputEvidence.Inputs | Select-Object -Skip 1)
    }
    Assert-ThrowsLike {
        Assert-GssExpectedPackageInputEvidence `
            -ExpectedEvidence $missingInputEvidence `
            -SourceDescriptors $packageInputDescriptors `
            -LedgerPath $ledgerPath
    } '*unexpected source*' 'Expected package-input evidence rejects an extra live input'

    $extraInputRecords = @($expectedPackageInputEvidence.Inputs) + [pscustomobject][ordered]@{
        PortablePath = 'gss/synthetic/extra-input.bin'
        ByteSize = [long]1
        Sha256 = ('1' * 64)
    }
    $extraInputEvidence = [pscustomobject]@{
        Inputs = $extraInputRecords
        SourceSetSha256 = Get-GssPackageInputEvidenceSha256 -Inputs $extraInputRecords
    }
    Assert-ThrowsLike {
        Assert-GssExpectedPackageInputEvidence `
            -ExpectedEvidence $extraInputEvidence `
            -SourceDescriptors $packageInputDescriptors `
            -LedgerPath $ledgerPath
    } '*missing from the current source set*' 'Expected package-input evidence rejects a missing live input'

    $driftedInputEvidence = Copy-TestJsonObject $expectedPackageInputEvidence
    $driftedInputEvidence.Inputs[0].Sha256 = ('2' * 64)
    $driftedInputEvidence.SourceSetSha256 = Get-GssPackageInputEvidenceSha256 -Inputs @($driftedInputEvidence.Inputs)
    Assert-ThrowsLike {
        Assert-GssExpectedPackageInputEvidence `
            -ExpectedEvidence $driftedInputEvidence `
            -SourceDescriptors $packageInputDescriptors `
            -LedgerPath $ledgerPath
    } '*changed after the committed snapshot*' 'Expected package-input evidence rejects hash drift'

    $priorRollingBytes = [System.IO.File]::ReadAllBytes($priorRolling)
    [System.IO.File]::AppendAllText($priorRolling, 'mutation after evidence capture')
    Assert-ThrowsLike {
        Assert-GssExpectedPackageInputEvidence `
            -ExpectedEvidence $expectedPackageInputEvidence `
            -SourceDescriptors $packageInputDescriptors `
            -LedgerPath $ledgerPath
    } '*changed after the committed snapshot*' 'Expected package-input evidence detects mutation after capture'
    [System.IO.File]::WriteAllBytes($priorRolling, $priorRollingBytes)

    $rollbackProbePath = Join-Path $temporaryRoot 'ledger-rollback-probe.json'
    $priorLedgerBytes = [System.Text.Encoding]::UTF8.GetBytes("{`"prior`":true}`r`n")
    [System.IO.File]::WriteAllBytes($rollbackProbePath, $priorLedgerBytes)
    $priorLedgerState = Get-GssFeedbackLedgerRollbackState -Path $rollbackProbePath
    [System.IO.File]::WriteAllText($rollbackProbePath, '{"prior":false}')
    Restore-GssFeedbackLedgerState -Path $rollbackProbePath -State $priorLedgerState
    Assert-Equal ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($rollbackProbePath))) ([Convert]::ToBase64String($priorLedgerBytes)) 'Ledger rollback restores exact prior bytes'
    Remove-Item -LiteralPath $rollbackProbePath -Force
    $absentLedgerState = Get-GssFeedbackLedgerRollbackState -Path $rollbackProbePath
    [System.IO.File]::WriteAllText($rollbackProbePath, '{"created":true}')
    Restore-GssFeedbackLedgerState -Path $rollbackProbePath -State $absentLedgerState
    Assert-True (-not (Test-Path -LiteralPath $rollbackProbePath)) 'Ledger rollback removes a ledger that was originally absent'

    Write-GssFeedbackLedger `
        -Ledger ([pscustomobject]@{ schema_version = $script:GssFeedbackLedgerVersion; entries = @() }) `
        -Path $ledgerPath
    Assert-ThrowsLike {
        Assert-GssExpectedPackageInputEvidence `
            -ExpectedEvidence $expectedPackageInputEvidence `
            -SourceDescriptors $packageInputDescriptors `
            -LedgerPath $ledgerPath
    } '*unexpected source*gss/_automation_runs/state/gss_feedback_first_seen.json*' 'Expected package-input evidence treats a newly appeared ledger as an extra input'
    $ledgerBoundEvidence = Get-GssCurrentPackageInputEvidence `
        -SourceDescriptors $packageInputDescriptors `
        -LedgerPath $ledgerPath
    Assert-Equal @($ledgerBoundEvidence.Inputs).Count 8 'Current package-input evidence includes the present feedback ledger bytes'
    $null = Assert-GssExpectedPackageInputEvidence `
        -ExpectedEvidence $ledgerBoundEvidence `
        -SourceDescriptors $packageInputDescriptors `
        -LedgerPath $ledgerPath
    Remove-Item -LiteralPath $ledgerPath -Force

    $originalModelingReason = [string](Get-GssPopulationModelingAvailability).Reason
    $analysis | Add-Member -NotePropertyName Modeling -NotePropertyValue (Get-GssPopulationModelingAvailability) -Force
    $analysis.Modeling.Reason = 'Synthetic network source //fileserver/restricted/report.xlsx must not publish.'
    Assert-ThrowsLike {
        New-GssEmailPackage `
            -FolderPath $folder `
            -RunLog $runLog `
            -AnalysisResult $analysis `
            -LedgerPath $ledgerPath `
            -ExpectedPackageInputEvidence $expectedPackageInputEvidence
    } '*machine-specific path*' 'End-to-end package publication rejects a forward-slash network path'
    $analysis.Modeling.Reason = $originalModelingReason

    $originalPublisher = (Get-Command Publish-GssStagedEmailPackage).ScriptBlock
    $publicationContextProbe = [pscustomobject]@{ InnerError = $null; Lock = $null; StagingPath = $null }
    $publicationContextError = $null
    try {
        Set-Item -LiteralPath Function:Publish-GssStagedEmailPackage -Value {
            param(
                [string]$StagingPath,
                [string]$PackagePath,
                [scriptblock]$ValidationOperation,
                [int]$MaxAttempts = 12,
                [int]$DelayMilliseconds = 500
            )
            $publicationContextProbe.StagingPath = $StagingPath
            $publicationContextProbe.Lock = [System.IO.File]::Open(
                (Join-Path $StagingPath 'READY'),
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::None
            )
            try {
                throw [System.IO.IOException]::new('Synthetic publication context failure')
            }
            catch {
                $publicationContextProbe.InnerError = $_
                throw
            }
        }
        try {
            $null = New-GssEmailPackage `
                -FolderPath $folder `
                -RunLog $runLog `
                -AnalysisResult $analysis `
                -LedgerPath $ledgerPath `
                -ExpectedPackageInputEvidence $expectedPackageInputEvidence
        }
        catch {
            $publicationContextError = $_
        }
    }
    finally {
        if ($publicationContextProbe.Lock) { $publicationContextProbe.Lock.Dispose() }
        Set-Item -LiteralPath Function:Publish-GssStagedEmailPackage -Value $originalPublisher
    }
    Assert-True ($null -ne $publicationContextError) 'Publication error is rethrown after staging cleanup'
    Assert-True ([object]::ReferenceEquals($publicationContextProbe.InnerError.Exception, $publicationContextError.Exception)) 'Publication catch preserves the original exception instance'
    Assert-Equal $publicationContextError.FullyQualifiedErrorId $publicationContextProbe.InnerError.FullyQualifiedErrorId 'Publication catch preserves the original error ID'
    Assert-Equal $publicationContextError.InvocationInfo.PositionMessage $publicationContextProbe.InnerError.InvocationInfo.PositionMessage 'Publication catch preserves the original error position'
    Assert-Equal $publicationContextError.ScriptStackTrace $publicationContextProbe.InnerError.ScriptStackTrace 'Publication catch preserves the original script stack'
    Assert-True $publicationContextError.Exception.Data.Contains('GssStagingCleanupError') 'Publication error retains staging cleanup failure detail'
    Assert-True (Test-Path -LiteralPath $publicationContextProbe.StagingPath -PathType Container) 'Locked staging directory remains available for recovery'
    Assert-True (-not (Test-Path -LiteralPath $ledgerPath)) 'Failed publication removes the ledger that was absent before the attempt'
    $retryableEvidence = Assert-GssExpectedPackageInputEvidence `
        -ExpectedEvidence $expectedPackageInputEvidence `
        -SourceDescriptors $packageInputDescriptors `
        -LedgerPath $ledgerPath
    Assert-Equal $retryableEvidence.SourceSetSha256 $expectedPackageInputEvidence.SourceSetSha256 'Failed publication leaves the exact source evidence retryable'
    Remove-Item -LiteralPath $publicationContextProbe.StagingPath -Recurse -Force

    $package = New-GssEmailPackage `
        -FolderPath $folder `
        -RunLog $runLog `
        -AnalysisResult $analysis `
        -LedgerPath $ledgerPath `
        -ExpectedPackageInputEvidence $expectedPackageInputEvidence
    Assert-Equal $package.EmailReadiness 'Ready' 'Package email readiness'
    Assert-True (-not [bool]$package.ExistingPackage) 'Retry after failed publication builds and promotes a new package'
    Assert-True (Test-Path -LiteralPath $package.ReadyMarkerPath -PathType Leaf) 'Ready marker exists'
    $initialLedgerWriteTimeUtc = (Get-Item -LiteralPath $ledgerPath).LastWriteTimeUtc
    $initialReadyWriteTimeUtc = (Get-Item -LiteralPath $package.ReadyMarkerPath).LastWriteTimeUtc

    $outerPackageMutex = New-Object System.Threading.Mutex($false, $script:GssTransactionMutexName)
    $ownsOuterPackageMutex = $false
    try {
        $ownsOuterPackageMutex = $outerPackageMutex.WaitOne(0)
        Assert-True $ownsOuterPackageMutex 'Reentrant publisher test acquires the outer transaction mutex'
        $invalidMutexRunLog = Copy-TestJsonObject $runLog
        $invalidMutexRunLog.Mode = 'CopyTest'
        Assert-ThrowsLike {
            New-GssEmailPackage `
                -FolderPath $folder `
                -RunLog $invalidMutexRunLog `
                -AnalysisResult $analysis `
                -LedgerPath $ledgerPath
        } '*successful live apply log*' 'Nested publisher mutex acquisition releases its own count after an exception'
        $reentrantPackage = New-GssEmailPackage `
            -FolderPath $folder `
            -RunLog $runLog `
            -AnalysisResult $analysis `
            -LedgerPath $ledgerPath
        Assert-True ([bool]$reentrantPackage.ExistingPackage) 'Publisher reenters the mutex already held by the safe coordinator thread'
    }
    finally {
        if ($ownsOuterPackageMutex) { [void]$outerPackageMutex.ReleaseMutex() }
        $outerPackageMutex.Dispose()
    }

    $ledgerBeforeConcurrentLoser = [System.IO.File]::ReadAllBytes($ledgerPath)
    $readyBeforeConcurrentLoser = [System.IO.File]::ReadAllBytes($package.ReadyMarkerPath)
    $outboxRoot = Split-Path -Parent $package.PackagePath
    $outboxInventoryBeforeConcurrentLoser = Get-TestDirectoryInventory -RootPath $outboxRoot
    $outboxBeforeConcurrentLoser = @(
        Get-ChildItem -LiteralPath $outboxRoot -Directory |
            Where-Object { $_.Name -notlike '.staging-*' }
    ).Count
    $mutexMarkerPath = Join-Path $temporaryRoot 'package-mutex-held.marker'
    $mutexReleasePath = Join-Path $temporaryRoot 'package-mutex-release.marker'
    $escapedMutexMarkerPath = $mutexMarkerPath.Replace("'", "''")
    $escapedMutexReleasePath = $mutexReleasePath.Replace("'", "''")
    $escapedMutexName = $script:GssTransactionMutexName.Replace("'", "''")
    $mutexHolderCommand = @"
`$mutex = New-Object System.Threading.Mutex(`$false, '$escapedMutexName')
`$ownsMutex = `$false
try {
    `$ownsMutex = `$mutex.WaitOne(0)
    if (-not `$ownsMutex) { exit 2 }
    [System.IO.File]::WriteAllText('$escapedMutexMarkerPath', 'held')
    for (`$attempt = 1; `$attempt -le 1000 -and -not (Test-Path -LiteralPath '$escapedMutexReleasePath' -PathType Leaf); `$attempt++) {
        Start-Sleep -Milliseconds 10
    }
}
finally {
    if (`$ownsMutex) { [void]`$mutex.ReleaseMutex() }
    `$mutex.Dispose()
}
"@
    $encodedMutexHolderCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($mutexHolderCommand))
    $mutexHolderProcess = $null
    try {
        $mutexHolderProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList @(
            '-NoLogo',
            '-NoProfile',
            '-NonInteractive',
            '-EncodedCommand',
            $encodedMutexHolderCommand
        ) -WindowStyle Hidden -PassThru
        for ($waitAttempt = 1; $waitAttempt -le 500 -and -not (Test-Path -LiteralPath $mutexMarkerPath -PathType Leaf); $waitAttempt++) {
            Start-Sleep -Milliseconds 10
        }
        Assert-True (Test-Path -LiteralPath $mutexMarkerPath -PathType Leaf) 'Competing publisher process acquires the shared transaction mutex'
        Assert-ThrowsLike {
            New-GssEmailPackage `
                -FolderPath $folder `
                -RunLog $runLog `
                -AnalysisResult $analysis `
                -LedgerPath $ledgerPath
        } '*Another GSS workbook transaction is already active*' 'Competing publisher is rejected before it can capture or roll back ledger state'
        Assert-Equal ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($ledgerPath))) ([Convert]::ToBase64String($ledgerBeforeConcurrentLoser)) 'Rejected competing publisher cannot undo the successful ledger'
        Assert-Equal ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($package.ReadyMarkerPath))) ([Convert]::ToBase64String($readyBeforeConcurrentLoser)) 'Rejected competing publisher cannot alter the successful READY package'
        Assert-Equal (Get-TestDirectoryInventory -RootPath $outboxRoot) $outboxInventoryBeforeConcurrentLoser 'Rejected competing publisher cannot alter any package file or leave a staging directory'
        Assert-Equal @(
            Get-ChildItem -LiteralPath $outboxRoot -Directory |
                Where-Object { $_.Name -notlike '.staging-*' }
        ).Count $outboxBeforeConcurrentLoser 'Rejected competing publisher creates no additional visible package'
    }
    finally {
        [System.IO.File]::WriteAllText($mutexReleasePath, 'release')
        if ($null -ne $mutexHolderProcess -and -not $mutexHolderProcess.WaitForExit(5000)) {
            $mutexHolderProcess.Kill()
            $mutexHolderProcess.WaitForExit()
        }
        if ($null -ne $mutexHolderProcess) { $mutexHolderProcess.Dispose() }
        foreach ($markerPath in @($mutexMarkerPath, $mutexReleasePath)) {
            if (Test-Path -LiteralPath $markerPath -PathType Leaf) { Remove-Item -LiteralPath $markerPath -Force }
        }
    }

    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    foreach ($portableName in @('email_manifest.json', 'analysis.json', 'commenter_lens.json', 'commenter_lens.csv', 'email_preview.txt', 'email_preview.html', 'RESTRICTED.txt')) {
        $portablePath = Join-Path $package.PackagePath $portableName
        $portableBytes = [System.IO.File]::ReadAllBytes($portablePath)
        $hasUtf8Bom = $portableBytes.Length -ge 3 -and
            $portableBytes[0] -eq 0xEF -and
            $portableBytes[1] -eq 0xBB -and
            $portableBytes[2] -eq 0xBF
        Assert-True (-not $hasUtf8Bom) "$portableName is UTF-8 without a BOM"
        try {
            $null = $strictUtf8.GetString($portableBytes)
        }
        catch [System.Text.DecoderFallbackException] {
            throw "Assertion failed for strict UTF-8 package output: $portableName."
        }
    }
    $manifest = Read-GssUtf8NoBomFile $package.ManifestPath | ConvertFrom-Json
    $analysisJson = Read-GssUtf8NoBomFile (Join-Path $package.PackagePath 'analysis.json') | ConvertFrom-Json
    $commenterLensJson = Read-GssUtf8NoBomFile $package.CommenterLensJsonPath | ConvertFrom-Json
    $commenterLensCsv = Read-GssUtf8NoBomFile $package.CommenterLensCsvPath
    Assert-Equal $manifest.schema_version 'gss-email-package/v2' 'Package schema version'
    Assert-Equal $manifest.policy_version 'gss-analysis-policy/v4' 'Versioned analysis policy'
    Assert-Equal $manifest.classification $script:GssRestrictedClassification 'Package restricted personal-data classification'
    Assert-True ([bool]$manifest.package_contains_personal_data) 'Package explicitly contains personal data'
    Assert-True (-not [bool]$manifest.distribution_controls.automatic_sending_enabled) 'Automatic sending remains disabled'
    Assert-True ([bool]$manifest.distribution_controls.human_review_required) 'Human review remains required'
    Assert-True ([bool]$manifest.distribution_controls.restricted_recipient_review.required_before_manual_send) 'Restricted-recipient review is required before manual send'
    Assert-Equal $manifest.distribution_controls.restricted_recipient_review.status 'pending_manual_confirmation' 'Restricted-recipient review status is explicit'
    Assert-Equal $manifest.distribution_controls.ready_marker_meaning 'Integrity-validated and ready for manual content/recipient review; never PII-free or send-approved.' 'READY meaning is unambiguous'
    Assert-True ([string]$manifest.distribution_controls.ready_marker_meaning -match '(?i)never PII-free') 'READY never means PII-free'
    Assert-Equal $analysisJson.classification $script:GssRestrictedClassification 'Analysis carries package classification'
    Assert-Equal $package.DataClassification $script:GssRestrictedClassification 'Package result carries data classification'
    Assert-True (-not [bool]$package.AutomaticSendingEnabled) 'Package result confirms automatic sending is disabled'
    Assert-True ([string]$manifest.feedback_selection_sha256 -match '^[a-f0-9]{64}$') 'Manifest carries the selected-feedback fingerprint'
    Assert-Equal $analysisJson.feedback_selection_sha256 $manifest.feedback_selection_sha256 'Analysis and manifest feedback selection agree'
    Assert-Equal @($manifest.attachments).Count 3 'Three attachment policy'
    Assert-Equal @($manifest.portable_artifacts).Count 6 'Six hash-bound non-attachment package artifacts'
    foreach ($artifact in @($manifest.portable_artifacts)) {
        $artifactPath = Join-Path $package.PackagePath ([string]$artifact.path).Replace('/', '\')
        Assert-True (@($manifest.attachments.path) -notcontains [string]$artifact.path) "Portable artifact is not an email attachment: $($artifact.role)"
        Assert-Equal (Get-GssSha256 $artifactPath) ([string]$artifact.sha256) "Portable artifact hash for $($artifact.role)"
        Assert-Equal ([long](Get-Item -LiteralPath $artifactPath).Length) ([long]$artifact.byte_size) "Portable artifact size for $($artifact.role)"
    }
    Assert-Equal $manifest.commenter_lens_json_path 'commenter_lens.json' 'Manifest points to the reviewer commenter-lens JSON'
    Assert-Equal $manifest.commenter_lens_csv_path 'commenter_lens.csv' 'Manifest points to the reviewer commenter-lens CSV'
    Assert-True (Test-Path -LiteralPath $package.CommenterLensJsonPath -PathType Leaf) 'Commenter-lens JSON exists'
    Assert-True (Test-Path -LiteralPath $package.CommenterLensCsvPath -PathType Leaf) 'Commenter-lens CSV exists'
    Assert-True (@($manifest.attachments.path) -notcontains 'commenter_lens.json') 'Commenter-lens JSON is not a fourth attachment'
    Assert-True (@($manifest.attachments.path) -notcontains 'commenter_lens.csv') 'Commenter-lens CSV is not a fourth attachment'
    $detailAttachment = @($manifest.attachments | Where-Object role -eq 'detail_workbook')
    Assert-Equal $detailAttachment.Count 1 'One raw guest detail attachment'
    Assert-True ([bool]$detailAttachment[0].contains_personal_data) 'Raw guest detail attachment is marked as containing personal data'
    Assert-Equal $detailAttachment[0].classification $script:GssRestrictedClassification 'Raw guest detail attachment restricted classification'
    Assert-True ([string]$detailAttachment[0].path -like '*RESTRICTED*') 'Raw guest detail attachment has a visible restricted filename'
    Assert-Equal $manifest.classification_notice_path 'RESTRICTED.txt' 'Package carries a visible restricted classification notice'
    Assert-True ((Read-GssUtf8NoBomFile (Join-Path $package.PackagePath 'RESTRICTED.txt')) -match [regex]::Escape($script:GssRestrictedClassification)) 'Package notice carries the exact restricted classification'
    Assert-Equal @($manifest.sources).Count 7 'Portable source evidence count including run log and archive detail'
    Assert-Equal @($manifest.sources | Where-Object role -eq 'run_log').Count 1 'Exact run log source evidence'
    Assert-Equal @($manifest.sources | Where-Object role -eq 'detail_archive_workbook').Count 1 'Archived detail source evidence'
    Assert-Equal $manifest.source_log_path ($manifest.sources | Where-Object role -eq 'run_log' | Select-Object -ExpandProperty path) 'Run-log path matches source descriptor'
    Assert-Equal $analysisJson.methodology.feedback_response_count 2 'Initial package includes only current-file new feedback'
    Assert-Equal $analysisJson.methodology.feedback_visit_date_start '2026-07-10' 'Actual feedback visit-date start'
    Assert-Equal $analysisJson.methodology.feedback_visit_date_end '2026-07-11' 'Actual feedback visit-date end'
    Assert-Equal $commenterLensJson.scope_label 'Among guests who provided comments' 'Packaged commenter lens carries the bounded scope'
    Assert-Equal $commenterLensJson.status 'Ready' 'Packaged commenter lens is reviewable'
    Assert-True (-not [bool]$commenterLensJson.source_design.population_raw_rows_available) 'Package explicitly records unavailable population raw rows'
    Assert-True (-not [bool]$commenterLensJson.reporting_window.exact_partition_alignment_verified) 'Package does not assert exact partition alignment'
    Assert-True ([string]$commenterLensJson.reporting_window.source_alignment -match '(?i)not exact source-report-week row alignment') 'Package discloses visit-date versus rolling-period alignment'
    Assert-True (
        @($commenterLensJson.limitations | Where-Object {
            $_ -match '(?i)stable vendor response IDs' -and $_ -match '(?i)low residual risk'
        }).Count -eq 1
    ) 'Packaged commenter lens carries the stable vendor response-ID limitation'
    Assert-Equal ([int](@($commenterLensJson.restaurants | Measure-Object -Property commenter_response_count -Sum).Sum)) 3 'Packaged commenter lens uses all deduplicated comment-bearing rows in the window'
    $packagedRichmondLens = @($commenterLensJson.restaurants | Where-Object restaurant_id -eq '9354')[0]
    $packagedOverallLens = @($packagedRichmondLens.metrics | Where-Object metric_id -eq 'low_overall')[0]
    $packagedRecommendLens = @($packagedRichmondLens.metrics | Where-Object metric_id -eq 'recommend_detractor')[0]
    Assert-Equal ([int]$packagedRichmondLens.commenter_response_count) 2 'Packaged Richmond commenter count'
    Assert-True ($null -eq $packagedRichmondLens.comment_coverage_pct) 'Packaged cross-source coverage is suppressed without verified alignment'
    Assert-Equal $packagedRichmondLens.comment_coverage_status 'SuppressedUnverifiedPartitionAlignment' 'Packaged coverage carries the alignment-suppression status'
    Assert-Equal ([int]$packagedOverallLens.commenter_event_count) 0 'Packaged Richmond low-overall commenter event count'
    Assert-True ($null -eq $packagedOverallLens.commenter_minus_population_percentage_points) 'Packaged commenter-minus-population gap is suppressed without verified alignment'
    Assert-Equal $packagedOverallLens.comparison_status 'SuppressedUnverifiedPartitionAlignment' 'Packaged population comparison carries the alignment-suppression status'
    Assert-Equal $packagedOverallLens.derived_non_comment_status 'SuppressedUnverifiedPartitionAlignment' 'Packaged non-comment derivation is alignment-gated'
    Assert-True ($null -eq $packagedOverallLens.derived_non_comment_response_count) 'Packaged lens has no derived non-comment denominator'
    Assert-True ($null -eq $packagedOverallLens.derived_non_comment_event_count) 'Packaged lens has no derived non-comment event count'
    Assert-True ($null -eq $packagedRecommendLens.population_metric) 'Packaged Recommend metric remains commenter-only'
    Assert-Equal $packagedRecommendLens.comparison_status 'NotAvailableNoExactPopulationMetric' 'Packaged Recommend metric is unavailable without causing data-quality review'
    Assert-True (-not [bool]$commenterLensJson.claims.population_prevalence_claimed) 'Packaged commenter lens makes no prevalence claim'
    Assert-True (-not [bool]$commenterLensJson.claims.statistical_significance_claimed) 'Packaged commenter lens makes no significance claim'
    Assert-True (-not [bool]$commenterLensJson.claims.causal_driver_claimed) 'Packaged commenter lens makes no causal claim'
    Assert-True (-not [bool]$commenterLensJson.claims.individual_prediction_produced) 'Packaged commenter lens makes no individual prediction'
    Assert-True (-not [bool]$commenterLensJson.row_level_data_persisted) 'Packaged commenter lens persists no row-level data'
    Assert-Equal $analysisJson.commenter_lens.scope_label $commenterLensJson.scope_label 'Analysis and restricted commenter-lens JSON agree'
    Assert-Equal $analysisJson.commenter_lens.status $commenterLensJson.status 'Analysis and restricted commenter-lens status agree'
    Assert-True ($commenterLensCsv -match 'exact_partition_alignment_verified') 'Commenter-lens CSV exposes the partition-alignment gate'
    Assert-True ($commenterLensCsv -match 'SuppressedUnverifiedPartitionAlignment') 'Commenter-lens CSV exposes non-comment suppression'
    Assert-True ($commenterLensCsv -notmatch '(?i)response_hash|guest_first_name|guest_last_name|source_path|source_row|sanitized_text') 'Commenter-lens CSV contains no row-level fields'
    Assert-Equal @($analysisJson.sanitized_feedback).Count 1 'DNC response excluded from individual evidence'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match 'Synthetic summary:"All set\."') 'Structured colon-quote text remains publishable'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match 'A:"Quoted\."') 'One-letter colon-quote text remains publishable end to end'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match '\[REDACTED URL\],') 'Bare URL is redacted without swallowing sentence punctuation'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match '\[REDACTED PHONE\]') 'Short local phone number is redacted'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match '\[REDACTED BOOKING ID\]') 'Reservation and Resy identifiers are redacted'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match '\[REDACTED CONTROL\]') 'Unsafe C0 and bidi controls are visibly removed'
    $serviceTheme = $analysisJson.feedback_themes | Where-Object category -eq 'service'
    Assert-Equal $serviceTheme.unique_response_count 2 'Theme requires two unique responses'
    Assert-Equal $serviceTheme.denominator_response_count 2 'Theme declares its unique-response denominator'
    Assert-Equal $serviceTheme.visit_date_start '2026-07-10' 'Theme denominator visit-date start'
    Assert-Equal $serviceTheme.visit_date_end '2026-07-11' 'Theme denominator visit-date end'
    Assert-True ([bool]$serviceTheme.categories_are_non_exclusive) 'Theme categories are explicitly non-exclusive'
    Assert-Equal $serviceTheme.do_not_contact_count 1 'DNC contributes only to aggregate theme counts'
    Assert-Equal @($serviceTheme.quotable_evidence_ids).Count 1 'DNC is not quotable'
    $serviceThemeEvidence = $analysisJson.theme_evidence | Where-Object theme_id -eq $serviceTheme.theme_id
    Assert-Equal $serviceThemeEvidence.theme_id $serviceTheme.theme_id 'Qualified theme uses the deterministic theme ID'
    Assert-True (-not [string]::IsNullOrWhiteSpace([string]$serviceThemeEvidence.display_text)) 'Qualified theme has deterministic display text'
    Assert-Equal $serviceThemeEvidence.restaurant_name 'Richmond' 'Theme evidence restaurant name'
    Assert-True (@($manifest.theme_ids) -contains $serviceThemeEvidence.theme_id) 'Manifest includes qualified theme ID'
    Assert-True (@($manifest.evidence_ids) -notcontains $serviceThemeEvidence.theme_id) 'Theme ID remains separate from evidence IDs'
    Assert-True (-not [string]::IsNullOrWhiteSpace([string]$analysisJson.metric_evidence[0].display_text)) 'Metric evidence has deterministic display text'
    Assert-Equal ([double]$analysisJson.metric_evidence[0].rolling_value) ([double]$finding.Current) 'High-precision native metric survives package serialization'
    Assert-Equal ([double]$analysisJson.metric_evidence[0].change_vs_previous_window) ([double]$finding.ChangeVsPreviousRollingWindow) 'Metric evidence previous-window delta contract'
    Assert-Equal ([double]$analysisJson.metric_evidence[0].change_vs_prior_year) ([double]$finding.YoYImprovement) 'Metric evidence prior-year delta contract'
    Assert-Equal ([double]$analysisJson.metric_evidence[0].vs_franchise) ([double]$finding.VsAllFranchisees) 'Metric evidence franchise delta contract'
    Assert-Equal $analysisJson.metric_evidence[0].metric_key $finding.RawMetric 'Metric evidence key contract'
    Assert-Equal $analysisJson.metric_evidence[0].direction 'higher_is_better' 'Metric evidence direction contract'
    Assert-True ($analysisJson.metric_evidence[0].lower_is_better -is [bool] -and -not $analysisJson.metric_evidence[0].lower_is_better) 'Metric evidence lower-is-better Boolean contract'
    Assert-Equal ([int]$serviceThemeEvidence.unique_response_count) 2 'Theme evidence response-count contract'
    Assert-Equal ([int]$serviceThemeEvidence.denominator_response_count) 2 'Theme evidence denominator contract'
    Assert-True ([string]$serviceThemeEvidence.display_text -match '2 of 2') 'Theme evidence uses N of M wording'
    Assert-True ([string]$serviceThemeEvidence.display_text -match '(?i)non-exclusive') 'Theme evidence states categories are non-exclusive'
    Assert-Equal ([int]$serviceThemeEvidence.concern_count) 0 'Theme evidence concern-count contract'
    Assert-Equal ([int]$serviceThemeEvidence.positive_count) 2 'Theme evidence positive-count contract'
    Assert-Equal ([int]$serviceThemeEvidence.do_not_contact_count) 1 'Theme evidence do-not-contact-count contract'
    Assert-True (Test-GssAnalysisEvidenceContract -Analysis $analysisJson) 'Serialized analysis satisfies structured evidence contract'
    Assert-Equal $analysisJson.metric_evidence[0].level.rolling_weeks 13 'Metric level is explicit'
    Assert-Equal $analysisJson.metric_evidence[0].response_volume_tier 'High' 'Metric response-volume tier is explicit'
    Assert-Equal $analysisJson.metric_evidence[0].confidence_tier 'High' 'Metric confidence tier is explicit'
    Assert-Equal $analysisJson.metric_evidence[0].confidence_tier $analysisJson.metric_evidence[0].response_volume_tier 'Confidence tier remains only a compatibility alias'
    Assert-Equal $analysisJson.metric_evidence[0].level.response_volume_tier 'High' 'Metric level carries response-volume tier'
    Assert-Equal $analysisJson.metric_evidence[0].movement.adjacent_window_overlap_weeks 12 'Metric movement declares 12-week adjacent overlap'
    Assert-Equal ([double]$analysisJson.metric_evidence[0].benchmark.vs_all_franchisees) ([double]$finding.VsAllFranchisees) 'Metric benchmark is explicit'
    Assert-Equal $analysisJson.methodology.analysis_description 'risk-reduced' 'Analysis is described as risk-reduced'
    Assert-True ([bool]$analysisJson.methodology.human_review_required) 'Analysis methodology requires human review'
    Assert-True (-not [bool]$analysisJson.methodology.commenter_exact_partition_alignment_verified) 'Methodology does not claim exact commenter/population partition alignment'
    Assert-True ([string]$analysisJson.methodology.commenter_alignment_basis -match '(?i)not exact source-report-week row alignment') 'Methodology discloses source-period alignment limits'
    Assert-Equal $analysisJson.reporting.adjacent_window_overlap_weeks 12 'Reporting metadata declares adjacent-window overlap'

    $nullableComparison = Copy-TestJsonObject $analysisJson
    $nullableMetricId = [string]$nullableComparison.metric_evidence[0].evidence_id
    $nullableComparison.metric_evidence[0].change_vs_previous_window = $null
    ($nullableComparison.evidence_cards | Where-Object { [string]$_.evidence_id -eq $nullableMetricId }).change_vs_previous_window = $null
    Assert-True (Test-GssAnalysisEvidenceContract -Analysis $nullableComparison) 'Nullable metric comparison remains valid'

    $missingRollingValue = Copy-TestJsonObject $analysisJson
    $missingRollingValue.metric_evidence[0].PSObject.Properties.Remove('rolling_value')
    Assert-ThrowsLike { Test-GssAnalysisEvidenceContract -Analysis $missingRollingValue } '*rolling_value*' 'Missing rolling value violates evidence contract'

    $stringRollingValue = Copy-TestJsonObject $analysisJson
    $stringRollingValue.metric_evidence[0].rolling_value = '78.1'
    Assert-ThrowsLike { Test-GssAnalysisEvidenceContract -Analysis $stringRollingValue } '*native finite number*' 'Quoted metric number violates evidence contract'

    $directionConflict = Copy-TestJsonObject $analysisJson
    $directionConflict.metric_evidence[0].lower_is_better = $true
    Assert-ThrowsLike { Test-GssAnalysisEvidenceContract -Analysis $directionConflict } '*direction conflicts*' 'Metric direction and Boolean must agree'

    $fractionalThemeCount = Copy-TestJsonObject $analysisJson
    $fractionalThemeCount.theme_evidence[0].concern_count = 0.5
    Assert-ThrowsLike { Test-GssAnalysisEvidenceContract -Analysis $fractionalThemeCount } '*nonnegative native integer*' 'Fractional theme count violates evidence contract'

    $negativeThemeCount = Copy-TestJsonObject $analysisJson
    $negativeThemeCount.theme_evidence[0].do_not_contact_count = -1
    Assert-ThrowsLike { Test-GssAnalysisEvidenceContract -Analysis $negativeThemeCount } '*nonnegative native integer*' 'Negative theme count violates evidence contract'
    Assert-Equal $analysisJson.restaurants[0].name 'Richmond' 'Restaurant display name omits source identifier'
    Assert-Equal @($analysisJson.portfolio_evidence).Count 1 'Portfolio reporting basis evidence exists without a portfolio finding'
    Assert-Equal $analysisJson.portfolio_evidence[0].restaurant_id 'portfolio' 'Portfolio evidence uses the cross-repo restaurant identifier'
    Assert-True (-not [string]::IsNullOrWhiteSpace([string]$analysisJson.portfolio_evidence[0].display_text)) 'Portfolio evidence has deterministic display text'
    Assert-True (@($manifest.evidence_ids) -contains $analysisJson.portfolio_evidence[0].evidence_id) 'Manifest includes portfolio evidence ID'
    Assert-True (@($analysisJson.evidence_cards.evidence_id) -contains $analysisJson.portfolio_evidence[0].evidence_id) 'Unified evidence cards include portfolio evidence'
    Assert-True ([bool]$analysisJson.privacy.pii_scan_passed) 'PII scan passed'
    Assert-True ([bool]$analysisJson.privacy.guest_name_fields_excluded) 'Guest name fields excluded'
    Assert-True ([bool]$analysisJson.privacy.package_contains_personal_data) 'Privacy metadata does not claim the package is PII-free'
    Assert-Equal $analysisJson.privacy.analysis_description 'risk-reduced' 'Privacy metadata uses risk-reduced wording'
    foreach ($source in @($manifest.sources)) {
        Assert-True (-not ([string]$source.path).Contains(':')) "Portable source path for $($source.role)"
        Assert-True ([string]$source.sha256 -match '^[a-f0-9]{64}$') "Source SHA-256 for $($source.role)"
    }
    foreach ($attachment in @($manifest.attachments)) {
        $attachmentPath = Join-Path $package.PackagePath ([string]$attachment.path).Replace('/', '\')
        Assert-Equal (Get-GssSha256 $attachmentPath) ([string]$attachment.sha256) "Attachment hash for $($attachment.role)"
        Assert-Equal ([long](Get-Item -LiteralPath $attachmentPath).Length) ([long]$attachment.byte_size) "Attachment size for $($attachment.role)"
    }
    $validationDescriptors = @($manifest.sources | ForEach-Object {
        [pscustomobject]@{
            role = $_.role
            source_path = $_.path
            byte_size = $_.byte_size
            sha256 = $_.sha256
        }
    })

    $missingPortableArtifactManifest = Copy-TestJsonObject $manifest
    $missingPortableArtifactManifest.portable_artifacts = @($missingPortableArtifactManifest.portable_artifacts | Select-Object -First 5)
    Write-GssUtf8NoBomFile -Path $package.ManifestPath -Value ($missingPortableArtifactManifest | ConvertTo-Json -Depth 20)
    Assert-ThrowsLike {
        Test-GssExistingEmailPackage -PackagePath $package.PackagePath -PackageId $package.PackageId -ExpectedSourceDescriptors $validationDescriptors -ExpectedFeedbackSelectionFingerprint $manifest.feedback_selection_sha256
    } '*exactly 6 portable artifact records*' 'Existing package rejects an incomplete portable-artifact inventory'

    $duplicatePortableArtifactManifest = Copy-TestJsonObject $manifest
    $duplicatePortableArtifactManifest.portable_artifacts[-1].role = 'analysis_json'
    Write-GssUtf8NoBomFile -Path $package.ManifestPath -Value ($duplicatePortableArtifactManifest | ConvertTo-Json -Depth 20)
    Assert-ThrowsLike {
        Test-GssExistingEmailPackage -PackagePath $package.PackagePath -PackageId $package.PackageId -ExpectedSourceDescriptors $validationDescriptors -ExpectedFeedbackSelectionFingerprint $manifest.feedback_selection_sha256
    } "*exactly one portable artifact for role 'analysis_json'*" 'Existing package rejects duplicate portable-artifact roles'

    $traversalPortableArtifactManifest = Copy-TestJsonObject $manifest
    ($traversalPortableArtifactManifest.portable_artifacts | Where-Object role -eq 'analysis_json').path = '../analysis.json'
    Write-GssUtf8NoBomFile -Path $package.ManifestPath -Value ($traversalPortableArtifactManifest | ConvertTo-Json -Depth 20)
    Assert-ThrowsLike {
        Test-GssExistingEmailPackage -PackagePath $package.PackagePath -PackageId $package.PackageId -ExpectedSourceDescriptors $validationDescriptors -ExpectedFeedbackSelectionFingerprint $manifest.feedback_selection_sha256
    } "*portable artifact path is invalid for role 'analysis_json'*" 'Existing package rejects portable-artifact path traversal'

    Write-GssUtf8NoBomFile -Path $package.ManifestPath -Value ($manifest | ConvertTo-Json -Depth 20)
    foreach ($artifact in @($manifest.portable_artifacts)) {
        $artifactPath = Join-Path $package.PackagePath ([string]$artifact.path).Replace('/', '\')
        $originalArtifactBytes = [System.IO.File]::ReadAllBytes($artifactPath)
        try {
            [System.IO.File]::AppendAllText($artifactPath, 'tampered')
            Assert-ThrowsLike {
                Test-GssExistingEmailPackage -PackagePath $package.PackagePath -PackageId $package.PackageId -ExpectedSourceDescriptors $validationDescriptors -ExpectedFeedbackSelectionFingerprint $manifest.feedback_selection_sha256
            } "*portable artifact does not match its manifest: $($artifact.role)*" "Existing package rejects mutated portable artifact $($artifact.role)"
        }
        finally {
            [System.IO.File]::WriteAllBytes($artifactPath, $originalArtifactBytes)
        }
    }

    $analysisPathForEqualityTest = Join-Path $package.PackagePath 'analysis.json'
    $originalAnalysisBytes = [System.IO.File]::ReadAllBytes($analysisPathForEqualityTest)
    try {
        $mismatchedAnalysis = Copy-TestJsonObject $analysisJson
        $mismatchedAnalysis.commenter_lens.reason = 'A different but otherwise schema-valid standalone reason.'
        Write-GssUtf8NoBomFile -Path $analysisPathForEqualityTest -Value ($mismatchedAnalysis | ConvertTo-Json -Depth 20)
        $mismatchedManifest = Copy-TestJsonObject $manifest
        $analysisArtifactRecord = $mismatchedManifest.portable_artifacts | Where-Object role -eq 'analysis_json'
        $analysisArtifactRecord.byte_size = [long](Get-Item -LiteralPath $analysisPathForEqualityTest).Length
        $analysisArtifactRecord.sha256 = Get-GssSha256 $analysisPathForEqualityTest
        Write-GssUtf8NoBomFile -Path $package.ManifestPath -Value ($mismatchedManifest | ConvertTo-Json -Depth 20)
        Assert-ThrowsLike {
            Test-GssExistingEmailPackage -PackagePath $package.PackagePath -PackageId $package.PackageId -ExpectedSourceDescriptors $validationDescriptors -ExpectedFeedbackSelectionFingerprint $manifest.feedback_selection_sha256
        } '*analysis and commenter-lens output disagree*' 'Existing package requires full embedded and standalone commenter-lens equality'
    }
    finally {
        [System.IO.File]::WriteAllBytes($analysisPathForEqualityTest, $originalAnalysisBytes)
        Write-GssUtf8NoBomFile -Path $package.ManifestPath -Value ($manifest | ConvertTo-Json -Depth 20)
    }

    $missingAttachmentManifest = Copy-TestJsonObject $manifest
    $missingAttachmentManifest.attachments = @($missingAttachmentManifest.attachments | Select-Object -First 2)
    Write-GssUtf8NoBomFile -Path $package.ManifestPath -Value ($missingAttachmentManifest | ConvertTo-Json -Depth 20)
    Assert-ThrowsLike {
        Test-GssExistingEmailPackage -PackagePath $package.PackagePath -PackageId $package.PackageId -ExpectedSourceDescriptors $validationDescriptors -ExpectedFeedbackSelectionFingerprint $manifest.feedback_selection_sha256
    } '*exactly three attachments*' 'Existing package rejects an attachment count other than three'

    $duplicateAttachmentRoleManifest = Copy-TestJsonObject $manifest
    $duplicateAttachmentRoleManifest.attachments[2].role = 'comparison_pdf'
    Write-GssUtf8NoBomFile -Path $package.ManifestPath -Value ($duplicateAttachmentRoleManifest | ConvertTo-Json -Depth 20)
    Assert-ThrowsLike {
        Test-GssExistingEmailPackage -PackagePath $package.PackagePath -PackageId $package.PackageId -ExpectedSourceDescriptors $validationDescriptors -ExpectedFeedbackSelectionFingerprint $manifest.feedback_selection_sha256
    } "*exactly one attachment for role 'comparison_pdf'*" 'Existing package rejects duplicate or missing required attachment roles'

    Write-GssUtf8NoBomFile -Path $package.ManifestPath -Value ($manifest | ConvertTo-Json -Depth 20)
    Assert-True ([bool](Test-GssExistingEmailPackage -PackagePath $package.PackagePath -PackageId $package.PackageId -ExpectedSourceDescriptors $validationDescriptors -ExpectedFeedbackSelectionFingerprint $manifest.feedback_selection_sha256)) 'Original three-attachment manifest remains valid'

    $portableTextByFile = [ordered]@{
        'email_manifest.json' = Get-Content -Raw -LiteralPath $package.ManifestPath
        'analysis.json' = Get-Content -Raw -LiteralPath (Join-Path $package.PackagePath 'analysis.json')
        'commenter_lens.json' = Get-Content -Raw -LiteralPath $package.CommenterLensJsonPath
        'commenter_lens.csv' = Get-Content -Raw -LiteralPath $package.CommenterLensCsvPath
        'email_preview.txt' = Get-Content -Raw -LiteralPath (Join-Path $package.PackagePath 'email_preview.txt')
        'email_preview.html' = Get-Content -Raw -LiteralPath (Join-Path $package.PackagePath 'email_preview.html')
        'RESTRICTED.txt' = Get-Content -Raw -LiteralPath (Join-Path $package.PackagePath 'RESTRICTED.txt')
    }
    $nonRaw = @($portableTextByFile.Values) -join ''
    Assert-True ($nonRaw -match '(?i)exact seven-day sample') 'Required methodology wording is not mistaken for a matching guest surname'
    Assert-True ($nonRaw -match '(?i)adjacent 13-week windows overlap by 12 weeks') 'Portable narrative discloses adjacent-window overlap'
    Assert-True ($nonRaw -match '(?i)theme categories are non-exclusive') 'Portable narrative discloses non-exclusive themes'
    Assert-True ($nonRaw -match '(?i)derived non-comment counts and rates are suppressed') 'Portable narrative discloses the partition-alignment gate'
    Assert-True ($nonRaw -match '(?i)response-volume tier') 'Portable narrative uses response-volume tier semantics'
    Assert-True ($nonRaw -match '(?i)risk-reduced') 'Portable narrative uses risk-reduced wording'
    Assert-True ($nonRaw -notmatch '(?i)\banonym(?:ous|ized|ised)\b') 'Portable narrative makes no anonymity claim'
    foreach ($forbidden in @('Casey', 'Testperson', 'Robin', 'Sample', 'casey@example.invalid', '212-555-0199', '555-1212', 'https://example.invalid', 'bare.example.invalid', 'ABC123', 'ZX9876')) {
        $leakedFiles = @($portableTextByFile.GetEnumerator() | Where-Object { ([string]$_.Value).Contains($forbidden) } | ForEach-Object { $_.Key })
        Assert-True ($leakedFiles.Count -eq 0) "PII canary removed: $forbidden; leaked files: $($leakedFiles -join ', ')"
    }
    Assert-True (-not [regex]::IsMatch($nonRaw, (Get-GssUnsafeControlPattern))) 'Portable output contains no unsafe control or bidi characters'
    $ledgerText = Get-Content -Raw -LiteralPath $ledgerPath
    Assert-True ($ledgerText -notmatch '(?i)Casey|Testperson|Robin|Sample|example\.invalid|212-555-0199|555-1212') 'Ledger remains hash-only'
    $ledger = $ledgerText | ConvertFrom-Json
    Assert-Equal @($ledger.entries).Count 3 'Ledger includes deduplicated current and baseline hashes'
    Assert-True ($initialLedgerWriteTimeUtc -le $initialReadyWriteTimeUtc) 'Ledger is persisted before the READY marker'

    $sameWeekSelection = Get-GssFeedbackSelection -Inventory $inventory -Ledger $ledger -ReportingDate ([datetime]'2026-07-12')
    Assert-Equal $sameWeekSelection.Fingerprint $manifest.feedback_selection_sha256 'Same-week rerun retains the selected-feedback fingerprint after ledger persistence'
    $priorSeenLedger = ($ledger | ConvertTo-Json -Depth 8 | ConvertFrom-Json)
    $contactableHash = [string]($inventory.UniqueResponses | Where-Object Text -match 'praised the service' | Select-Object -First 1 -ExpandProperty ResponseHash)
    $priorSeenEntry = $priorSeenLedger.entries | Where-Object response_hash -eq $contactableHash | Select-Object -First 1
    $priorSeenEntry.first_seen_reporting_date = '2026-07-05'
    $priorSeenSelection = Get-GssFeedbackSelection -Inventory $inventory -Ledger $priorSeenLedger -ReportingDate ([datetime]'2026-07-12')
    Assert-True ($priorSeenSelection.Fingerprint -ne $sameWeekSelection.Fingerprint) 'A response first seen before this report changes the selected-feedback fingerprint'

    $futureDetail = Join-Path $detailFolder 'Sorensen Future.xlsx'
    $future = New-TestResponse '9355 Virginia Beach' '07/20/2026' '8:00 PM' 'Future visit should wait for its reporting cutoff.' 'Future' 'Guest'
    New-TestDetailWorkbook -Path $futureDetail -Headers $headers19 -Records @($future)
    $futureInventory = Get-GssDetailInventory -FolderPath $folder -ReportingDate ([datetime]'2026-07-12')
    $futureFeedback = New-GssSanitizedFeedback -Inventory $futureInventory -Ledger $ledger -PackageId 'future-cutoff-check' -ReportingDate ([datetime]'2026-07-12')
    Assert-Equal $futureFeedback.ResponseCount 2 'Future response is excluded from current reporting selection'
    Assert-Equal @($futureFeedback.NextLedger.entries).Count 3 'Future response is not written to first-seen ledger early'
    Remove-Item -LiteralPath $futureDetail -Force

    (Get-Item -LiteralPath $rolling).LastWriteTime = (Get-Date).AddYears(-5)
    $second = New-GssEmailPackage -FolderPath $folder -RunLog $runLog -AnalysisResult $analysis -LedgerPath $ledgerPath
    Assert-Equal $second.PackageId $package.PackageId 'Package ID ignores modification times and machine root'
    Assert-True ([bool]$second.ExistingPackage) 'Identical rerun reuses one validated package'
    Assert-Equal @(Get-ChildItem -LiteralPath (Split-Path -Parent $package.PackagePath) -Directory | Where-Object { $_.Name -notlike '.staging-*' }).Count 1 'No duplicate package directory'

    'incomplete' | Set-Content -LiteralPath $package.ReadyMarkerPath -Encoding ASCII
    Assert-ThrowsLike { New-GssEmailPackage -FolderPath $folder -RunLog $runLog -AnalysisResult $analysis -LedgerPath $ledgerPath } '*READY marker does not match package ID*' 'Existing package requires exact READY marker contents'
    "$($manifest.schema_version)`n$($manifest.package_id)" | Set-Content -LiteralPath $package.ReadyMarkerPath -Encoding ASCII

    $descriptors = @($manifest.sources | ForEach-Object { [pscustomobject]@{ role = $_.role; source_path = $_.path; byte_size = $_.byte_size; sha256 = $_.sha256 } })
    $sameId = Get-GssDeterministicPackageId -ReportingDate '2026-07-12' -SourceDescriptors $descriptors -FeedbackSelectionFingerprint $manifest.feedback_selection_sha256
    Assert-Equal $sameId $package.PackageId 'Deterministic package ID snapshot'
    $legacyPackageId = Get-GssDeterministicPackageId -ReportingDate '2026-07-12' -SourceDescriptors $descriptors -FeedbackSelectionFingerprint $manifest.feedback_selection_sha256 -SchemaVersion 'gss-email-package/v1'
    Assert-True ($legacyPackageId -ne $package.PackageId) 'Schema v2 uses a new deterministic package ID and cannot collide with legacy v1 output'
    $changedDescriptors = @($descriptors | ForEach-Object { [pscustomobject]@{ role = $_.role; source_path = $_.source_path; byte_size = $_.byte_size; sha256 = $_.sha256 } })
    $changedDescriptors[0].sha256 = ('0' * 64)
    Assert-True ((Get-GssDeterministicPackageId -ReportingDate '2026-07-12' -SourceDescriptors $changedDescriptors -FeedbackSelectionFingerprint $manifest.feedback_selection_sha256) -ne $package.PackageId) 'Source hash changes package ID'
    $changedArchiveDescriptors = @($descriptors | ForEach-Object { [pscustomobject]@{ role = $_.role; source_path = $_.source_path; byte_size = $_.byte_size; sha256 = $_.sha256 } })
    ($changedArchiveDescriptors | Where-Object role -eq 'detail_archive_workbook').sha256 = ('1' * 64)
    Assert-True ((Get-GssDeterministicPackageId -ReportingDate '2026-07-12' -SourceDescriptors $changedArchiveDescriptors -FeedbackSelectionFingerprint $manifest.feedback_selection_sha256) -ne $package.PackageId) 'Archive-detail hash changes package ID'
    Assert-True ((Get-GssDeterministicPackageId -ReportingDate '2026-07-12' -SourceDescriptors $descriptors -FeedbackSelectionFingerprint $priorSeenSelection.Fingerprint) -ne $package.PackageId) 'Selected-feedback change changes package ID'

    $mixedCaseSources = @(
        [pscustomobject]@{ role = 'detail_archive_workbook'; source_path = 'Archive/z.xlsx'; byte_size = 10; sha256 = ('a' * 64) },
        [pscustomobject]@{ role = 'detail_archive_workbook'; source_path = 'archive/A.xlsx'; byte_size = 11; sha256 = ('b' * 64) },
        [pscustomobject]@{ role = 'run_log'; source_path = '_automation_runs/Logs/run.json'; byte_size = 12; sha256 = ('c' * 64) }
    )
    $mixedCaseId = Get-GssDeterministicPackageId -ReportingDate '2026-07-12' -SourceDescriptors $mixedCaseSources -FeedbackSelectionFingerprint $manifest.feedback_selection_sha256
    [array]::Reverse($mixedCaseSources)
    Assert-Equal (Get-GssDeterministicPackageId -ReportingDate '2026-07-12' -SourceDescriptors $mixedCaseSources -FeedbackSelectionFingerprint $manifest.feedback_selection_sha256) $mixedCaseId 'Mixed-case source paths use stable case-insensitive ordering'

    [System.IO.File]::AppendAllText($pdf, 'changed')
    Assert-ThrowsLike { New-GssEmailPackage -FolderPath $folder -RunLog $runLog -AnalysisResult $analysis -LedgerPath $ledgerPath } '*File changed after the successful workbook run: comparison_pdf*' 'Attachment mutation blocks package readiness'
    Assert-Equal @(Get-ChildItem -LiteralPath (Split-Path -Parent $package.PackagePath) -Directory | Where-Object { $_.Name -notlike '.staging-*' }).Count 1 'Blocked mutation creates no ready package'

    $emptyFile = Join-Path $folder '04 Email Comparison PDFs\empty.pdf'
    [System.IO.File]::WriteAllBytes($emptyFile, [byte[]]@())
    Assert-ThrowsLike { Get-GssSourceDescriptor -Role 'comparison_pdf' -Path $emptyFile -FolderPath $folder } '*empty or not fully synced*' 'Zero-byte partial sync blocks readiness'

    Write-Host 'GSS email package tests passed.'
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot -PathType Container) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}
