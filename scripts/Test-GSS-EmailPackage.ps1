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

function Assert-ThrowsLike {
    param([scriptblock]$Script, [string]$Pattern, [string]$Name)
    try { & $Script; throw "Assertion failed for $Name. Expected an exception." }
    catch {
        if ($_.Exception.Message -notlike $Pattern) {
            throw "Assertion failed for $Name. Expected '$Pattern'; actual '$($_.Exception.Message)'."
        }
    }
}

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

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('gss_email_package_test_' + [guid]::NewGuid().ToString('N'))
$folder = Join-Path $temporaryRoot 'GSS Surveys'
try {
    $detailFolder = Join-Path $folder '03 Uploaded Survey Workbooks'
    $archiveFolder = Join-Path $detailFolder 'Archive - Previous Uploads'
    $headers18 = @('Restaurant Name', 'Reservation Date', 'Reservation Time', 'Text', 'Overall', 'Service', 'Culinary', 'Value', 'Pace of Meal', 'Recommend', 'Manager Visit', 'Steak Cooked Correctly', 'Event Booking Process', 'First Visit', 'Guest First Name', 'Guest Last Name', 'Sorensen', 'Sorensen Weekly Comments')
    $headers19 = @('Text', 'Restaurant Name', 'Reservation Time', 'Reservation Date', 'Service', 'Overall', 'Culinary', 'Value', 'Pace of Meal', 'Recommend', 'Manager Visit', 'Steak Cooked Correctly', 'Alert Guests DO NOT CONTACT', 'Event Booking Process', 'First Visit', 'Guest Last Name', 'Guest First Name', 'Sorensen Weekly Comments', 'Sorensen')
    $unsafeBidi = [char]0x202E
    $unsafeC0 = [char]0x0001
    $contactableText = "Casey Testperson praised the service. Contact casey@example.invalid or 212-555-0199 and 555-1212; https://example.invalid and bare.example.invalid/path, Reservation ABC123, Resy ZX9876. Unsafe ${unsafeBidi}bidi control."
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
    foreach ($phoneProbe in @('212-555-0199', '555-1212', '(212) 555-0199', '212.555.0199', '2125550199', '+44 20 7946 0958')) {
        Assert-True ([regex]::IsMatch($phoneProbe, $phonePattern)) "Phone pattern recognizes $phoneProbe"
    }
    foreach ($nonPhoneProbe in @('78.0952380952381', '2026-07-12', 'metric-9354-service')) {
        Assert-True (-not [regex]::IsMatch($nonPhoneProbe, $phonePattern)) "Phone pattern rejects non-phone value $nonPhoneProbe"
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

    $inventory = Get-GssDetailInventory -FolderPath $folder -ReportingDate ([datetime]'2026-07-12')
    Assert-Equal $inventory.CurrentWorkbook.PortablePath '03 Uploaded Survey Workbooks/Sorensen Current.xlsx' 'Current detail selection by visit date'
    Assert-Equal $inventory.UniqueResponses.Count 3 'Overlapping exports deduplicate responses'
    Assert-Equal $inventory.DuplicateResponseCount 1 'Duplicate response count'
    $knownGuestNames = @(Get-GssKnownGuestNames $inventory.AllResponseInstances)
    Assert-True ($knownGuestNames -contains 'Casey' -and $knownGuestNames -contains 'Robin') 'Distinct same-length guest names remain in the redaction set'

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
    $analysis = [pscustomobject]@{
        WorkbookStatus = 'Ready'
        AnalysisStatus = 'Review'
        EmailReadiness = 'Ready'
        LogPath = $logPath
        MetricDetail = @($finding)
        RestaurantFindings = @(Select-GssRestaurantFindings @($finding))
    }
    $ledgerPath = Join-Path $folder '_automation_runs\state\fixture_ledger.json'
    $package = New-GssEmailPackage -FolderPath $folder -RunLog $runLog -AnalysisResult $analysis -LedgerPath $ledgerPath
    Assert-Equal $package.EmailReadiness 'Ready' 'Package email readiness'
    Assert-True (Test-Path -LiteralPath $package.ReadyMarkerPath -PathType Leaf) 'Ready marker exists'
    $manifest = Get-Content -Raw -LiteralPath $package.ManifestPath | ConvertFrom-Json
    $analysisJson = Get-Content -Raw -LiteralPath (Join-Path $package.PackagePath 'analysis.json') | ConvertFrom-Json
    Assert-Equal $manifest.schema_version 'gss-email-package/v1' 'Package schema version'
    Assert-True ([string]$manifest.feedback_selection_sha256 -match '^[a-f0-9]{64}$') 'Manifest carries the selected-feedback fingerprint'
    Assert-Equal $analysisJson.feedback_selection_sha256 $manifest.feedback_selection_sha256 'Analysis and manifest feedback selection agree'
    Assert-Equal @($manifest.attachments).Count 3 'Three attachment policy'
    Assert-Equal @($manifest.sources).Count 7 'Portable source evidence count including run log and archive detail'
    Assert-Equal @($manifest.sources | Where-Object role -eq 'run_log').Count 1 'Exact run log source evidence'
    Assert-Equal @($manifest.sources | Where-Object role -eq 'detail_archive_workbook').Count 1 'Archived detail source evidence'
    Assert-Equal $manifest.source_log_path ($manifest.sources | Where-Object role -eq 'run_log' | Select-Object -ExpandProperty path) 'Run-log path matches source descriptor'
    Assert-Equal $analysisJson.methodology.feedback_response_count 2 'Initial package includes only current-file new feedback'
    Assert-Equal $analysisJson.methodology.feedback_visit_date_start '2026-07-10' 'Actual feedback visit-date start'
    Assert-Equal $analysisJson.methodology.feedback_visit_date_end '2026-07-11' 'Actual feedback visit-date end'
    Assert-Equal @($analysisJson.sanitized_feedback).Count 1 'DNC response excluded from individual evidence'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match '\[REDACTED URL\],') 'Bare URL is redacted without swallowing sentence punctuation'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match '\[REDACTED PHONE\]') 'Short local phone number is redacted'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match '\[REDACTED BOOKING ID\]') 'Reservation and Resy identifiers are redacted'
    Assert-True ([string]$analysisJson.sanitized_feedback[0].sanitized_text -match '\[REDACTED CONTROL\]') 'Unsafe C0 and bidi controls are visibly removed'
    $serviceTheme = $analysisJson.feedback_themes | Where-Object category -eq 'service'
    Assert-Equal $serviceTheme.unique_response_count 2 'Theme requires two unique responses'
    Assert-Equal $serviceTheme.do_not_contact_count 1 'DNC contributes only to anonymous theme counts'
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
    Assert-Equal ([int]$serviceThemeEvidence.concern_count) 0 'Theme evidence concern-count contract'
    Assert-Equal ([int]$serviceThemeEvidence.positive_count) 2 'Theme evidence positive-count contract'
    Assert-Equal ([int]$serviceThemeEvidence.do_not_contact_count) 1 'Theme evidence do-not-contact-count contract'
    Assert-True (Test-GssAnalysisEvidenceContract -Analysis $analysisJson) 'Serialized analysis satisfies structured evidence contract'

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
    foreach ($source in @($manifest.sources)) {
        Assert-True (-not ([string]$source.path).Contains(':')) "Portable source path for $($source.role)"
        Assert-True ([string]$source.sha256 -match '^[a-f0-9]{64}$') "Source SHA-256 for $($source.role)"
    }
    foreach ($attachment in @($manifest.attachments)) {
        $attachmentPath = Join-Path $package.PackagePath ([string]$attachment.path).Replace('/', '\')
        Assert-Equal (Get-GssSha256 $attachmentPath) ([string]$attachment.sha256) "Attachment hash for $($attachment.role)"
        Assert-Equal ([long](Get-Item -LiteralPath $attachmentPath).Length) ([long]$attachment.byte_size) "Attachment size for $($attachment.role)"
    }
    $portableTextByFile = [ordered]@{
        'email_manifest.json' = Get-Content -Raw -LiteralPath $package.ManifestPath
        'analysis.json' = Get-Content -Raw -LiteralPath (Join-Path $package.PackagePath 'analysis.json')
        'email_preview.txt' = Get-Content -Raw -LiteralPath (Join-Path $package.PackagePath 'email_preview.txt')
        'email_preview.html' = Get-Content -Raw -LiteralPath (Join-Path $package.PackagePath 'email_preview.html')
    }
    $nonRaw = @($portableTextByFile.Values) -join ''
    Assert-True ($nonRaw -match '(?i)exact seven-day sample') 'Required methodology wording is not mistaken for a matching guest surname'
    foreach ($forbidden in @('Casey', 'Testperson', 'Robin', 'Sample', 'casey@example.invalid', '212-555-0199', '555-1212', 'https://example.invalid', 'bare.example.invalid', 'ABC123', 'ZX9876')) {
        $leakedFiles = @($portableTextByFile.GetEnumerator() | Where-Object { ([string]$_.Value).Contains($forbidden) } | ForEach-Object { $_.Key })
        Assert-True ($leakedFiles.Count -eq 0) "PII canary removed: $forbidden; leaked files: $($leakedFiles -join ', ')"
    }
    Assert-True (-not [regex]::IsMatch($nonRaw, (Get-GssUnsafeControlPattern))) 'Portable output contains no unsafe control or bidi characters'
    $ledgerText = Get-Content -Raw -LiteralPath $ledgerPath
    Assert-True ($ledgerText -notmatch '(?i)Casey|Testperson|Robin|Sample|example\.invalid|555') 'Ledger remains hash-only'
    $ledger = $ledgerText | ConvertFrom-Json
    Assert-Equal @($ledger.entries).Count 3 'Ledger includes deduplicated current and baseline hashes'
    Assert-True ((Get-Item -LiteralPath $ledgerPath).LastWriteTimeUtc -le (Get-Item -LiteralPath $package.ReadyMarkerPath).LastWriteTimeUtc) 'Ledger is persisted before the READY marker'

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
