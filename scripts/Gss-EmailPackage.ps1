$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
if (-not (Get-Command ConvertTo-GssDropboxRelativePath -ErrorAction SilentlyContinue)) {
    . (Join-Path $scriptRoot 'Gss-Common.ps1')
}

$script:GssEmailPackageSchemaVersion = 'gss-email-package/v2'
$script:GssFeedbackLedgerVersion = 'gss-feedback-first-seen/v1'
$script:GssHistoricalRecoveryManifestVersion = 'gss-historical-recovery/v1'
$script:GssHistoricalRecoveryReceiptVersion = 'gss-historical-recovery-receipt/v1'
$script:GssHistoricalResponseSetVersion = 'gss-historical-response-set/v1'
$script:GssHistoricalRecoveryArchivePrefix = '03 Uploaded Survey Workbooks/Archive - Previous Uploads/Recovered Historical Detail'
$script:GssTransactionMutexName = 'Global\GSSSurveyWorkbookAutomationTransaction'
$script:GssThemeNames = @('service', 'culinary', 'pace', 'value', 'hospitality/recovery', 'recognition')
$script:GssRestrictedClassification = 'CONTAINS PERSONAL DATA ' + [char]0x2014 + ' RESTRICTED'
$script:GssPortableArtifactPaths = [ordered]@{
    analysis_json = 'analysis.json'
    commenter_lens_json = 'commenter_lens.json'
    commenter_lens_csv = 'commenter_lens.csv'
    email_preview_text = 'email_preview.txt'
    email_preview_html = 'email_preview.html'
    classification_notice = 'RESTRICTED.txt'
}
$script:GssCommenterLensCsvColumns = @(
    'status',
    'scope_label',
    'reporting_window_start',
    'reporting_window_end',
    'exact_partition_alignment_verified',
    'restaurant_id',
    'restaurant_status',
    'population_response_count',
    'commenter_response_count',
    'comment_coverage_pct',
    'comment_coverage_status',
    'metric_id',
    'response_field',
    'population_metric',
    'commenter_scored_response_count',
    'commenter_missing_score_count',
    'commenter_event_count',
    'commenter_event_rate_pct',
    'commenter_rate_denominator_label',
    'population_event_rate_pct',
    'population_rate_denominator_label',
    'denominator_alignment_status',
    'commenter_minus_population_percentage_points',
    'material_gap',
    'reconstructed_population_event_count',
    'derived_non_comment_response_count',
    'derived_non_comment_event_count',
    'derived_non_comment_event_rate_pct',
    'commenter_minus_non_comment_percentage_points',
    'comparison_status',
    'derived_non_comment_status'
)

function Read-GssAnalysisPolicy {
    $policyPath = Join-Path (Split-Path -Parent $scriptRoot) 'config\analysis-policy.json'
    if (-not (Test-Path -LiteralPath $policyPath -PathType Leaf)) {
        throw "GSS analysis policy is missing: $policyPath"
    }

    try {
        $policy = Get-Content -Raw -LiteralPath $policyPath | ConvertFrom-Json
    }
    catch {
        throw "GSS analysis policy is not valid JSON: $policyPath. $($_.Exception.Message)"
    }

    if ([string]$policy.schema_version -notmatch '^gss-analysis-policy/v\d+$') {
        throw "GSS analysis policy has an invalid schema version: $($policy.schema_version)"
    }
    if ([double]$policy.thresholds.previous_window.candidate_points -ne 1 -or
        [double]$policy.thresholds.previous_window.action_points -ne 2 -or
        [double]$policy.thresholds.prior_year.candidate_points -ne 2 -or
        [double]$policy.thresholds.prior_year.action_points -ne 5) {
        throw 'GSS analysis policy must preserve the approved 1/2-point previous-window and 2/5-point prior-year thresholds.'
    }
    if ([int]$policy.confidence.high_minimum_responses -ne 100 -or
        [int]$policy.confidence.developing_minimum_responses -ne 50 -or
        [int]$policy.confidence.low_minimum_responses -ne 1 -or
        -not [bool]$policy.confidence.developing_requires_threshold_and_corroboration -or
        [bool]$policy.confidence.low_eligible_for_top_findings -or
        [bool]$policy.confidence.not_scored_eligible_for_top_findings) {
        throw 'GSS analysis policy must preserve the approved High, Developing, Low, and Not scored confidence boundaries.'
    }
    if ([int]$policy.rolling_windows.weeks -ne 13 -or
        [int]$policy.rolling_windows.adjacent_overlap_weeks -ne 12) {
        throw 'GSS analysis policy must identify 13-week rolling windows with 12 weeks of overlap between adjacent windows.'
    }
    if ([int]$policy.guest_feedback.minimum_unique_responses_per_theme -ne 2 -or
        -not [bool]$policy.guest_feedback.theme_categories_are_non_exclusive) {
        throw 'GSS analysis policy must preserve the approved two-response theme minimum and non-exclusive theme categories.'
    }
    if (-not [bool]$policy.review_controls.human_review_required -or
        [bool]$policy.review_controls.automatic_sending_enabled -or
        [bool]$policy.review_controls.causation_claimed -or
        [bool]$policy.review_controls.statistical_significance_claimed) {
        throw 'GSS analysis policy must require human review, keep automatic sending disabled, and make no causation or statistical-significance claim.'
    }
    $populationRawRowsProperty = $policy.source_design.PSObject.Properties['population_raw_rows_available']
    if ($null -eq $populationRawRowsProperty -or
        $populationRawRowsProperty.Value -isnot [bool] -or
        [bool]$populationRawRowsProperty.Value -or
        [string]$policy.source_design.population_scores_source -ne 'rolling_aggregate_workbook' -or
        [string]$policy.source_design.row_level_scores_source -ne 'surveys_with_comments_only' -or
        [bool]$policy.source_design.non_comment_row_level_scores_available -or
        [bool]$policy.source_design.commenter_rows_population_representative -or
        [bool]$policy.source_design.population_driver_modeling_supported -or
        [string]$policy.source_design.population_driver_modeling_status -ne 'PopulationRawDataUnavailable') {
        throw 'GSS analysis policy must identify aggregate population scores, commenter-only row scores, and unavailable population driver modeling.'
    }
    if (-not [bool]$policy.commenter_lens.enabled -or
        [string]$policy.commenter_lens.scope_label -ne 'Among guests who provided comments' -or
        [int]$policy.commenter_lens.reporting_window_weeks -ne 13 -or
        [double]$policy.commenter_lens.material_gap_percentage_points -lt 0 -or
        [double]$policy.commenter_lens.maximum_reconstructed_event_count_error -lt 0 -or
        [bool]$policy.commenter_lens.population_prevalence_claim_allowed -or
        [bool]$policy.commenter_lens.statistical_significance_claim_allowed -or
        [bool]$policy.commenter_lens.individual_prediction_allowed) {
        throw 'GSS analysis policy must keep the commenter lens descriptive, scoped to comment-providing guests, and free of population prevalence, significance, and individual-prediction claims.'
    }
    $commenterMetricIds = @()
    foreach ($metric in @($policy.commenter_lens.metrics)) {
        if ([string]::IsNullOrWhiteSpace([string]$metric.id) -or
            [string]::IsNullOrWhiteSpace([string]$metric.response_field) -or
            $null -eq $metric.event_minimum -or
            $null -eq $metric.event_maximum -or
            [double]$metric.event_minimum -gt [double]$metric.event_maximum) {
            throw 'Every commenter-lens metric must define an ID, response field, exact population metric, and valid inclusive event range.'
        }
        $commenterMetricIds += [string]$metric.id
    }
    if ($commenterMetricIds.Count -eq 0 -or
        @($commenterMetricIds | Group-Object | Where-Object Count -gt 1).Count -gt 0 -or
        'commenter_lens.json' -notin @($policy.commenter_lens.outputs) -or
        'commenter_lens.csv' -notin @($policy.commenter_lens.outputs)) {
        throw 'GSS analysis policy must define distinct commenter-lens metrics and the aggregate JSON and CSV outputs.'
    }
    if ([string]$policy.modeling.status -ne 'PopulationRawDataUnavailable' -or
        [string]::IsNullOrWhiteSpace([string]$policy.modeling.reason) -or
        -not [bool]$policy.modeling.nonblocking -or
        [bool]$policy.modeling.blocks_workbook_or_backup_or_package -or
        [bool]$policy.modeling.email_attachment_allowed -or
        [bool]$policy.modeling.row_level_persistence_allowed -or
        [bool]$policy.commenter_lens.causation_claim_allowed -or
        [bool]$policy.commenter_lens.email_attachment_allowed -or
        [bool]$policy.commenter_lens.row_level_persistence_allowed) {
        throw 'GSS analysis policy must keep population modeling unavailable, nonblocking, unattached, and free of row-level persistence.'
    }

    return $policy
}

$script:GssAnalysisPolicy = Read-GssAnalysisPolicy
$script:GssAnalysisPolicyVersion = [string]$script:GssAnalysisPolicy.schema_version

function Get-GssStringSha256 {
    param([Parameter(Mandatory)][string]$Value)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
        return (($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    finally {
        $sha.Dispose()
    }
}

function Write-GssUtf8NoBomFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Value
    )

    $encoding = New-Object System.Text.UTF8Encoding($false, $true)
    [System.IO.File]::WriteAllText($Path, $Value + [Environment]::NewLine, $encoding)
}

function Read-GssUtf8NoBomFile {
    param([Parameter(Mandatory)][string]$Path)

    $encoding = New-Object System.Text.UTF8Encoding($false, $true)
    return [System.IO.File]::ReadAllText($Path, $encoding)
}

function Assert-GssExactObjectSchema {
    param(
        [Parameter(Mandatory)][AllowNull()][object]$Value,
        [Parameter(Mandatory)][string[]]$AllowedProperties,
        [Parameter(Mandatory)][string]$Label
    )

    if ($null -eq $Value) {
        throw "$Label must be a structured object."
    }
    $actualProperties = @($Value.PSObject.Properties | ForEach-Object { [string]$_.Name })
    $unexpected = @($actualProperties | Where-Object { $AllowedProperties -cnotcontains $_ })
    if ($unexpected.Count -gt 0) {
        throw "$Label contains unsupported property '$($unexpected[0])'."
    }
    $missing = @($AllowedProperties | Where-Object { $actualProperties -cnotcontains $_ })
    if ($missing.Count -gt 0) {
        throw "$Label is missing required property '$($missing[0])'."
    }
}

function Assert-GssScalarValue {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory)][string]$Label,
        [switch]$AllowNull
    )

    if ($null -eq $Value) {
        if ($AllowNull) { return }
        throw "$Label must not be null."
    }
    $baseValue = $Value.PSObject.BaseObject
    if ($baseValue -is [string] -or $baseValue -is [System.ValueType]) { return }
    throw "$Label must be a scalar value."
}

function Assert-GssUtf8NoBomFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Label
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        throw "Package output must be BOM-free UTF-8: $Label"
    }

    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    try {
        $null = $strictUtf8.GetString($bytes)
    }
    catch [System.Text.DecoderFallbackException] {
        throw "Package output must be valid UTF-8: $Label"
    }
}

function Get-GssZipEntryText {
    param([object]$Archive, [string]$EntryName)

    $entry = $Archive.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        try { return $reader.ReadToEnd() }
        finally { $reader.Dispose() }
    }
    finally { $stream.Dispose() }
}

function ConvertFrom-GssCellReference {
    param([string]$Reference)

    $letters = ([regex]::Match($Reference, '^[A-Z]+')).Value
    $column = 0
    foreach ($character in $letters.ToCharArray()) {
        $column = ($column * 26) + ([int][char]$character - [int][char]'A' + 1)
    }
    return ($column - 1)
}

function Normalize-GssFeedbackHeader {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return [regex]::Replace(([string]$Value).ToLowerInvariant(), '[^a-z0-9]+', '')
}

function Normalize-GssFeedbackText {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return ([regex]::Replace(([string]$Value).Trim().ToLowerInvariant(), '\s+', ' '))
}

function Get-GssRestaurantDisplayName {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return ([regex]::Replace($Value.Trim(), '^\d{4}\s+', '')).Trim()
}

function Read-GssXlsxFirstWorksheet {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw 'Detail workbook is missing.'
    }
    if ((Get-Item -LiteralPath $Path).Length -le 0) {
        throw 'Detail workbook is empty or not fully synced.'
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        [xml]$workbookXml = Get-GssZipEntryText $archive 'xl/workbook.xml'
        [xml]$relationshipXml = Get-GssZipEntryText $archive 'xl/_rels/workbook.xml.rels'
        if (-not $workbookXml -or -not $relationshipXml) {
            throw 'The workbook package is missing its workbook metadata.'
        }

        $workbookNs = New-Object System.Xml.XmlNamespaceManager($workbookXml.NameTable)
        $workbookNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $workbookNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $sheet = $workbookXml.SelectSingleNode('//m:sheets/m:sheet[1]', $workbookNs)
        if (-not $sheet) { throw 'The detail workbook has no worksheet.' }

        $relationshipMap = @{}
        foreach ($relationship in $relationshipXml.Relationships.Relationship) {
            $relationshipMap[[string]$relationship.Id] = [string]$relationship.Target
        }
        $relationshipId = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $target = $relationshipMap[$relationshipId]
        if ([string]::IsNullOrWhiteSpace($target)) { throw 'The first worksheet relationship is missing.' }
        $entryName = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { 'xl/' + $target.TrimStart('/') }

        $sharedStrings = @()
        $sharedText = Get-GssZipEntryText $archive 'xl/sharedStrings.xml'
        if ($sharedText) {
            [xml]$sharedXml = $sharedText
            $sharedNs = New-Object System.Xml.XmlNamespaceManager($sharedXml.NameTable)
            $sharedNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            foreach ($item in $sharedXml.SelectNodes('//m:si', $sharedNs)) {
                $parts = @($item.SelectNodes('.//m:t', $sharedNs) | ForEach-Object { $_.InnerText })
                $sharedStrings += ($parts -join '')
            }
        }

        [xml]$worksheetXml = Get-GssZipEntryText $archive $entryName
        if (-not $worksheetXml) { throw 'The first worksheet XML is missing.' }
        $worksheetNs = New-Object System.Xml.XmlNamespaceManager($worksheetXml.NameTable)
        $worksheetNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $rowNodes = @($worksheetXml.SelectNodes('//m:sheetData/m:row', $worksheetNs))
        if ($rowNodes.Count -lt 2) { throw 'The detail workbook has no response rows.' }

        $parsedRows = @()
        foreach ($rowNode in $rowNodes) {
            $values = @{}
            foreach ($cell in $rowNode.SelectNodes('./m:c', $worksheetNs)) {
                $columnIndex = ConvertFrom-GssCellReference ([string]$cell.r)
                $valueNode = $cell.SelectSingleNode('./m:v', $worksheetNs)
                $inlineNode = $cell.SelectSingleNode('./m:is', $worksheetNs)
                $value = ''
                if ([string]$cell.t -eq 's' -and $valueNode) {
                    $value = $sharedStrings[[int]$valueNode.InnerText]
                }
                elseif ([string]$cell.t -eq 'inlineStr' -and $inlineNode) {
                    $value = (@($inlineNode.SelectNodes('.//m:t', $worksheetNs) | ForEach-Object { $_.InnerText }) -join '')
                }
                elseif ($valueNode) {
                    $value = $valueNode.InnerText
                }
                $values[$columnIndex] = [string]$value
            }
            $parsedRows += ,$values
        }

        $headers = @()
        $headerMax = ($parsedRows[0].Keys | Measure-Object -Maximum).Maximum
        for ($columnIndex = 0; $columnIndex -le $headerMax; $columnIndex++) {
            $headers += if ($parsedRows[0].ContainsKey($columnIndex)) { [string]$parsedRows[0][$columnIndex] } else { '' }
        }
        $normalizedHeaders = @($headers | ForEach-Object { Normalize-GssFeedbackHeader $_ })
        $duplicateHeaders = @($normalizedHeaders | Where-Object { $_ } | Group-Object | Where-Object { $_.Count -gt 1 })
        if ($duplicateHeaders.Count -gt 0) {
            throw 'Detail workbook contains a duplicate normalized header.'
        }

        $records = @()
        for ($rowIndex = 1; $rowIndex -lt $parsedRows.Count; $rowIndex++) {
            $sourceRow = $parsedRows[$rowIndex]
            $record = [ordered]@{}
            $hasContent = $false
            for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
                $normalizedHeader = $normalizedHeaders[$columnIndex]
                if ([string]::IsNullOrWhiteSpace($normalizedHeader)) { continue }
                $cellValue = if ($sourceRow.ContainsKey($columnIndex)) { [string]$sourceRow[$columnIndex] } else { '' }
                if (-not [string]::IsNullOrWhiteSpace($cellValue)) { $hasContent = $true }
                $record[$normalizedHeader] = $cellValue
            }
            if ($hasContent) { $records += [pscustomobject]$record }
        }

        return [pscustomobject]@{
            SheetName = [string]$sheet.name
            Headers = $headers
            NormalizedHeaders = $normalizedHeaders
            Records = $records
        }
    }
    catch {
        throw 'Detail workbook is corrupt or unsupported.'
    }
    finally {
        if ($archive) { $archive.Dispose() }
    }
}

function ConvertTo-GssFeedbackDate {
    param([object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $number = 0.0
    if ([double]::TryParse([string]$Value, [ref]$number) -and $number -gt 1000) {
        return [datetime]::FromOADate($number).Date
    }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [ref]$date)) { return $date.Date }
    return $null
}

function Get-GssFeedbackProperty {
    param([object]$Record, [string]$Name)
    $property = $Record.PSObject.Properties[$Name]
    if ($property) { return [string]$property.Value }
    return ''
}

function Test-GssFeedbackAnswers {
    param([object]$Record)

    $ranges = [ordered]@{
        overall = @(1, 5)
        service = @(1, 5)
        culinary = @(1, 5)
        value = @(1, 5)
        paceofmeal = @(1, 5)
        recommend = @(0, 10)
        eventbookingprocess = @(1, 5)
    }
    foreach ($name in $ranges.Keys) {
        $raw = Get-GssFeedbackProperty $Record $name
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $number = 0.0
        if (-not [double]::TryParse($raw, [ref]$number) -or $number -lt $ranges[$name][0] -or $number -gt $ranges[$name][1]) {
            throw "Invalid $name answer in the detail workbook."
        }
    }
}

function Read-GssDetailWorkbook {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$FolderPath
    )

    $table = Read-GssXlsxFirstWorksheet $Path
    foreach ($required in @('restaurantname', 'reservationdate', 'text')) {
        if (-not $table.NormalizedHeaders.Contains($required)) {
            throw "Detail workbook is missing required header '$required'."
        }
    }

    $portablePath = ConvertTo-GssDropboxRelativePath -Path $Path -FolderPath $FolderPath
    $responses = @()
    $rowNumber = 1
    foreach ($record in $table.Records) {
        $rowNumber++
        $restaurant = (Get-GssFeedbackProperty $record 'restaurantname').Trim()
        $visitDate = ConvertTo-GssFeedbackDate (Get-GssFeedbackProperty $record 'reservationdate')
        if ([string]::IsNullOrWhiteSpace($restaurant) -and $null -eq $visitDate -and [string]::IsNullOrWhiteSpace((Get-GssFeedbackProperty $record 'text'))) {
            continue
        }
        if ([string]::IsNullOrWhiteSpace($restaurant) -or $null -eq $visitDate) {
            throw "Detail workbook contains an incomplete response at row ${rowNumber}."
        }
        Test-GssFeedbackAnswers -Record $record

        $restaurantIdMatch = [regex]::Match($restaurant, '^\s*(\d{4})\b')
        $restaurantId = if ($restaurantIdMatch.Success) { $restaurantIdMatch.Groups[1].Value } else { Normalize-GssFeedbackHeader $restaurant }
        $time = (Get-GssFeedbackProperty $record 'reservationtime').Trim()
        $text = (Get-GssFeedbackProperty $record 'text').Trim()
        $answerNames = @('overall', 'service', 'culinary', 'value', 'paceofmeal', 'recommend', 'managervisit', 'steakcookedcorrectly', 'eventbookingprocess', 'firstvisit')
        $answers = [ordered]@{}
        foreach ($answerName in $answerNames) { $answers[$answerName] = (Get-GssFeedbackProperty $record $answerName).Trim() }
        $timestampKey = "$($visitDate.ToString('yyyy-MM-dd')) $((Normalize-GssFeedbackText $time))".Trim()
        $identity = @(
            (Normalize-GssFeedbackText $restaurant),
            $timestampKey,
            (Normalize-GssFeedbackText $text),
            (($answerNames | ForEach-Object { Normalize-GssFeedbackText $answers[$_] }) -join '|')
        ) -join "`n"
        $dncValue = (Get-GssFeedbackProperty $record 'alertguestsdonotcontact').Trim()
        $doNotContact = (-not [string]::IsNullOrWhiteSpace($dncValue)) -and ($dncValue -notmatch '^(0|n|no|false)$')

        $responses += [pscustomobject]@{
            ResponseHash = Get-GssStringSha256 $identity
            RestaurantId = $restaurantId
            Restaurant = $restaurant
            VisitDate = $visitDate
            ReservationTime = $time
            Text = $text
            Answers = [pscustomobject]$answers
            DoNotContact = $doNotContact
            GuestFirstName = (Get-GssFeedbackProperty $record 'guestfirstname').Trim()
            GuestLastName = (Get-GssFeedbackProperty $record 'guestlastname').Trim()
            SourcePath = $portablePath
            SourceRow = $rowNumber
        }
    }
    if ($responses.Count -eq 0) { throw 'Detail workbook contains no usable responses.' }

    return [pscustomobject]@{
        Path = $Path
        PortablePath = $portablePath
        HeaderCount = @($table.NormalizedHeaders | Where-Object { $_ }).Count
        Headers = $table.NormalizedHeaders
        Responses = $responses
        VisitDateStart = ($responses.VisitDate | Measure-Object -Minimum).Minimum
        VisitDateEnd = ($responses.VisitDate | Measure-Object -Maximum).Maximum
    }
}

function Get-GssKnownGuestNames {
    param([object[]]$Responses)

    $names = @()
    foreach ($response in $Responses) {
        foreach ($value in @($response.GuestFirstName, $response.GuestLastName)) {
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            $names += $value.Trim()
            $names += @([regex]::Matches($value, "[A-Za-z][A-Za-z'-]{1,}") | ForEach-Object { $_.Value })
        }
    }
    $uniqueNames = @($names | Where-Object { $_.Length -ge 2 } | Sort-Object -Unique)
    return @($uniqueNames | Sort-Object `
        @{ Expression = { $_.Length }; Descending = $true }, `
        @{ Expression = { $_ }; Descending = $false })
}

function Get-GssPiiRedactionRules {
    return @(
        [pscustomobject]@{ Label = 'email'; Pattern = '(?i)\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b'; Replacement = '[REDACTED EMAIL]' },
        [pscustomobject]@{ Label = 'url'; Pattern = '(?i)(?<![@\w])(?<url>(?:(?:https?://|www\.)[^\s<>"\x27]*?|(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}(?:/[^\s<>"\x27]*?)?))(?<punct>[.,;:!?)}\]]*)(?=\s|$)'; Replacement = '[REDACTED URL]${punct}' },
        [pscustomobject]@{ Label = 'phone'; Pattern = '(?<![\w\d.])(?:(?:\+?1[\s.-]+)?[2-9]\d{2}\.[2-9]\d{6}|(?!\d+\.\d+(?![\w\d]|\.\d))(?:\+?\d{10,15}|(?:\+?\d{1,3}[\s.-]+)?(?:\(\d{2,4}\)[\s.-]*|\d{2,4}[\s.-]+)\d{2,4}[\s.-]*\d{4}|\d{3}[- ]\d{4}))(?![\w\d]|\.\d)'; Replacement = '[REDACTED PHONE]' },
        [pscustomobject]@{ Label = 'booking_identifier'; Pattern = '(?i)\b(?:check|confirmation|booking|reservation|resy)\s*(?:(?:id|number|no\.?)\s*)?(?:#|:)?\s*(?=[A-Z0-9-]{4,}\b)(?=[A-Z0-9-]*\d)[A-Z0-9-]{4,}\b'; Replacement = '[REDACTED BOOKING ID]' }
    )
}

function Get-GssUnsafeControlPattern {
    # Permit normal horizontal tab and line endings while rejecting other C0/C1
    # controls and Unicode bidi controls that can disguise package text.
    return '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\u061C\u200E\u200F\u202A-\u202E\u2066-\u2069]'
}

function Get-GssRemainingPiiTypes {
    param([string]$Text, [string[]]$KnownNames)

    $remainingPii = @()
    foreach ($rule in @(Get-GssPiiRedactionRules)) {
        if ([regex]::IsMatch([string]$Text, $rule.Pattern)) { $remainingPii += $rule.Label }
    }
    if ([regex]::IsMatch([string]$Text, (Get-GssUnsafeControlPattern))) { $remainingPii += 'unsafe_control' }
    foreach ($name in @($KnownNames)) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ([regex]::IsMatch([string]$Text, '(?i)(?<![A-Za-z])' + [regex]::Escape($name) + '(?![A-Za-z])')) {
            $remainingPii += 'known_name'
        }
    }
    return @($remainingPii | Select-Object -Unique)
}

function Test-GssTextContainsMachineSpecificPath {
    param([AllowNull()][AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrEmpty($Text)) { return $false }

    $trimmedText = $Text.Trim()
    $looksLikeJsonContainer = $trimmedText.Length -ge 2 -and (
        ($trimmedText[0] -eq '{' -and $trimmedText[$trimmedText.Length - 1] -eq '}') -or
        ($trimmedText[0] -eq '[' -and $trimmedText[$trimmedText.Length - 1] -eq ']')
    )
    if ($looksLikeJsonContainer) {
        $decodedValue = $null
        try {
            # Inspect decoded values, not JSON escape syntax. In particular,
            # A:\" inside serialized JSON is the text A:" and is not a drive root.
            $decodedValue = $trimmedText | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            # Malformed JSON is ordinary text and must still be scanned below.
            $decodedValue = $null
        }
        if ($null -ne $decodedValue) {
            return Test-GssStructuredValueContainsMachineSpecificPath -Value $decodedValue
        }
    }

    # Rooted and drive-relative paths must begin at a token boundary. A
    # drive-relative form must contain a later separator so labels such as
    # "A: acceptable" are not treated as paths. Separators that only escape a
    # JSON quote are excluded. UNC/network forms require server and share names.
    $rootedDrivePathPattern = '(?i)(?<![A-Za-z0-9])[A-Z]:[\\/](?!["''])'
    $relativeDrivePathPattern = '(?i)(?<![A-Za-z0-9])[A-Z]:(?![\\/\s"''])(?=[^\\/\s"''<>|]+[\\/](?!["'']))'
    $backslashUncPathPattern = '(?<![\\])\\\\(?![\\/"''])[^\\/\s"''<>|]+[\\/](?![\\/"''])[^\\/\s"''<>|]+'
    $forwardSlashNetworkPathPattern = '(?<![:/A-Za-z0-9._~-])//(?![/"''])[^/\s"''<>]+/(?![/"''])[^/\s"''<>]+'

    # Scheme-qualified HTTP URLs are portable even when their URL path or query
    # contains path-shaped text. Before removing a URL token, introduce a scan
    # boundary at an adjacent Windows path so URL stripping cannot consume it.
    # A forward slash is excluded as a delimiter because /C:/... may itself be
    # an ordinary URL path segment.
    $drivePathAfterUrlDelimiterPattern = '(?i)(?<=[^A-Za-z0-9_/])(?=(?:[A-Z]:[\\/](?!["''])|[A-Z]:(?![\\/\s"''])(?=[^\\/\s"''<>|]+[\\/](?!["'']))))'
    $uncPathAfterUrlPattern = '(?<![\\])(?=\\\\(?![\\/"''])[^\\/\s"''<>|]+[\\/](?![\\/"''])[^\\/\s"''<>|]+)'
    $scanText = [regex]::Replace($Text, $drivePathAfterUrlDelimiterPattern, ' ')
    $scanText = [regex]::Replace($scanText, $uncPathAfterUrlPattern, ' ')
    $scanText = [regex]::Replace($scanText, '(?i)\bhttps?://[^\s<>"'']+', '')

    return (
        [regex]::IsMatch($scanText, $rootedDrivePathPattern) -or
        [regex]::IsMatch($scanText, $relativeDrivePathPattern) -or
        [regex]::IsMatch($scanText, $backslashUncPathPattern) -or
        [regex]::IsMatch($scanText, $forwardSlashNetworkPathPattern)
    )
}

function Test-GssStructuredValueContainsMachineSpecificPath {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return $false }
    if ($Value -is [string]) {
        return Test-GssTextContainsMachineSpecificPath -Text ([string]$Value)
    }
    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in $Value.Keys) {
            if (Test-GssTextContainsMachineSpecificPath -Text ([string]$key)) { return $true }
            if (Test-GssStructuredValueContainsMachineSpecificPath -Value $Value[$key]) { return $true }
        }
        return $false
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        foreach ($item in $Value) {
            if (Test-GssStructuredValueContainsMachineSpecificPath -Value $item) { return $true }
        }
        return $false
    }
    if ($Value -isnot [System.Management.Automation.PSCustomObject]) { return $false }

    foreach ($property in $Value.PSObject.Properties) {
        if ($property.MemberType -notin @('NoteProperty', 'Property')) { continue }
        if (Test-GssTextContainsMachineSpecificPath -Text ([string]$property.Name)) { return $true }
        if (Test-GssStructuredValueContainsMachineSpecificPath -Value $property.Value) { return $true }
    }
    return $false
}

function Assert-GssPortableContentHasNoMachineSpecificPath {
    param(
        [AllowEmptyCollection()][object[]]$StructuredValues = @(),
        [AllowEmptyCollection()][string[]]$TextValues = @()
    )

    foreach ($value in @($StructuredValues)) {
        if (Test-GssStructuredValueContainsMachineSpecificPath -Value $value) {
            throw 'A machine-specific path leaked into a portable package file.'
        }
    }
    foreach ($text in @($TextValues)) {
        if (Test-GssTextContainsMachineSpecificPath -Text $text) {
            throw 'A machine-specific path leaked into a portable package file.'
        }
    }
}

function Protect-GssFeedbackText {
    param([string]$Text, [string[]]$KnownNames)

    $sanitized = if ($null -eq $Text) { '' } else { [string]$Text }
    $redactionCount = 0
    $controlMatches = [regex]::Matches($sanitized, (Get-GssUnsafeControlPattern))
    $redactionCount += $controlMatches.Count
    $sanitized = [regex]::Replace($sanitized, (Get-GssUnsafeControlPattern), '[REDACTED CONTROL]')
    $rules = @(Get-GssPiiRedactionRules)
    foreach ($rule in $rules) {
        $matches = [regex]::Matches($sanitized, $rule.Pattern)
        $redactionCount += $matches.Count
        $sanitized = [regex]::Replace($sanitized, $rule.Pattern, $rule.Replacement)
    }
    foreach ($name in @($KnownNames)) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $pattern = '(?i)(?<![A-Za-z])' + [regex]::Escape($name) + '(?![A-Za-z])'
        $matches = [regex]::Matches($sanitized, $pattern)
        $redactionCount += $matches.Count
        $sanitized = [regex]::Replace($sanitized, $pattern, '[REDACTED NAME]')
    }
    $sanitized = [regex]::Replace($sanitized, '\s+', ' ').Trim()

    $remainingPii = @(Get-GssRemainingPiiTypes -Text $sanitized -KnownNames $KnownNames)

    return [pscustomobject]@{
        Text = $sanitized
        RedactionCount = $redactionCount
        PiiScanPassed = ($remainingPii.Count -eq 0)
        RemainingPiiTypes = @($remainingPii | Select-Object -Unique)
    }
}

function Get-GssFeedbackThemes {
    param([object]$Response, [string]$SanitizedText)

    $text = Normalize-GssFeedbackText $SanitizedText
    $themes = @()
    $patterns = [ordered]@{
        service = '\b(service|server|waiter|waitress|staff|attentive|courteous|professional)\b'
        culinary = '\b(food|steak|meal|dish|salad|potato|crab|lamb|cook|cooked|temperature|delicious)\b'
        pace = '\b(wait|waiting|slow|rushed|timing|delay|minutes|pace)\b'
        value = '\b(value|price|priced|cost|expensive|worth|portion|portions)\b'
        'hospitality/recovery' = '\b(manager|management|apology|apologize|resolve|resolved|recovery|follow.?up|care|complaint)\b'
        recognition = '\b(birthday|anniversary|celebration|occasion|recognition|recognize|acknowledge|memorable)\b'
    }
    foreach ($theme in $patterns.Keys) {
        if ($text -match $patterns[$theme]) { $themes += $theme }
    }
    $answerThemeMap = [ordered]@{ service = 'service'; culinary = 'culinary'; paceofmeal = 'pace'; value = 'value' }
    foreach ($answerName in $answerThemeMap.Keys) {
        $raw = Get-GssFeedbackProperty $Response.Answers $answerName
        $number = 0.0
        if ([double]::TryParse($raw, [ref]$number) -and $number -le 3) { $themes += $answerThemeMap[$answerName] }
    }
    $overall = 0.0
    $manager = Get-GssFeedbackProperty $Response.Answers 'managervisit'
    if ([double]::TryParse((Get-GssFeedbackProperty $Response.Answers 'overall'), [ref]$overall) -and $overall -le 3 -and $manager -match '^(?i:n|no)$') {
        $themes += 'hospitality/recovery'
    }
    return @($themes | Select-Object -Unique)
}

function Get-GssFeedbackSentiment {
    param([object]$Response, [string]$SanitizedText)

    $text = Normalize-GssFeedbackText $SanitizedText
    $concern = $text -match '\b(disappoint|disappointed|poor|cold|overcooked|undercooked|rude|slow|expensive|complaint|worst|not good|burnt|scorched)\b'
    $positive = $text -match '\b(excellent|wonderful|great|outstanding|delicious|fantastic|perfect|memorable|exceptional)\b'
    foreach ($name in @('overall', 'service', 'culinary', 'value', 'paceofmeal')) {
        $number = 0.0
        if ([double]::TryParse((Get-GssFeedbackProperty $Response.Answers $name), [ref]$number)) {
            if ($number -le 3) { $concern = $true }
            if ($number -ge 4) { $positive = $true }
        }
    }
    if ($concern -and $positive) { return 'mixed' }
    if ($concern) { return 'concern' }
    if ($positive) { return 'positive' }
    return 'neutral'
}

function Read-GssFeedbackLedger {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ schema_version = $script:GssFeedbackLedgerVersion; entries = @() }
    }
    try {
        $ledger = Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json
        if ($ledger.schema_version -ne $script:GssFeedbackLedgerVersion) {
            throw "Unsupported feedback ledger version: $($ledger.schema_version)"
        }
        return $ledger
    }
    catch {
        throw "Feedback ledger is corrupt: $($_.Exception.Message)"
    }
}

function Invoke-GssFileSystemRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Operation,
        [int]$MaxAttempts = 12,
        [int]$DelayMilliseconds = 500
    )

    if ($MaxAttempts -lt 1) { throw 'MaxAttempts must be at least 1.' }
    if ($DelayMilliseconds -lt 0) { throw 'DelayMilliseconds cannot be negative.' }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return (& $Operation)
        }
        catch {
            if ($attempt -ge $MaxAttempts) { throw }
            if ($DelayMilliseconds -gt 0) {
                Start-Sleep -Milliseconds $DelayMilliseconds
            }
        }
    }
}

function Publish-GssStagedEmailPackage {
    param(
        [Parameter(Mandatory)][string]$StagingPath,
        [Parameter(Mandatory)][string]$PackagePath,
        [Parameter(Mandatory)][scriptblock]$ValidationOperation,
        [int]$MaxAttempts = 12,
        [int]$DelayMilliseconds = 500
    )

    $resolvedStaging = [System.IO.Path]::GetFullPath($StagingPath)
    $resolvedPackage = [System.IO.Path]::GetFullPath($PackagePath)
    if ([string]::Equals($resolvedStaging, $resolvedPackage, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Staging and package directories must be different.'
    }
    $stagingParent = [System.IO.Path]::GetDirectoryName($resolvedStaging)
    $packageParent = [System.IO.Path]::GetDirectoryName($resolvedPackage)
    if (-not [string]::Equals($stagingParent, $packageParent, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Staging and package directories must share the same outbox parent.'
    }

    $promoted = $false
    try {
        Invoke-GssFileSystemRetry -MaxAttempts $MaxAttempts -DelayMilliseconds $DelayMilliseconds -Operation {
            # Same-parent paths guarantee a same-volume rename. Directory.Move
            # fails if another publisher created the destination first instead
            # of nesting this staging directory inside that publisher's package.
            [System.IO.Directory]::Move($resolvedStaging, $resolvedPackage)
        }
        $promoted = $true
        return (Invoke-GssFileSystemRetry -MaxAttempts $MaxAttempts -DelayMilliseconds $DelayMilliseconds -Operation $ValidationOperation)
    }
    catch {
        $publicationError = $_
        $cleanupPath = if ($promoted) { $resolvedPackage } else { $resolvedStaging }
        if (Test-Path -LiteralPath $cleanupPath -PathType Container) {
            try {
                Invoke-GssFileSystemRetry -MaxAttempts $MaxAttempts -DelayMilliseconds $DelayMilliseconds -Operation {
                    Remove-Item -LiteralPath $cleanupPath -Recurse -Force
                }
            }
            catch {
                $key = if ($promoted) { 'GssPromotedPackageCleanupError' } else { 'GssStagingCleanupError' }
                $publicationError.Exception.Data[$key] = $_.Exception.Message
            }
        }
        throw
    }
}

function Write-GssFeedbackLedger {
    param([object]$Ledger, [string]$Path)

    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $temporary = "$Path.tmp.$([guid]::NewGuid().ToString('N'))"
    try {
        $Ledger | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $temporary -Encoding UTF8
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Get-GssByteArraySha256 {
    param([Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return (($sha.ComputeHash($Bytes) | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    finally {
        $sha.Dispose()
    }
}

function Get-GssFeedbackLedgerRollbackState {
    param([Parameter(Mandatory)][string]$Path)

    if (Test-Path -LiteralPath $Path) {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            throw 'The feedback ledger path exists but is not a file.'
        }
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        return [pscustomobject]@{
            Existed = $true
            Bytes = [byte[]]$bytes
            ByteSize = [long]$bytes.Length
            Sha256 = Get-GssByteArraySha256 -Bytes $bytes
        }
    }

    return [pscustomobject]@{
        Existed = $false
        Bytes = [byte[]]@()
        ByteSize = [long]0
        Sha256 = $null
    }
}

function Restore-GssFeedbackLedgerState {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$State
    )

    if (-not [bool]$State.Existed) {
        if (Test-Path -LiteralPath $Path) {
            if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
                throw 'The feedback ledger rollback target is not a file.'
            }
            Invoke-GssFileSystemRetry -Operation {
                Remove-Item -LiteralPath $Path -Force
            }
        }
        if (Test-Path -LiteralPath $Path) {
            throw 'The newly created feedback ledger could not be removed during rollback.'
        }
        return
    }

    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    if (Test-Path -LiteralPath $Path -PathType Container) {
        throw 'The feedback ledger rollback target is not a file.'
    }

    $temporary = Join-Path $directory ('.gss-ledger-rollback-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
    $replacementBackup = Join-Path $directory ('.gss-ledger-rollback-{0}.bak' -f [guid]::NewGuid().ToString('N'))
    try {
        [System.IO.File]::WriteAllBytes($temporary, [byte[]]$State.Bytes)
        Invoke-GssFileSystemRetry -Operation {
            if (Test-Path -LiteralPath $Path -PathType Leaf) {
                [System.IO.File]::Replace($temporary, $Path, $replacementBackup, $true)
            }
            else {
                [System.IO.File]::Move($temporary, $Path)
            }
        }

        $restoredBytes = [System.IO.File]::ReadAllBytes($Path)
        if ([long]$restoredBytes.Length -ne [long]$State.ByteSize -or
            (Get-GssByteArraySha256 -Bytes $restoredBytes) -cne [string]$State.Sha256) {
            throw 'Feedback ledger rollback did not restore the exact prior bytes.'
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporary -PathType Leaf) {
            Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $replacementBackup -PathType Leaf) {
            Remove-Item -LiteralPath $replacementBackup -Force -ErrorAction SilentlyContinue
        }
    }
}

function Write-GssFeedbackLedgerWithRollbackOnFailure {
    param(
        [Parameter(Mandatory)][object]$Ledger,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$RollbackState
    )

    try {
        Write-GssFeedbackLedger -Ledger $Ledger -Path $Path
    }
    catch {
        $ledgerWriteError = $_
        try {
            Restore-GssFeedbackLedgerState -Path $Path -State $RollbackState
        }
        catch {
            $ledgerWriteError.Exception.Data['GssFeedbackLedgerRollbackError'] = $_.Exception.Message
        }
        throw
    }
}

function ConvertTo-GssPackageInputPortablePath {
    param([Parameter(Mandatory)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or $Path -cne $Path.Trim() -or $Path.Contains('\')) {
        throw 'Package-input evidence contains a noncanonical portable path.'
    }
    if (-not $Path.StartsWith('gss/', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Package-input evidence portable paths must start with gss/.'
    }
    $segments = @($Path.Split('/'))
    if ($segments.Count -lt 2 -or
        @($segments | Where-Object { [string]::IsNullOrWhiteSpace($_) -or $_ -eq '.' -or $_ -eq '..' }).Count -gt 0) {
        throw 'Package-input evidence contains a noncanonical portable path.'
    }
    return ($segments -join '/')
}

function ConvertTo-GssPackageInputRecord {
    param(
        [Parameter(Mandatory)][object]$InputRecord,
        [Parameter(Mandatory)][string]$Label
    )

    Assert-GssExactObjectSchema `
        -Value $InputRecord `
        -AllowedProperties @('PortablePath', 'ByteSize', 'Sha256') `
        -Label $Label
    $portablePath = ConvertTo-GssPackageInputPortablePath -Path ([string]$InputRecord.PortablePath)
    $byteSizeText = [string]$InputRecord.ByteSize
    $byteSize = [long]0
    if ($byteSizeText -notmatch '^(?:0|[1-9][0-9]*)$' -or
        -not [long]::TryParse($byteSizeText, [ref]$byteSize)) {
        throw "$Label has an invalid byte size."
    }
    $sha256 = [string]$InputRecord.Sha256
    if ($sha256 -cnotmatch '^[a-f0-9]{64}$') {
        throw "$Label has an invalid SHA-256 hash."
    }
    return [pscustomobject][ordered]@{
        PortablePath = $portablePath
        ByteSize = $byteSize
        Sha256 = $sha256
    }
}

function Get-GssPackageInputEvidenceSha256 {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Inputs)

    $canonicalEvidence = @($Inputs |
        Sort-Object @{ Expression = { ([string]$_.PortablePath).ToLowerInvariant() } } |
        ForEach-Object {
            "$(([string]$_.PortablePath).ToLowerInvariant()):$([long]$_.ByteSize):$(([string]$_.Sha256).ToLowerInvariant())"
        }) -join "`n"
    return Get-GssStringSha256 -Value $canonicalEvidence
}

function Get-GssCurrentPackageInputEvidence {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SourceDescriptors,
        [Parameter(Mandatory)][string]$LedgerPath,
        [object]$LedgerState
    )

    $records = @()
    foreach ($source in $SourceDescriptors) {
        $sourcePath = [string]$source.source_path
        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            throw 'A captured package source descriptor is missing its portable path.'
        }
        $portablePath = ConvertTo-GssPackageInputPortablePath -Path ('gss/' + $sourcePath.Replace('\', '/'))
        $sourceFullPath = [string]$source.source_full_path
        if (-not (Test-Path -LiteralPath $sourceFullPath -PathType Leaf)) {
            throw "A captured package source is missing: $portablePath"
        }
        $sourceItem = Get-Item -LiteralPath $sourceFullPath
        $records += [pscustomobject][ordered]@{
            PortablePath = $portablePath
            ByteSize = [long]$sourceItem.Length
            Sha256 = Get-GssSha256 -Path $sourceFullPath
        }
    }

    $ledgerEvidenceState = if ($null -ne $LedgerState) {
        $LedgerState
    }
    else {
        Get-GssFeedbackLedgerRollbackState -Path $LedgerPath
    }
    if ([bool]$ledgerEvidenceState.Existed) {
        $records += [pscustomobject][ordered]@{
            PortablePath = 'gss/_automation_runs/state/gss_feedback_first_seen.json'
            ByteSize = [long]$ledgerEvidenceState.ByteSize
            Sha256 = [string]$ledgerEvidenceState.Sha256
        }
    }

    $byPath = @{}
    foreach ($record in $records) {
        $key = ([string]$record.PortablePath).ToLowerInvariant()
        if ($byPath.ContainsKey($key)) {
            throw "A package input was captured more than once: $($record.PortablePath)"
        }
        $byPath[$key] = $record
    }
    $sortedRecords = @($byPath.Keys | Sort-Object | ForEach-Object { $byPath[$_] })
    return [pscustomobject]@{
        Inputs = $sortedRecords
        SourceSetSha256 = Get-GssPackageInputEvidenceSha256 -Inputs $sortedRecords
    }
}

function Assert-GssExpectedPackageInputEvidence {
    param(
        [AllowNull()][object]$ExpectedEvidence,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SourceDescriptors,
        [Parameter(Mandatory)][string]$LedgerPath,
        [object]$LedgerState
    )

    if ($null -eq $ExpectedEvidence) { return $null }
    if ($null -eq $ExpectedEvidence.PSObject.Properties['Inputs'] -or
        $null -eq $ExpectedEvidence.PSObject.Properties['SourceSetSha256']) {
        throw 'Expected package-input evidence must contain Inputs and SourceSetSha256.'
    }

    $expectedRecords = @()
    $expectedByPath = @{}
    $recordIndex = 0
    foreach ($record in @($ExpectedEvidence.Inputs)) {
        $recordIndex++
        $normalized = ConvertTo-GssPackageInputRecord `
            -InputRecord $record `
            -Label "Expected package-input record $recordIndex"
        $key = ([string]$normalized.PortablePath).ToLowerInvariant()
        if ($expectedByPath.ContainsKey($key)) {
            throw "Expected package-input evidence contains a duplicate path: $($normalized.PortablePath)"
        }
        $expectedByPath[$key] = $normalized
        $expectedRecords += $normalized
    }
    if ($expectedRecords.Count -eq 0) {
        throw 'Expected package-input evidence must contain at least one input.'
    }

    $expectedSourceSetSha256 = [string]$ExpectedEvidence.SourceSetSha256
    if ($expectedSourceSetSha256 -cnotmatch '^[a-f0-9]{64}$') {
        throw 'Expected package-input evidence has an invalid SourceSetSha256.'
    }
    $computedExpectedSha256 = Get-GssPackageInputEvidenceSha256 -Inputs $expectedRecords
    if ($computedExpectedSha256 -cne $expectedSourceSetSha256) {
        throw 'Expected package-input evidence SourceSetSha256 does not match its Inputs.'
    }

    $actualEvidence = Get-GssCurrentPackageInputEvidence `
        -SourceDescriptors $SourceDescriptors `
        -LedgerPath $LedgerPath `
        -LedgerState $LedgerState
    $actualByPath = @{}
    foreach ($record in @($actualEvidence.Inputs)) {
        $actualByPath[([string]$record.PortablePath).ToLowerInvariant()] = $record
    }
    foreach ($key in @($actualByPath.Keys | Sort-Object)) {
        if (-not $expectedByPath.ContainsKey($key)) {
            throw "Current package inputs contain an unexpected source: $($actualByPath[$key].PortablePath)"
        }
    }
    foreach ($key in @($expectedByPath.Keys | Sort-Object)) {
        if (-not $actualByPath.ContainsKey($key)) {
            throw "Expected package input is missing from the current source set: $($expectedByPath[$key].PortablePath)"
        }
        $expected = $expectedByPath[$key]
        $actual = $actualByPath[$key]
        if ([long]$actual.ByteSize -ne [long]$expected.ByteSize -or
            [string]$actual.Sha256 -cne [string]$expected.Sha256) {
            throw "Current package input changed after the committed snapshot: $($actual.PortablePath)"
        }
    }
    if ($actualByPath.Count -ne $expectedByPath.Count -or
        [string]$actualEvidence.SourceSetSha256 -cne $expectedSourceSetSha256) {
        throw 'Current package inputs do not match the expected exact source set.'
    }
    return $actualEvidence
}

function Assert-GssSundayReportingDate {
    param([Parameter(Mandatory)][datetime]$ReportingDate)

    if ($ReportingDate.Date.DayOfWeek -ne [System.DayOfWeek]::Sunday) {
        throw "GSS reporting date must be a Sunday: $($ReportingDate.ToString('yyyy-MM-dd'))."
    }
}

function Get-GssHistoricalRecoveryResponseSetSha256 {
    param([Parameter(Mandatory)][object]$Workbook)

    $responseHashes = @(
        $Workbook.Responses.ResponseHash |
            ForEach-Object { ([string]$_).ToLowerInvariant() } |
            Sort-Object -Unique
    )
    if ($responseHashes.Count -eq 0) {
        throw 'A recovered historical detail workbook has no response identities.'
    }
    $material = "$($script:GssHistoricalResponseSetVersion)`n" + ($responseHashes -join "`n")
    return Get-GssStringSha256 $material
}

function Get-GssVerifiedHistoricalRecoveryInventory {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.IO.FileInfo[]]$RecoveredFiles
    )

    $resolvedFolder = [System.IO.Path]::GetFullPath($FolderPath).TrimEnd('\', '/')
    $recoveredRoot = [System.IO.Path]::GetFullPath(
        (Join-Path $resolvedFolder $script:GssHistoricalRecoveryArchivePrefix.Replace('/', '\'))
    ).TrimEnd('\', '/')
    $recoveredPrefix = "$recoveredRoot\"

    $runtimeRoot = Join-Path $resolvedFolder '_automation_runs\historical-recovery'
    if (-not (Test-Path -LiteralPath $runtimeRoot -PathType Container)) {
        if (@($RecoveredFiles).Count -eq 0) {
            return [pscustomobject]@{
                DescriptorsByPath = @{}
                ManifestSha256 = @()
            }
        }
        throw 'Recovered historical detail exists without a historical-recovery transaction directory.'
    }

    $receiptPaths = @(
        Get-ChildItem -LiteralPath $runtimeRoot -Directory |
            ForEach-Object {
                $candidate = Join-Path $_.FullName 'transaction-receipt.json'
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { $candidate }
            } |
            Sort-Object
    )
    $expectedByPath = @{}
    $manifestHashes = @()
    foreach ($receiptPath in $receiptPaths) {
        try {
            $receipt = Get-Content -Raw -LiteralPath $receiptPath | ConvertFrom-Json
        }
        catch {
            throw "Historical recovery receipt is not valid JSON: $receiptPath"
        }
        if ([string]$receipt.schema_version -ne $script:GssHistoricalRecoveryReceiptVersion) {
            throw "Unsupported historical recovery receipt version: $($receipt.schema_version)"
        }
        if ([string]$receipt.state -ne 'Committed') {
            continue
        }

        $manifestSha256 = ([string]$receipt.manifest_sha256).ToLowerInvariant()
        if ($manifestSha256 -notmatch '^[a-f0-9]{64}$') {
            throw "Committed historical recovery receipt has an invalid manifest SHA-256: $receiptPath"
        }
        $transactionDirectory = Split-Path -Parent $receiptPath
        if ((Split-Path -Leaf $transactionDirectory) -cne $manifestSha256) {
            throw "Committed historical recovery receipt is outside its manifest-hash transaction directory: $receiptPath"
        }
        if ([string]$receipt.transaction_id -ne "historical-recovery:$manifestSha256") {
            throw "Committed historical recovery receipt has an invalid transaction ID: $receiptPath"
        }
        if ((Split-Path -Leaf ([string]$receipt.manifest_snapshot_path)) -cne 'recovery-manifest.json') {
            throw "Committed historical recovery receipt has an invalid manifest snapshot name: $receiptPath"
        }

        $manifestPath = Join-Path $transactionDirectory 'recovery-manifest.json'
        if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
            throw "Committed historical recovery manifest snapshot is missing: $manifestPath"
        }
        if ((Get-GssSha256 $manifestPath) -cne $manifestSha256) {
            throw "Committed historical recovery manifest snapshot hash mismatch: $manifestPath"
        }
        try {
            $manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
        }
        catch {
            throw "Committed historical recovery manifest is not valid JSON: $manifestPath"
        }
        if ([string]$manifest.schema_version -ne $script:GssHistoricalRecoveryManifestVersion) {
            throw "Unsupported historical recovery manifest version: $($manifest.schema_version)"
        }
        $fiscalYear = [string]$manifest.fiscal_year
        if ($fiscalYear -notmatch '^FY\d{2}$') {
            throw "Committed historical recovery manifest has an invalid fiscal year: '$fiscalYear'."
        }
        $manifestSources = @($manifest.sources)
        $receiptFiles = @($receipt.files)
        if ($manifestSources.Count -eq 0 -or
            $receiptFiles.Count -ne $manifestSources.Count -or
            [int]$receipt.published_file_count -ne $manifestSources.Count) {
            throw "Committed historical recovery receipt and manifest file counts disagree: $receiptPath"
        }

        for ($index = 0; $index -lt $manifestSources.Count; $index++) {
            $source = $manifestSources[$index]
            $sourceIndex = $index + 1
            $reportWeek = [string]$source.source_report_week
            if ($reportWeek -notmatch '^FY\d{2} FW(?:[1-9]|[1-4]\d|5[0-3])$' -or
                -not $reportWeek.StartsWith("$fiscalYear ", [System.StringComparison]::Ordinal)) {
                throw "Committed historical recovery manifest source $sourceIndex has an invalid report week: '$reportWeek'."
            }
            $sha256 = ([string]$source.sha256).ToLowerInvariant()
            $responseSetSha256 = ([string]$source.response_set_sha256).ToLowerInvariant()
            if ($sha256 -notmatch '^[a-f0-9]{64}$' -or $responseSetSha256 -notmatch '^[a-f0-9]{64}$') {
                throw "Committed historical recovery manifest source $sourceIndex has an invalid SHA-256."
            }
            try {
                $byteSize = [long]$source.byte_size
                $rowCount = [int]$source.row_count
            }
            catch {
                throw "Committed historical recovery manifest source $sourceIndex has invalid numeric evidence."
            }
            if ($byteSize -lt 1 -or $rowCount -lt 1) {
                throw "Committed historical recovery manifest source $sourceIndex has nonpositive size or row-count evidence."
            }

            $expectedDestination = "$($script:GssHistoricalRecoveryArchivePrefix)/$fiscalYear/$($reportWeek.Replace(' ', '-'))-$sha256.xlsx"
            $portableDestination = ([string]$source.destination_path).Replace('\', '/').Trim('/')
            if (-not $portableDestination.Equals($expectedDestination, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Committed historical recovery manifest source $sourceIndex has a noncanonical destination."
            }
            $destinationFullPath = [System.IO.Path]::GetFullPath(
                (Join-Path $resolvedFolder $portableDestination.Replace('/', '\'))
            )
            if (-not $destinationFullPath.StartsWith($recoveredPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Committed historical recovery destination escapes the recovered-detail root: $portableDestination"
            }

            $receiptMatches = @($receiptFiles | Where-Object { [int]$_.source_index -eq $sourceIndex })
            if ($receiptMatches.Count -ne 1) {
                throw "Committed historical recovery receipt does not contain exactly one file for source $sourceIndex."
            }
            $receiptFile = $receiptMatches[0]
            if ([string]$receiptFile.sha256 -cne $sha256 -or
                [long]$receiptFile.byte_size -ne $byteSize -or
                [int]$receiptFile.row_count -ne $rowCount -or
                [string]$receiptFile.response_set_sha256 -cne $responseSetSha256 -or
                -not ([string]$receiptFile.destination_path).Replace('\', '/').Trim('/').Equals(
                    $portableDestination,
                    [System.StringComparison]::OrdinalIgnoreCase
                )) {
                throw "Committed historical recovery receipt source $sourceIndex disagrees with its manifest."
            }

            $pathKey = $destinationFullPath.ToLowerInvariant()
            if ($expectedByPath.ContainsKey($pathKey)) {
                throw "Recovered historical detail destination is attested more than once: $portableDestination"
            }
            $expectedByPath[$pathKey] = [pscustomobject]@{
                Path = $destinationFullPath
                PortablePath = $portableDestination
                SourceReportWeek = $reportWeek
                Sha256 = $sha256
                ByteSize = $byteSize
                RowCount = $rowCount
                ResponseSetSha256 = $responseSetSha256
                ManifestSha256 = $manifestSha256
            }
        }
        $manifestHashes += $manifestSha256
    }

    if ($expectedByPath.Count -eq 0) {
        if (@($RecoveredFiles).Count -eq 0) {
            return [pscustomobject]@{
                DescriptorsByPath = @{}
                ManifestSha256 = @()
            }
        }
        throw 'Recovered historical detail exists but no committed historical-recovery receipt covers it.'
    }
    $actualByPath = @{}
    foreach ($file in @($RecoveredFiles)) {
        $actualPath = [System.IO.Path]::GetFullPath($file.FullName)
        if (-not $actualPath.StartsWith($recoveredPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Recovered historical detail inventory contains an out-of-scope file: $actualPath"
        }
        $actualByPath[$actualPath.ToLowerInvariant()] = $file
        if (-not $expectedByPath.ContainsKey($actualPath.ToLowerInvariant())) {
            throw "Recovered historical detail file is not covered by a committed manifest: $actualPath"
        }
    }
    foreach ($pathKey in @($expectedByPath.Keys)) {
        if (-not $actualByPath.ContainsKey($pathKey)) {
            throw "Committed historical recovery file is missing: $($expectedByPath[$pathKey].PortablePath)"
        }
    }
    if ($actualByPath.Count -ne $expectedByPath.Count) {
        throw 'Recovered historical detail exact-set verification failed.'
    }

    return [pscustomobject]@{
        DescriptorsByPath = $expectedByPath
        ManifestSha256 = @($manifestHashes | Sort-Object -Unique)
    }
}

function Assert-GssHistoricalRecoveryFileIntegrity {
    param(
        [Parameter(Mandatory)][object]$Descriptor,
        [Parameter(Mandatory)][string]$Phase
    )

    if (-not (Test-Path -LiteralPath $Descriptor.Path -PathType Leaf)) {
        throw "Recovered historical detail file disappeared during $Phase verification: $($Descriptor.PortablePath)"
    }
    $item = Get-Item -LiteralPath $Descriptor.Path
    if ([long]$item.Length -ne [long]$Descriptor.ByteSize -or
        (Get-GssSha256 $Descriptor.Path) -cne [string]$Descriptor.Sha256) {
        throw "Recovered historical detail file hash or size mismatch during $Phase verification: $($Descriptor.PortablePath)"
    }
}

function Get-GssDetailInventory {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][datetime]$ReportingDate
    )

    Assert-GssSundayReportingDate -ReportingDate $ReportingDate
    $detailFolder = Join-Path $FolderPath '03 Uploaded Survey Workbooks'
    if (-not (Test-Path -LiteralPath $detailFolder -PathType Container)) {
        throw "Guest-detail folder is missing: 03 Uploaded Survey Workbooks"
    }
    $files = @(Get-ChildItem -LiteralPath $detailFolder -File -Filter '*.xlsx' -Recurse |
        Where-Object { $_.Name -notlike '~$*' -and $_.Length -gt 0 } |
        Sort-Object FullName)
    if ($files.Count -eq 0) { throw 'No guest-detail workbooks were found.' }

    $recoveredRoot = [System.IO.Path]::GetFullPath(
        (Join-Path $FolderPath $script:GssHistoricalRecoveryArchivePrefix.Replace('/', '\'))
    ).TrimEnd('\', '/')
    $recoveredPrefix = "$recoveredRoot\"
    $recoveredFiles = @($files | Where-Object {
        ([System.IO.Path]::GetFullPath($_.FullName)).StartsWith(
            $recoveredPrefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    })
    $recoveryInventory = Get-GssVerifiedHistoricalRecoveryInventory `
        -FolderPath $FolderPath `
        -RecoveredFiles $recoveredFiles

    $workbooks = @()
    foreach ($file in $files) {
        $pathKey = ([System.IO.Path]::GetFullPath($file.FullName)).ToLowerInvariant()
        $recoveryDescriptor = if ($recoveryInventory.DescriptorsByPath.ContainsKey($pathKey)) {
            $recoveryInventory.DescriptorsByPath[$pathKey]
        }
        else {
            $null
        }
        if ($null -ne $recoveryDescriptor) {
            Assert-GssHistoricalRecoveryFileIntegrity -Descriptor $recoveryDescriptor -Phase 'pre-parse'
        }
        $workbook = Read-GssDetailWorkbook -Path $file.FullName -FolderPath $FolderPath
        if ($null -ne $recoveryDescriptor) {
            Assert-GssHistoricalRecoveryFileIntegrity -Descriptor $recoveryDescriptor -Phase 'post-parse'
            if (@($workbook.Responses).Count -ne [int]$recoveryDescriptor.RowCount -or
                (Get-GssHistoricalRecoveryResponseSetSha256 -Workbook $workbook) -cne [string]$recoveryDescriptor.ResponseSetSha256) {
                throw "Recovered historical detail workbook row or response-set evidence mismatch: $($recoveryDescriptor.PortablePath)"
            }
        }
        $workbooks += $workbook
    }

    $directPrefix = [System.IO.Path]::GetFullPath($detailFolder).TrimEnd('\') + '\'
    $direct = @($workbooks | Where-Object {
        $parent = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($_.Path)).TrimEnd('\') + '\'
        $parent.Equals($directPrefix, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($direct.Count -eq 0) { throw 'No current guest-detail workbook exists directly in 03 Uploaded Survey Workbooks.' }
    $eligible = @($direct | Where-Object { $_.VisitDateEnd.Date -le $ReportingDate.Date })
    if ($eligible.Count -eq 0) {
        throw "No current guest-detail workbook has a visit date on or before reporting date $($ReportingDate.ToString('yyyy-MM-dd'))."
    }
    $current = $eligible |
        Sort-Object @{ Expression = { $_.VisitDateEnd }; Descending = $true }, @{ Expression = { $_.PortablePath }; Descending = $false } |
        Select-Object -First 1

    $groups = @{}
    foreach ($workbook in $workbooks) {
        foreach ($response in $workbook.Responses) {
            if (-not $groups.ContainsKey($response.ResponseHash)) { $groups[$response.ResponseHash] = @() }
            $groups[$response.ResponseHash] += $response
        }
    }
    $uniqueResponses = @()
    foreach ($hash in @($groups.Keys | Sort-Object)) {
        $instances = @($groups[$hash] | Sort-Object SourcePath, SourceRow)
        $canonical = $instances[0]
        $uniqueResponses += [pscustomobject]@{
            ResponseHash = $hash
            RestaurantId = $canonical.RestaurantId
            Restaurant = $canonical.Restaurant
            VisitDate = $canonical.VisitDate
            Text = $canonical.Text
            Answers = $canonical.Answers
            DoNotContact = [bool](@($instances | Where-Object { $_.DoNotContact }).Count -gt 0)
            GuestFirstName = $canonical.GuestFirstName
            GuestLastName = $canonical.GuestLastName
            SourcePaths = @($instances.SourcePath | Sort-Object -Unique)
        }
    }

    return [pscustomobject]@{
        CurrentWorkbook = $current
        Workbooks = $workbooks
        AllResponseInstances = @($workbooks.Responses)
        UniqueResponses = $uniqueResponses
        DuplicateResponseCount = (@($workbooks.Responses).Count - $uniqueResponses.Count)
    }
}

function Get-GssCommenterLens {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseSingularNouns',
        '',
        Justification = 'Commenter Lens is the established singular domain term and published artifact name.'
    )]
    param(
        [Parameter(Mandatory)][object]$Inventory,
        [Parameter(Mandatory)][object[]]$MetricDetail,
        [Parameter(Mandatory)][datetime]$ReportingDate,
        [switch]$ExactPartitionAlignmentVerified
    )

    Assert-GssSundayReportingDate -ReportingDate $ReportingDate
    $lensPolicy = $script:GssAnalysisPolicy.commenter_lens
    $windowWeeks = [int]$lensPolicy.reporting_window_weeks
    $windowEnd = $ReportingDate.Date
    $windowStart = $windowEnd.AddDays(-(($windowWeeks * 7) - 1))
    $eligibleResponses = @($Inventory.UniqueResponses | Where-Object {
        $null -ne $_.VisitDate -and
        $_.VisitDate.Date -ge $windowStart -and
        $_.VisitDate.Date -le $windowEnd
    })
    $restaurantIds = @(
        @($MetricDetail | ForEach-Object { [string]$_.RestaurantId }) +
        @($eligibleResponses | ForEach-Object { [string]$_.RestaurantId }) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )

    $hasDataQualityIssue = $false
    $hasUsableCommenterData = $false
    $restaurantResults = @()
    foreach ($restaurantId in $restaurantIds) {
        $restaurantResponses = @($eligibleResponses | Where-Object { [string]$_.RestaurantId -eq $restaurantId })
        $commenterCount = $restaurantResponses.Count
        if ($commenterCount -gt 0) { $hasUsableCommenterData = $true }

        $countCandidates = @(
            $MetricDetail |
                Where-Object {
                    [string]$_.RestaurantId -eq $restaurantId -and
                    $null -ne $_.CurrentCount
                } |
                ForEach-Object {
                    $candidate = 0.0
                    if ([double]::TryParse(
                        [string]$_.CurrentCount,
                        [System.Globalization.NumberStyles]::Float,
                        [System.Globalization.CultureInfo]::InvariantCulture,
                        [ref]$candidate
                    ) -and
                        -not [double]::IsNaN($candidate) -and
                        -not [double]::IsInfinity($candidate) -and
                        $candidate -ge 0 -and
                        [math]::Abs($candidate - [math]::Round($candidate)) -le 0.0000001) {
                        [long][math]::Round($candidate)
                    }
                } |
                Sort-Object -Unique
        )
        $populationCount = if ($countCandidates.Count -eq 1) { [long]$countCandidates[0] } else { $null }
        $restaurantIssues = @()
        if ($null -eq $populationCount) {
            $restaurantIssues += 'population_response_count_missing_or_inconsistent'
        }
        elseif ($ExactPartitionAlignmentVerified -and $commenterCount -gt $populationCount) {
            $restaurantIssues += 'commenter_count_exceeds_population_count'
        }

        $coverageIsStructurallyValid = $null -ne $populationCount -and
            $populationCount -gt 0 -and
            $commenterCount -le $populationCount
        $coverageIsValid = $coverageIsStructurallyValid -and [bool]$ExactPartitionAlignmentVerified
        $coverage = if ($coverageIsValid) {
            [math]::Round((100.0 * $commenterCount / $populationCount), 8)
        }
        else {
            $null
        }
        $coverageStatus = if (-not $ExactPartitionAlignmentVerified) {
            'SuppressedUnverifiedPartitionAlignment'
        }
        elseif ($coverageIsStructurallyValid) {
            'Ready'
        }
        else {
            'SuppressedInvalidCoverage'
        }

        $metricResults = @()
        foreach ($metricPolicy in @($lensPolicy.metrics)) {
            $metricId = [string]$metricPolicy.id
            $responseField = [string]$metricPolicy.response_field
            $populationMetricName = if ($null -eq $metricPolicy.population_metric) {
                $null
            }
            else {
                [string]$metricPolicy.population_metric
            }
            $eventMinimum = [double]$metricPolicy.event_minimum
            $eventMaximum = [double]$metricPolicy.event_maximum

            $scoredValues = @()
            foreach ($response in $restaurantResponses) {
                $rawScore = Get-GssFeedbackProperty $response.Answers $responseField
                $score = 0.0
                if (-not [string]::IsNullOrWhiteSpace($rawScore) -and
                    [double]::TryParse(
                        $rawScore,
                        [System.Globalization.NumberStyles]::Float,
                        [System.Globalization.CultureInfo]::InvariantCulture,
                        [ref]$score
                    ) -and
                    -not [double]::IsNaN($score) -and
                    -not [double]::IsInfinity($score)) {
                    $scoredValues += $score
                }
            }
            $scoredCount = $scoredValues.Count
            $missingScoreCount = $commenterCount - $scoredCount
            $eventCount = @($scoredValues | Where-Object { $_ -ge $eventMinimum -and $_ -le $eventMaximum }).Count
            $commenterEventRate = if ($scoredCount -gt 0) {
                [math]::Round((100.0 * $eventCount / $scoredCount), 8)
            }
            else {
                $null
            }

            $hasExactPopulationMetric = -not [string]::IsNullOrWhiteSpace($populationMetricName)
            $normalizedPopulationMetric = if ($hasExactPopulationMetric) {
                Normalize-GssFeedbackHeader $populationMetricName
            }
            else {
                ''
            }
            $populationMetricRows = @(
                if ($hasExactPopulationMetric) {
                    $MetricDetail | Where-Object {
                        [string]$_.RestaurantId -eq $restaurantId -and
                        (
                            (Normalize-GssFeedbackHeader ([string]$_.RawMetric)) -eq $normalizedPopulationMetric -or
                            (Normalize-GssFeedbackHeader ([string]$_.Metric)) -eq $normalizedPopulationMetric
                        )
                    }
                }
            )
            $populationEventRate = $null
            $populationMetricValid = $false
            if ($populationMetricRows.Count -eq 1 -and $null -ne $populationMetricRows[0].Current) {
                $candidateRate = 0.0
                if ([double]::TryParse(
                    [string]$populationMetricRows[0].Current,
                    [System.Globalization.NumberStyles]::Float,
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [ref]$candidateRate
                ) -and
                    -not [double]::IsNaN($candidateRate) -and
                    -not [double]::IsInfinity($candidateRate) -and
                    $candidateRate -ge 0 -and
                    $candidateRate -le 100) {
                    $populationEventRate = $candidateRate
                    $populationMetricValid = $true
                }
            }

            $comparisonStatus = 'Ready'
            if ($scoredCount -eq 0) {
                $comparisonStatus = 'SuppressedNoScoredCommenterResponses'
            }
            elseif (-not $hasExactPopulationMetric) {
                $comparisonStatus = 'NotAvailableNoExactPopulationMetric'
            }
            elseif (-not $ExactPartitionAlignmentVerified) {
                $comparisonStatus = 'SuppressedUnverifiedPartitionAlignment'
            }
            elseif (-not $coverageIsValid) {
                $comparisonStatus = 'SuppressedInvalidCoverage'
            }
            elseif ($missingScoreCount -gt 0) {
                $comparisonStatus = 'SuppressedMissingCommenterScores'
            }
            elseif ($populationMetricRows.Count -ne 1) {
                $comparisonStatus = 'SuppressedPopulationMetricDefinitionMismatch'
                $restaurantIssues += "population_metric_definition_mismatch:$metricId"
            }
            elseif (-not $populationMetricValid) {
                $comparisonStatus = 'SuppressedInvalidPopulationMetric'
                $restaurantIssues += "invalid_population_metric:$metricId"
            }

            $populationGap = if ($comparisonStatus -eq 'Ready') {
                [math]::Round(($commenterEventRate - $populationEventRate), 8)
            }
            else {
                $null
            }
            $materialGap = if ($null -ne $populationGap) {
                [math]::Abs($populationGap) -ge [double]$lensPolicy.material_gap_percentage_points
            }
            else {
                $null
            }
            $negativeOverrepresentation = if ($null -ne $populationGap) {
                $populationGap -ge [double]$lensPolicy.material_gap_percentage_points
            }
            else {
                $null
            }

            $reconstructedPopulationEventCount = $null
            $reconstructionError = $null
            $nonCommentCount = if ($coverageIsValid -and $ExactPartitionAlignmentVerified) {
                [long]($populationCount - $commenterCount)
            }
            else {
                $null
            }
            $derivedNonCommentEventCount = $null
            $derivedNonCommentEventRate = $null
            $nonCommentGap = $null
            $derivedStatus = 'Suppressed'
            if ($comparisonStatus -eq 'Ready') {
                $rawPopulationEventCount = $populationEventRate * $populationCount / 100.0
                $nearestPopulationEventCount = [math]::Round(
                    $rawPopulationEventCount,
                    0,
                    [System.MidpointRounding]::AwayFromZero
                )
                $reconstructionError = [math]::Abs($rawPopulationEventCount - $nearestPopulationEventCount)
                if ($reconstructionError -le [double]$lensPolicy.maximum_reconstructed_event_count_error) {
                    $reconstructedPopulationEventCount = [long]$nearestPopulationEventCount
                }

                if (-not $ExactPartitionAlignmentVerified) {
                    $derivedStatus = 'SuppressedUnverifiedPartitionAlignment'
                }
                elseif ($null -eq $reconstructedPopulationEventCount) {
                    $derivedStatus = 'SuppressedInexactPopulationEventCount'
                }
                elseif ($nonCommentCount -le 0) {
                    $derivedStatus = 'SuppressedNoNonCommentResponses'
                }
                else {
                    $candidateNonCommentEventCount = $reconstructedPopulationEventCount - $eventCount
                    if ($candidateNonCommentEventCount -lt 0 -or $candidateNonCommentEventCount -gt $nonCommentCount) {
                        $derivedStatus = 'SuppressedInvalidEventArithmetic'
                        $restaurantIssues += "invalid_event_arithmetic:$metricId"
                    }
                    else {
                        $derivedNonCommentEventCount = [long]$candidateNonCommentEventCount
                        $derivedNonCommentEventRate = [math]::Round(
                            (100.0 * $derivedNonCommentEventCount / $nonCommentCount),
                            8
                        )
                        $nonCommentGap = [math]::Round(
                            ($commenterEventRate - $derivedNonCommentEventRate),
                            8
                        )
                        $derivedStatus = 'Ready'
                    }
                }
            }
            elseif ($comparisonStatus -eq 'SuppressedNoScoredCommenterResponses') {
                $derivedStatus = 'SuppressedNoScoredCommenterResponses'
            }
            elseif ($comparisonStatus -eq 'SuppressedUnverifiedPartitionAlignment') {
                $derivedStatus = 'SuppressedUnverifiedPartitionAlignment'
            }
            elseif ($comparisonStatus -eq 'SuppressedInvalidCoverage') {
                $derivedStatus = 'SuppressedInvalidCoverage'
            }
            elseif ($comparisonStatus -eq 'SuppressedPopulationMetricDefinitionMismatch') {
                $derivedStatus = 'SuppressedPopulationMetricDefinitionMismatch'
            }
            elseif ($comparisonStatus -eq 'SuppressedInvalidPopulationMetric') {
                $derivedStatus = 'SuppressedInvalidPopulationMetric'
            }
            elseif ($comparisonStatus -eq 'SuppressedMissingCommenterScores') {
                $derivedStatus = 'SuppressedMissingCommenterScores'
            }
            elseif ($comparisonStatus -eq 'NotAvailableNoExactPopulationMetric') {
                $derivedStatus = 'NotAvailableNoExactPopulationMetric'
            }

            $metricResults += [pscustomobject][ordered]@{
                metric_id = $metricId
                response_field = $responseField
                population_metric = $populationMetricName
                event_minimum = $eventMinimum
                event_maximum = $eventMaximum
                commenter_scored_response_count = $scoredCount
                commenter_missing_score_count = $missingScoreCount
                commenter_event_count = $eventCount
                commenter_event_rate_pct = $commenterEventRate
                commenter_rate_denominator_label = "commenter responses with a nonmissing $responseField score"
                population_event_rate_pct = $populationEventRate
                population_rate_denominator_label = 'vendor rolling population aggregate metric denominator'
                denominator_alignment_status = if ($missingScoreCount -eq 0) { 'CommenterMetricComplete' } else { 'CommenterMetricMissingScores' }
                commenter_minus_population_percentage_points = $populationGap
                material_gap = $materialGap
                negative_experience_overrepresented_among_commenters = $negativeOverrepresentation
                population_event_count_reconstruction_error = $reconstructionError
                reconstructed_population_event_count = $reconstructedPopulationEventCount
                derived_non_comment_response_count = $nonCommentCount
                derived_non_comment_event_count = $derivedNonCommentEventCount
                derived_non_comment_event_rate_pct = $derivedNonCommentEventRate
                commenter_minus_non_comment_percentage_points = $nonCommentGap
                comparison_status = $comparisonStatus
                derived_non_comment_status = $derivedStatus
            }
        }

        $restaurantIssues = @($restaurantIssues | Sort-Object -Unique)
        $restaurantStatus = if ($restaurantIssues.Count -gt 0) {
            $hasDataQualityIssue = $true
            'DataQualityReview'
        }
        elseif ($commenterCount -eq 0) {
            'InsufficientData'
        }
        else {
            'Ready'
        }
        $restaurantResults += [pscustomobject][ordered]@{
            restaurant_id = $restaurantId
            status = $restaurantStatus
            population_response_count = $populationCount
            commenter_response_count = $commenterCount
            comment_coverage_pct = $coverage
            comment_coverage_status = $coverageStatus
            data_quality_issues = $restaurantIssues
            metrics = $metricResults
        }
    }

    $status = if ($hasDataQualityIssue) {
        'DataQualityReview'
    }
    elseif (-not $hasUsableCommenterData) {
        'InsufficientData'
    }
    else {
        'Ready'
    }
    return [pscustomobject][ordered]@{
        schema_version = 'gss-commenter-lens/v1'
        status = $status
        reason_code = if ($status -eq 'DataQualityReview') { 'commenter_lens_data_quality_review' } elseif ($status -eq 'InsufficientData') { 'no_commenter_rows_in_window' } else { 'commenter_lens_ready' }
        reason = if ($status -eq 'DataQualityReview') { 'One or more aggregate commenter results require data-quality review.' } elseif ($status -eq 'InsufficientData') { 'No comment-bearing responses were available in the visit-date reporting window.' } elseif ($ExactPartitionAlignmentVerified) { 'Aggregate commenter results and explicitly aligned cross-source comparisons are ready for human review.' } else { 'Aggregate commenter-only counts and rates are ready; cross-source coverage and comparisons are suppressed because exact partition alignment is unverified.' }
        scope_label = [string]$lensPolicy.scope_label
        population_modeling_status = [string]$script:GssAnalysisPolicy.modeling.status
        reporting_window = [pscustomobject][ordered]@{
            weeks = $windowWeeks
            start_date = $windowStart.ToString('yyyy-MM-dd')
            end_date = $windowEnd.ToString('yyyy-MM-dd')
            inclusive = $true
            commenter_date_basis = 'guest_detail_visit_date'
            population_date_basis = 'rolling_workbook_reporting_period'
            exact_partition_alignment_verified = [bool]$ExactPartitionAlignmentVerified
            source_alignment = if ($ExactPartitionAlignmentVerified) {
                'exact commenter/population reporting partition alignment explicitly verified by the caller'
            }
            else {
                'visit-date commenter window compared with the population rolling aggregate through the reporting date; not exact source-report-week row alignment'
            }
        }
        source_design = [pscustomobject][ordered]@{
            population_scores_source = [string]$script:GssAnalysisPolicy.source_design.population_scores_source
            population_raw_rows_available = [bool]$script:GssAnalysisPolicy.source_design.population_raw_rows_available
            row_level_scores_source = [string]$script:GssAnalysisPolicy.source_design.row_level_scores_source
            non_comment_row_level_scores_available = [bool]$script:GssAnalysisPolicy.source_design.non_comment_row_level_scores_available
            commenter_rows_population_representative = [bool]$script:GssAnalysisPolicy.source_design.commenter_rows_population_representative
        }
        claims = [pscustomobject][ordered]@{
            population_prevalence_claimed = $false
            statistical_significance_claimed = $false
            causal_driver_claimed = $false
            individual_prediction_produced = $false
        }
        row_level_data_persisted = $false
        restaurants = $restaurantResults
        limitations = @(
            'Raw scores are available only for surveys with comments.',
            'Commenter score distributions are not representative estimates for all guests.',
            'Population aggregate rates remain separately labeled and are not treated as aligned with commenter rows unless exact partition alignment is explicitly verified.',
            'Commenter event rates use only comment-providing responses with a nonmissing metric score; population rates retain the vendor aggregate denominator.',
            'Cross-source coverage and commenter-minus-population gaps are suppressed unless exact commenter/population partition alignment is explicitly verified.',
            'Derived non-comment counts and rates are suppressed unless exact commenter/population partition alignment is explicitly verified.',
            'Commenter rows are aligned by guest-detail visit date, not by an asserted exact source-report-week row match.',
            'Stable vendor response IDs are not consistently available across exports; deterministic content-based deduplication retains a low residual risk of false matches or missed matches.'
        )
    }
}

function Get-GssCommenterLensFallback {
    param(
        [Parameter(Mandatory)][datetime]$ReportingDate,
        [ValidateSet('InsufficientData', 'DataQualityReview')][string]$Status = 'DataQualityReview',
        [Parameter(Mandatory)][string]$Reason
    )

    Assert-GssSundayReportingDate -ReportingDate $ReportingDate
    $windowWeeks = [int]$script:GssAnalysisPolicy.commenter_lens.reporting_window_weeks
    $windowEnd = $ReportingDate.Date
    $windowStart = $windowEnd.AddDays(-(($windowWeeks * 7) - 1))
    return [pscustomobject][ordered]@{
        schema_version = 'gss-commenter-lens/v1'
        status = $Status
        reason_code = 'commenter_detail_inventory_unavailable'
        reason = $Reason
        scope_label = [string]$script:GssAnalysisPolicy.commenter_lens.scope_label
        population_modeling_status = [string]$script:GssAnalysisPolicy.modeling.status
        reporting_window = [pscustomobject][ordered]@{
            weeks = $windowWeeks
            start_date = $windowStart.ToString('yyyy-MM-dd')
            end_date = $windowEnd.ToString('yyyy-MM-dd')
            inclusive = $true
            commenter_date_basis = 'guest_detail_visit_date'
            population_date_basis = 'rolling_workbook_reporting_period'
            exact_partition_alignment_verified = $false
            source_alignment = 'visit-date commenter window compared with the population rolling aggregate through the reporting date; not exact source-report-week row alignment'
        }
        source_design = [pscustomobject][ordered]@{
            population_scores_source = [string]$script:GssAnalysisPolicy.source_design.population_scores_source
            population_raw_rows_available = [bool]$script:GssAnalysisPolicy.source_design.population_raw_rows_available
            row_level_scores_source = [string]$script:GssAnalysisPolicy.source_design.row_level_scores_source
            non_comment_row_level_scores_available = [bool]$script:GssAnalysisPolicy.source_design.non_comment_row_level_scores_available
            commenter_rows_population_representative = [bool]$script:GssAnalysisPolicy.source_design.commenter_rows_population_representative
        }
        claims = [pscustomobject][ordered]@{
            population_prevalence_claimed = $false
            statistical_significance_claimed = $false
            causal_driver_claimed = $false
            individual_prediction_produced = $false
        }
        row_level_data_persisted = $false
        restaurants = @()
        limitations = @(
            $Reason,
            'Raw scores are available only for surveys with comments.',
            'Commenter rows are aligned by guest-detail visit date, not by an asserted exact source-report-week row match.',
            'Stable vendor response IDs are not consistently available across exports; deterministic content-based deduplication retains a low residual risk of false matches or missed matches.'
        )
    }
}

function ConvertTo-GssCommenterLensCsv {
    param([Parameter(Mandatory)][object]$CommenterLens)

    $rows = @()
    foreach ($restaurant in @($CommenterLens.restaurants)) {
        foreach ($metric in @($restaurant.metrics)) {
            $rows += [pscustomobject][ordered]@{
                status = [string]$CommenterLens.status
                scope_label = [string]$CommenterLens.scope_label
                reporting_window_start = [string]$CommenterLens.reporting_window.start_date
                reporting_window_end = [string]$CommenterLens.reporting_window.end_date
                exact_partition_alignment_verified = [bool]$CommenterLens.reporting_window.exact_partition_alignment_verified
                restaurant_id = [string]$restaurant.restaurant_id
                restaurant_status = [string]$restaurant.status
                population_response_count = $restaurant.population_response_count
                commenter_response_count = $restaurant.commenter_response_count
                comment_coverage_pct = $restaurant.comment_coverage_pct
                comment_coverage_status = [string]$restaurant.comment_coverage_status
                metric_id = [string]$metric.metric_id
                response_field = [string]$metric.response_field
                population_metric = [string]$metric.population_metric
                commenter_scored_response_count = $metric.commenter_scored_response_count
                commenter_missing_score_count = $metric.commenter_missing_score_count
                commenter_event_count = $metric.commenter_event_count
                commenter_event_rate_pct = $metric.commenter_event_rate_pct
                commenter_rate_denominator_label = [string]$metric.commenter_rate_denominator_label
                population_event_rate_pct = $metric.population_event_rate_pct
                population_rate_denominator_label = [string]$metric.population_rate_denominator_label
                denominator_alignment_status = [string]$metric.denominator_alignment_status
                commenter_minus_population_percentage_points = $metric.commenter_minus_population_percentage_points
                material_gap = $metric.material_gap
                reconstructed_population_event_count = $metric.reconstructed_population_event_count
                derived_non_comment_response_count = $metric.derived_non_comment_response_count
                derived_non_comment_event_count = $metric.derived_non_comment_event_count
                derived_non_comment_event_rate_pct = $metric.derived_non_comment_event_rate_pct
                commenter_minus_non_comment_percentage_points = $metric.commenter_minus_non_comment_percentage_points
                comparison_status = [string]$metric.comparison_status
                derived_non_comment_status = [string]$metric.derived_non_comment_status
            }
        }
    }
    if ($rows.Count -eq 0) {
        return ($script:GssCommenterLensCsvColumns -join ',')
    }
    return (@($rows | ConvertTo-Csv -NoTypeInformation) -join [Environment]::NewLine)
}

function Assert-GssCommenterLensCsvContract {
    param([Parameter(Mandatory)][string]$CsvText)

    $firstLine = @($CsvText -split '\r?\n', 2)[0]
    $actualColumns = @($firstLine -split ',' | ForEach-Object { ([string]$_).Trim().Trim('"') })
    if ($actualColumns.Count -ne $script:GssCommenterLensCsvColumns.Count) {
        throw 'GSS commenter-lens CSV contract violation: header column count is invalid.'
    }
    for ($index = 0; $index -lt $script:GssCommenterLensCsvColumns.Count; $index++) {
        if ([string]$actualColumns[$index] -cne [string]$script:GssCommenterLensCsvColumns[$index]) {
            throw "GSS commenter-lens CSV contract violation: unsupported column '$($actualColumns[$index])'."
        }
    }
    if ($CsvText -match '(?i)response_hash|guest_first_name|guest_last_name|sanitized_text|source_path|source_row|reservation_time|individual_prediction') {
        throw 'GSS commenter-lens CSV contract violation: row-level fields are forbidden.'
    }
}

function Get-GssFeedbackSelection {
    param(
        [Parameter(Mandatory)][object]$Inventory,
        [Parameter(Mandatory)][object]$Ledger,
        [Parameter(Mandatory)][datetime]$ReportingDate
    )

    $reportingDateText = $ReportingDate.ToString('yyyy-MM-dd')
    $ledgerByHash = @{}
    foreach ($entry in @($Ledger.entries)) { $ledgerByHash[[string]$entry.response_hash] = $entry }
    $isInitialLedger = ($ledgerByHash.Count -eq 0)
    $currentPath = $Inventory.CurrentWorkbook.PortablePath
    $selected = @()
    foreach ($response in @($Inventory.UniqueResponses | Sort-Object ResponseHash)) {
        if ($response.VisitDate.Date -gt $ReportingDate.Date) { continue }
        $existing = $ledgerByHash[$response.ResponseHash]
        $inCurrentAttachment = @($response.SourcePaths | Where-Object { $_ -eq $currentPath }).Count -gt 0
        if ($existing) {
            if ($existing.first_seen_package_id -ne 'baseline' -and $existing.first_seen_reporting_date -eq $reportingDateText) { $selected += $response }
        }
        elseif (-not $isInitialLedger -or $inCurrentAttachment) {
            $selected += $response
        }
    }
    $canonicalHashes = @($selected.ResponseHash | Sort-Object -Unique)
    $fingerprintMaterial = "gss-feedback-selection/v1`n" + ($canonicalHashes -join "`n")
    return [pscustomobject]@{
        Responses = $selected
        ResponseHashes = $canonicalHashes
        Fingerprint = Get-GssStringSha256 $fingerprintMaterial
        LedgerByHash = $ledgerByHash
        IsInitialLedger = $isInitialLedger
    }
}

function New-GssSanitizedFeedback {
    param(
        [Parameter(Mandatory)][object]$Inventory,
        [Parameter(Mandatory)][object]$Ledger,
        [Parameter(Mandatory)][string]$PackageId,
        [Parameter(Mandatory)][datetime]$ReportingDate,
        [object]$FeedbackSelection
    )

    $reportingDateText = $ReportingDate.ToString('yyyy-MM-dd')
    $selection = if ($FeedbackSelection) { $FeedbackSelection } else { Get-GssFeedbackSelection -Inventory $Inventory -Ledger $Ledger -ReportingDate $ReportingDate }
    $ledgerByHash = $selection.LedgerByHash
    $isInitialLedger = [bool]$selection.IsInitialLedger
    $currentPath = $Inventory.CurrentWorkbook.PortablePath
    $selected = @($selection.Responses)

    $knownNames = @(Get-GssKnownGuestNames $Inventory.AllResponseInstances)
    $internalCards = @()
    $piiRedactions = 0
    $piiFailures = @()
    foreach ($response in $selected) {
        $protected = Protect-GssFeedbackText -Text $response.Text -KnownNames $knownNames
        $piiRedactions += $protected.RedactionCount
        if (-not $protected.PiiScanPassed) { $piiFailures += $protected.RemainingPiiTypes }
        $themes = @(Get-GssFeedbackThemes -Response $response -SanitizedText $protected.Text)
        $unknownThemes = @($themes | Where-Object { $_ -notin $script:GssThemeNames })
        if ($unknownThemes.Count -gt 0) { throw "Unknown guest-feedback theme: $($unknownThemes[0])" }
        $internalCards += [pscustomobject]@{
            EvidenceId = 'feedback-' + $response.ResponseHash.Substring(0, 16)
            ResponseHash = $response.ResponseHash
            RestaurantId = $response.RestaurantId
            Restaurant = $response.Restaurant
            VisitDate = $response.VisitDate.ToString('yyyy-MM-dd')
            ThemeIds = $themes
            Sentiment = Get-GssFeedbackSentiment -Response $response -SanitizedText $protected.Text
            DoNotContact = $response.DoNotContact
            SanitizedText = $protected.Text
        }
    }

    $themes = @()
    foreach ($restaurantGroup in @($internalCards | Group-Object RestaurantId)) {
        $restaurantResponses = @($restaurantGroup.Group | Sort-Object ResponseHash -Unique)
        $denominatorCount = $restaurantResponses.Count
        $denominatorVisitStart = if ($denominatorCount -gt 0) { ($restaurantResponses.VisitDate | Sort-Object | Select-Object -First 1) } else { $null }
        $denominatorVisitEnd = if ($denominatorCount -gt 0) { ($restaurantResponses.VisitDate | Sort-Object -Descending | Select-Object -First 1) } else { $null }
        foreach ($themeName in $script:GssThemeNames) {
            $matching = @($restaurantResponses | Where-Object { $_.ThemeIds -contains $themeName })
            if ($matching.Count -lt [int]$script:GssAnalysisPolicy.guest_feedback.minimum_unique_responses_per_theme) { continue }
            $slug = $themeName.Replace('/', '-').Replace(' ', '-')
            $themes += [pscustomobject][ordered]@{
                theme_id = "theme-$($restaurantGroup.Name)-$slug"
                restaurant_id = $restaurantGroup.Name
                category = $themeName
                unique_response_count = $matching.Count
                denominator_response_count = $denominatorCount
                visit_date_start = $denominatorVisitStart
                visit_date_end = $denominatorVisitEnd
                categories_are_non_exclusive = [bool]$script:GssAnalysisPolicy.guest_feedback.theme_categories_are_non_exclusive
                concern_count = @($matching | Where-Object { $_.Sentiment -in @('concern', 'mixed') }).Count
                positive_count = @($matching | Where-Object { $_.Sentiment -in @('positive', 'mixed') }).Count
                do_not_contact_count = @($matching | Where-Object { $_.DoNotContact }).Count
                quotable_evidence_ids = @($matching | Where-Object { -not $_.DoNotContact -and $_.SanitizedText } | Select-Object -ExpandProperty EvidenceId)
            }
        }
    }

    $publicCards = @()
    foreach ($card in @($internalCards | Where-Object { -not $_.DoNotContact })) {
        $restaurantDisplayName = Get-GssRestaurantDisplayName $card.Restaurant
        $publicCards += [pscustomobject][ordered]@{
            evidence_id = $card.EvidenceId
            restaurant_id = $card.RestaurantId
            restaurant = $restaurantDisplayName
            restaurant_name = $restaurantDisplayName
            source_entity = $card.Restaurant
            kind = 'guest_feedback_response'
            finding_type = 'response'
            visit_date = $card.VisitDate
            theme_ids = @($themes | Where-Object { $_.restaurant_id -eq $card.RestaurantId -and $card.ThemeIds -contains $_.category } | Select-Object -ExpandProperty theme_id)
            sentiment = $card.Sentiment
            quote_allowed = $true
            outreach_allowed = $true
            sanitized_text = $card.SanitizedText
            display_text = $card.SanitizedText
        }
    }

    $newEntries = @($Ledger.entries)
    foreach ($response in $Inventory.UniqueResponses) {
        if ($response.VisitDate.Date -gt $ReportingDate.Date) { continue }
        if ($ledgerByHash.ContainsKey($response.ResponseHash)) { continue }
        $inCurrentAttachment = @($response.SourcePaths | Where-Object { $_ -eq $currentPath }).Count -gt 0
        $firstPackage = if ($isInitialLedger -and -not $inCurrentAttachment) { 'baseline' } else { $PackageId }
        $firstDate = if ($isInitialLedger -and -not $inCurrentAttachment) { $response.VisitDate.ToString('yyyy-MM-dd') } else { $reportingDateText }
        $newEntries += [pscustomobject][ordered]@{
            response_hash = $response.ResponseHash
            first_seen_reporting_date = $firstDate
            first_seen_package_id = $firstPackage
        }
    }
    $nextLedger = [pscustomobject][ordered]@{
        schema_version = $script:GssFeedbackLedgerVersion
        entries = @($newEntries | Sort-Object response_hash)
    }

    $dateStart = if ($selected.Count -gt 0) { ($selected.VisitDate | Measure-Object -Minimum).Minimum.ToString('yyyy-MM-dd') } else { $null }
    $dateEnd = if ($selected.Count -gt 0) { ($selected.VisitDate | Measure-Object -Maximum).Maximum.ToString('yyyy-MM-dd') } else { $null }
    return [pscustomobject]@{
        ResponseCount = $selected.Count
        VisitDateStart = $dateStart
        VisitDateEnd = $dateEnd
        Cards = $publicCards
        InternalCards = $internalCards
        Themes = $themes
        SelectionFingerprint = [string]$selection.Fingerprint
        NextLedger = $nextLedger
        Privacy = [pscustomobject][ordered]@{
            pii_scan_passed = ($piiFailures.Count -eq 0)
            pii_scan_scope = 'risk-reduced analysis text only; raw detail attachment is excluded'
            pii_redaction_count = $piiRedactions
            do_not_contact_response_count = @($internalCards | Where-Object { $_.DoNotContact }).Count
            do_not_contact_text_excluded = $true
            guest_name_fields_excluded = $true
            raw_detail_contains_personal_data = $true
            package_contains_personal_data = $true
            classification = $script:GssRestrictedClassification
            analysis_description = 'risk-reduced'
            remaining_pii_types = @($piiFailures | Select-Object -Unique)
        }
    }
}

function Get-GssDeterministicPackageId {
    param(
        [string]$ReportingDate,
        [object[]]$SourceDescriptors,
        [Parameter(Mandatory)][string]$FeedbackSelectionFingerprint,
        [string]$SchemaVersion = $script:GssEmailPackageSchemaVersion,
        [string]$PolicyVersion = $script:GssAnalysisPolicyVersion
    )

    $normalizedSources = @($SourceDescriptors | ForEach-Object {
        $portablePath = if (-not [string]::IsNullOrWhiteSpace([string]$_.source_path)) { [string]$_.source_path } else { [string]$_.path }
        if ([string]::IsNullOrWhiteSpace($portablePath)) { throw "Package source descriptor is missing its portable path: $($_.role)" }
        if ($null -eq $_.byte_size) { throw "Package source descriptor is missing its byte size: $($_.role)" }
        [pscustomobject]@{
            role = [string]$_.role
            path = $portablePath.Replace('\', '/')
            byte_size = [long]$_.byte_size
            sha256 = ([string]$_.sha256).ToLowerInvariant()
        }
    })
    $sourceText = @($normalizedSources |
        Sort-Object @{ Expression = { $_.role.ToLowerInvariant() } }, @{ Expression = { $_.path.ToLowerInvariant() } } |
        ForEach-Object { "$($_.role):$($_.path):$($_.byte_size):$($_.sha256)" }) -join "`n"
    if ($FeedbackSelectionFingerprint -notmatch '^[a-f0-9]{64}$') { throw 'Feedback selection fingerprint must be a lowercase SHA-256 hash.' }
    $material = "$SchemaVersion`n$PolicyVersion`n$ReportingDate`nfeedback_selection:$FeedbackSelectionFingerprint`n$sourceText"
    return 'gss-' + $ReportingDate.Replace('-', '') + '-' + (Get-GssStringSha256 $material).Substring(0, 20)
}

function Get-GssSourceDescriptor {
    param([string]$Role, [string]$Path, [string]$FolderPath)

    $resolved = Resolve-GssDropboxPath -Path $Path -FolderPath $FolderPath -RequireFile
    $item = Get-Item -LiteralPath $resolved
    if ($item.Length -le 0) { throw "Attachment is empty or not fully synced: $Role" }
    return [pscustomobject][ordered]@{
        role = $Role
        source_path = ConvertTo-GssDropboxRelativePath -Path $resolved -FolderPath $FolderPath
        source_full_path = $resolved
        byte_size = [long]$item.Length
        sha256 = Get-GssSha256 $resolved
    }
}

function Get-GssPortableArtifactDescriptor {
    param(
        [Parameter(Mandatory)][string]$PackagePath,
        [Parameter(Mandatory)][string]$Role,
        [Parameter(Mandatory)][string]$RelativePath
    )

    $artifactPath = Join-Path $PackagePath $RelativePath
    if (-not (Test-Path -LiteralPath $artifactPath -PathType Leaf)) {
        throw "Portable package artifact is missing: $Role"
    }
    $item = Get-Item -LiteralPath $artifactPath
    if ($item.Length -le 0) {
        throw "Portable package artifact is empty: $Role"
    }
    return [pscustomobject][ordered]@{
        role = $Role
        path = $RelativePath
        byte_size = [long]$item.Length
        sha256 = Get-GssSha256 $artifactPath
    }
}

function Assert-GssPortableArtifactInventory {
    param(
        [Parameter(Mandatory)][string]$PackagePath,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Artifacts
    )

    if (@($Artifacts).Count -ne $script:GssPortableArtifactPaths.Count) {
        throw "Existing package must contain exactly $($script:GssPortableArtifactPaths.Count) portable artifact records."
    }
    foreach ($expected in $script:GssPortableArtifactPaths.GetEnumerator()) {
        $artifactMatches = @($Artifacts | Where-Object { [string]$_.role -ceq [string]$expected.Key })
        if ($artifactMatches.Count -ne 1) {
            throw "Existing package must contain exactly one portable artifact for role '$($expected.Key)'."
        }
        $artifact = $artifactMatches[0]
        Assert-GssExactObjectSchema `
            -Value $artifact `
            -AllowedProperties @('role', 'path', 'byte_size', 'sha256') `
            -Label "Portable artifact '$($expected.Key)'"
        if ([string]$artifact.path -cne [string]$expected.Value) {
            throw "Existing package portable artifact path is invalid for role '$($expected.Key)'."
        }
        if ($null -eq $artifact.byte_size -or
            -not (Test-GssNativeFiniteNumber $artifact.byte_size) -or
            [double]$artifact.byte_size -le 0 -or
            [double]$artifact.byte_size -ne [math]::Truncate([double]$artifact.byte_size) -or
            [string]$artifact.sha256 -cnotmatch '^[a-f0-9]{64}$') {
            throw "Existing package portable artifact evidence is invalid for role '$($expected.Key)'."
        }
        $artifactPath = Join-Path $PackagePath ([string]$expected.Value)
        if (-not (Test-Path -LiteralPath $artifactPath -PathType Leaf)) {
            throw "Existing package portable artifact is missing: $($expected.Key)"
        }
        $item = Get-Item -LiteralPath $artifactPath
        if ([long]$item.Length -ne [long]$artifact.byte_size -or
            (Get-GssSha256 $artifactPath) -cne [string]$artifact.sha256) {
            throw "Existing package portable artifact does not match its manifest: $($expected.Key)"
        }
    }
}

function Assert-GssLoggedFileEvidence {
    param([object]$RunLog, [object[]]$Descriptors)

    $logged = @($RunLog.FileEvidence)
    if ($logged.Count -eq 0) {
        throw 'Run log does not contain the required file size and SHA-256 evidence.'
    }
    $loggedRoles = @('comparison_pdf', 'rolling_workbook', 'prior_year_rolling_workbook', 'live_workbook')
    foreach ($descriptor in @($Descriptors | Where-Object { $_.role -in $loggedRoles })) {
        $entry = $logged | Where-Object { ([string]$_.Role) -eq $descriptor.role } | Select-Object -First 1
        if (-not $entry) { throw "Run log is missing file evidence for $($descriptor.role)." }
        if ([long]$entry.ByteSize -ne [long]$descriptor.byte_size -or ([string]$entry.Sha256).ToLowerInvariant() -ne $descriptor.sha256) {
            throw "File changed after the successful workbook run: $($descriptor.role)"
        }
        $loggedRelative = ([string]$entry.RelativePath).Replace('\', '/')
        if ($loggedRelative -ne $descriptor.source_path) {
            throw "Run log path does not match the current file for $($descriptor.role)."
        }
    }
}

function Format-GssEvidenceNumber {
    param([object]$Value, [switch]$Signed)
    if ($null -eq $Value) { return 'n/a' }
    $number = [math]::Round([double]$Value, 1, [System.MidpointRounding]::AwayFromZero)
    $format = if ($Signed) { '+0.0;-0.0;0.0' } else { '0.0' }
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:$format}", $number)
}

function ConvertTo-GssMetricEvidenceCard {
    param([object]$Item, [string]$Kind)

    $restaurantDisplayName = Get-GssRestaurantDisplayName $Item.Entity
    $confidenceTier = if ($Item.PSObject.Properties['ConfidenceTier']) { [string]$Item.ConfidenceTier } else { 'Not scored' }
    $responseVolumeTier = if ($Item.PSObject.Properties['ResponseVolumeTier']) { [string]$Item.ResponseVolumeTier } else { $confidenceTier }
    $displayText = '{0}: Level: 13-week rolling result {1} ({2} response-volume tier, n={3}); Movement: {4} points versus the previous 13-week rolling window and {5} points versus prior year; Benchmark: {6} points versus all franchisees. Adjacent 13-week windows overlap by 12 weeks.' -f `
        $Item.Metric, (Format-GssEvidenceNumber $Item.Current), $responseVolumeTier, (Format-GssEvidenceNumber $Item.CurrentCount), (Format-GssEvidenceNumber $Item.ChangeVsPreviousRollingWindow -Signed), (Format-GssEvidenceNumber $Item.YoYImprovement -Signed), (Format-GssEvidenceNumber $Item.VsAllFranchisees -Signed)
    return [pscustomobject][ordered]@{
        evidence_id = $Item.EvidenceId
        restaurant_id = $Item.RestaurantId
        restaurant = $restaurantDisplayName
        restaurant_name = $restaurantDisplayName
        source_entity = $Item.Entity
        kind = $Kind
        finding_type = $Kind
        metric_key = $Item.RawMetric
        metric = $Item.Metric
        direction = if ($Item.LowerIsBetter) { 'lower_is_better' } else { 'higher_is_better' }
        finding_direction = $Item.CandidateDirection
        lower_is_better = [bool]$Item.LowerIsBetter
        rolling_value = $Item.Current
        response_count = $Item.CurrentCount
        response_volume_tier = $responseVolumeTier
        confidence_tier = $confidenceTier
        change_vs_previous_window = $Item.ChangeVsPreviousRollingWindow
        change_vs_prior_year = $Item.YoYImprovement
        vs_franchise = $Item.VsAllFranchisees
        level = [pscustomobject][ordered]@{
            rolling_weeks = [int]$script:GssAnalysisPolicy.rolling_windows.weeks
            value = $Item.Current
            response_count = $Item.CurrentCount
            response_volume_tier = $responseVolumeTier
            confidence_tier = $confidenceTier
        }
        movement = [pscustomobject][ordered]@{
            change_vs_previous_window = $Item.ChangeVsPreviousRollingWindow
            change_vs_prior_year = $Item.YoYImprovement
            adjacent_window_overlap_weeks = [int]$script:GssAnalysisPolicy.rolling_windows.adjacent_overlap_weeks
        }
        benchmark = [pscustomobject][ordered]@{
            vs_sorensen_total = $Item.VsSorensenTotal
            vs_all_franchisees = $Item.VsAllFranchisees
        }
        persistent = [bool]$Item.PersistentMovement
        corroboration = @($Item.Corroboration)
        display_text = $displayText
    }
}

function Test-GssNativeFiniteNumber {
    param([object]$Value)

    $isNativeNumber =
        $Value -is [byte] -or
        $Value -is [sbyte] -or
        $Value -is [int16] -or
        $Value -is [uint16] -or
        $Value -is [int32] -or
        $Value -is [uint32] -or
        $Value -is [int64] -or
        $Value -is [uint64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]
    if (-not $isNativeNumber) { return $false }

    $number = [double]$Value
    return -not [double]::IsNaN($number) -and -not [double]::IsInfinity($number)
}

function Get-GssRequiredEvidenceProperty {
    param(
        [Parameter(Mandatory)][object]$Card,
        [Parameter(Mandatory)][string]$PropertyName,
        [Parameter(Mandatory)][string]$CardLabel
    )

    $property = $Card.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        throw "GSS analysis evidence contract violation: $CardLabel is missing required property '$PropertyName'."
    }
    return $property.Value
}

function Assert-GssMetricEvidenceCardContract {
    param(
        [Parameter(Mandatory)][object]$Card,
        [Parameter(Mandatory)][string]$CardLabel
    )

    foreach ($propertyName in @('evidence_id', 'restaurant_id', 'metric_key', 'metric', 'response_volume_tier', 'confidence_tier')) {
        $value = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $propertyName -CardLabel $CardLabel
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            throw "GSS analysis evidence contract violation: $CardLabel has an empty '$propertyName'."
        }
    }

    $direction = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName 'direction' -CardLabel $CardLabel
    if ($direction -notin @('higher_is_better', 'lower_is_better')) {
        throw "GSS analysis evidence contract violation: $CardLabel has invalid direction '$direction'."
    }
    $lowerIsBetter = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName 'lower_is_better' -CardLabel $CardLabel
    if ($lowerIsBetter -isnot [bool]) {
        throw "GSS analysis evidence contract violation: $CardLabel property 'lower_is_better' must be Boolean."
    }
    if ([bool]$lowerIsBetter -ne ($direction -eq 'lower_is_better')) {
        throw "GSS analysis evidence contract violation: $CardLabel direction conflicts with 'lower_is_better'."
    }

    $rollingValue = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName 'rolling_value' -CardLabel $CardLabel
    if (-not (Test-GssNativeFiniteNumber $rollingValue)) {
        throw "GSS analysis evidence contract violation: $CardLabel property 'rolling_value' must be a native finite number."
    }
    $confidenceTier = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName 'confidence_tier' -CardLabel $CardLabel
    if ($confidenceTier -notin @('High', 'Developing', 'Low', 'Not scored')) {
        throw "GSS analysis evidence contract violation: $CardLabel has invalid confidence tier '$confidenceTier'."
    }
    if ([string]$Card.response_volume_tier -ne [string]$confidenceTier) {
        throw "GSS analysis evidence contract violation: $CardLabel response-volume tier conflicts with its compatibility confidence-tier alias."
    }
    foreach ($propertyName in @('response_count', 'change_vs_previous_window', 'change_vs_prior_year', 'vs_franchise')) {
        $value = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $propertyName -CardLabel $CardLabel
        if ($null -ne $value -and -not (Test-GssNativeFiniteNumber $value)) {
            throw "GSS analysis evidence contract violation: $CardLabel property '$propertyName' must be null or a native finite number."
        }
    }
    foreach ($sectionName in @('level', 'movement', 'benchmark')) {
        $section = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $sectionName -CardLabel $CardLabel
        if ($null -eq $section) {
            throw "GSS analysis evidence contract violation: $CardLabel property '$sectionName' must be an object."
        }
    }
    if ([int]$Card.level.rolling_weeks -ne 13 -or [int]$Card.movement.adjacent_window_overlap_weeks -ne 12) {
        throw "GSS analysis evidence contract violation: $CardLabel must identify 13-week windows with 12 weeks of adjacent overlap."
    }
    if ([string]$Card.level.confidence_tier -ne [string]$Card.confidence_tier) {
        throw "GSS analysis evidence contract violation: $CardLabel level confidence tier conflicts with the top-level value."
    }
    if ([string]$Card.level.response_volume_tier -ne [string]$Card.response_volume_tier) {
        throw "GSS analysis evidence contract violation: $CardLabel level response-volume tier conflicts with the top-level value."
    }
}

function Assert-GssThemeEvidenceCardContract {
    param(
        [Parameter(Mandatory)][object]$Card,
        [Parameter(Mandatory)][string]$CardLabel
    )

    foreach ($propertyName in @('theme_id', 'restaurant_id', 'category')) {
        $value = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $propertyName -CardLabel $CardLabel
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            throw "GSS analysis evidence contract violation: $CardLabel has an empty '$propertyName'."
        }
    }

    foreach ($propertyName in @('unique_response_count', 'denominator_response_count', 'concern_count', 'positive_count', 'do_not_contact_count')) {
        $value = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $propertyName -CardLabel $CardLabel
        if (-not (Test-GssNativeFiniteNumber $value) -or [double]$value -lt 0 -or [double]$value -ne [math]::Truncate([double]$value)) {
            throw "GSS analysis evidence contract violation: $CardLabel property '$propertyName' must be a nonnegative native integer."
        }
    }
    if ([int]$Card.unique_response_count -gt [int]$Card.denominator_response_count) {
        throw "GSS analysis evidence contract violation: $CardLabel theme numerator exceeds its denominator."
    }
    foreach ($propertyName in @('visit_date_start', 'visit_date_end', 'categories_are_non_exclusive')) {
        $null = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $propertyName -CardLabel $CardLabel
    }
    if ($Card.categories_are_non_exclusive -isnot [bool] -or -not [bool]$Card.categories_are_non_exclusive) {
        throw "GSS analysis evidence contract violation: $CardLabel must identify theme categories as non-exclusive."
    }
}

function Assert-GssCommenterLensContract {
    param([Parameter(Mandatory)][object]$CommenterLens)

    Assert-GssExactObjectSchema `
        -Value $CommenterLens `
        -AllowedProperties @(
            'schema_version',
            'status',
            'reason_code',
            'reason',
            'scope_label',
            'population_modeling_status',
            'reporting_window',
            'source_design',
            'claims',
            'row_level_data_persisted',
            'restaurants',
            'limitations'
        ) `
        -Label 'GSS commenter lens'
    Assert-GssExactObjectSchema `
        -Value $CommenterLens.reporting_window `
        -AllowedProperties @(
            'weeks',
            'start_date',
            'end_date',
            'inclusive',
            'commenter_date_basis',
            'population_date_basis',
            'exact_partition_alignment_verified',
            'source_alignment'
        ) `
        -Label 'GSS commenter-lens reporting window'
    Assert-GssExactObjectSchema `
        -Value $CommenterLens.source_design `
        -AllowedProperties @(
            'population_scores_source',
            'population_raw_rows_available',
            'row_level_scores_source',
            'non_comment_row_level_scores_available',
            'commenter_rows_population_representative'
        ) `
        -Label 'GSS commenter-lens source design'
    Assert-GssExactObjectSchema `
        -Value $CommenterLens.claims `
        -AllowedProperties @(
            'population_prevalence_claimed',
            'statistical_significance_claimed',
            'causal_driver_claimed',
            'individual_prediction_produced'
        ) `
        -Label 'GSS commenter-lens claims'
    foreach ($propertyName in @('schema_version', 'status', 'reason_code', 'reason', 'scope_label', 'population_modeling_status', 'row_level_data_persisted')) {
        Assert-GssScalarValue -Value $CommenterLens.$propertyName -Label "GSS commenter-lens property '$propertyName'"
    }
    foreach ($propertyName in @('weeks', 'start_date', 'end_date', 'inclusive', 'commenter_date_basis', 'population_date_basis', 'exact_partition_alignment_verified', 'source_alignment')) {
        Assert-GssScalarValue -Value $CommenterLens.reporting_window.$propertyName -Label "GSS commenter-lens reporting-window property '$propertyName'"
    }
    foreach ($propertyName in @('population_scores_source', 'population_raw_rows_available', 'row_level_scores_source', 'non_comment_row_level_scores_available', 'commenter_rows_population_representative')) {
        Assert-GssScalarValue -Value $CommenterLens.source_design.$propertyName -Label "GSS commenter-lens source-design property '$propertyName'"
    }
    foreach ($propertyName in @('population_prevalence_claimed', 'statistical_significance_claimed', 'causal_driver_claimed', 'individual_prediction_produced')) {
        Assert-GssScalarValue -Value $CommenterLens.claims.$propertyName -Label "GSS commenter-lens claims property '$propertyName'"
    }
    foreach ($limitation in @($CommenterLens.limitations)) {
        if ($limitation -isnot [string]) {
            throw 'GSS commenter-lens contract violation: every limitation must be a scalar string.'
        }
    }

    if ([string]$CommenterLens.schema_version -ne 'gss-commenter-lens/v1' -or
        [string]$CommenterLens.status -notin @('Ready', 'InsufficientData', 'DataQualityReview') -or
        [string]::IsNullOrWhiteSpace([string]$CommenterLens.reason_code) -or
        [string]::IsNullOrWhiteSpace([string]$CommenterLens.reason) -or
        [string]$CommenterLens.scope_label -ne [string]$script:GssAnalysisPolicy.commenter_lens.scope_label -or
        [string]$CommenterLens.population_modeling_status -ne 'PopulationRawDataUnavailable') {
        throw 'GSS commenter-lens contract violation: schema, status, scope, or population-modeling status is invalid.'
    }
    $populationRawRowsProperty = $CommenterLens.source_design.PSObject.Properties['population_raw_rows_available']
    if ($null -eq $populationRawRowsProperty -or
        $populationRawRowsProperty.Value -isnot [bool] -or
        [bool]$populationRawRowsProperty.Value -or
        $CommenterLens.source_design.non_comment_row_level_scores_available -isnot [bool] -or
        [bool]$CommenterLens.source_design.non_comment_row_level_scores_available -or
        $CommenterLens.source_design.commenter_rows_population_representative -isnot [bool] -or
        [bool]$CommenterLens.source_design.commenter_rows_population_representative -or
        $CommenterLens.row_level_data_persisted -isnot [bool] -or
        [bool]$CommenterLens.row_level_data_persisted -or
        @(
            $CommenterLens.claims.population_prevalence_claimed,
            $CommenterLens.claims.statistical_significance_claimed,
            $CommenterLens.claims.causal_driver_claimed,
            $CommenterLens.claims.individual_prediction_produced
        ).Where({ $_ -isnot [bool] -or [bool]$_ }).Count -gt 0) {
        throw 'GSS commenter-lens contract violation: population raw-row availability, prohibited persistence, or analytical claims are enabled.'
    }
    $alignmentProperty = $CommenterLens.reporting_window.PSObject.Properties['exact_partition_alignment_verified']
    if ($null -eq $alignmentProperty -or $alignmentProperty.Value -isnot [bool]) {
        throw 'GSS commenter-lens contract violation: exact partition-alignment verification must be an explicit Boolean.'
    }
    $exactPartitionAlignmentVerified = [bool]$alignmentProperty.Value
    $restaurantIds = @()
    foreach ($restaurant in @($CommenterLens.restaurants)) {
        Assert-GssExactObjectSchema `
            -Value $restaurant `
            -AllowedProperties @(
                'restaurant_id',
                'status',
                'population_response_count',
                'commenter_response_count',
                'comment_coverage_pct',
                'comment_coverage_status',
                'data_quality_issues',
                'metrics'
            ) `
            -Label 'GSS commenter-lens restaurant'
        foreach ($propertyName in @(
            'restaurant_id',
            'status',
            'population_response_count',
            'commenter_response_count',
            'comment_coverage_pct',
            'comment_coverage_status'
        )) {
            Assert-GssScalarValue `
                -Value $restaurant.$propertyName `
                -AllowNull:($propertyName -in @('population_response_count', 'comment_coverage_pct')) `
                -Label "GSS commenter-lens restaurant property '$propertyName'"
        }
        foreach ($issue in @($restaurant.data_quality_issues)) {
            if ($issue -isnot [string]) {
                throw "GSS commenter-lens contract violation: restaurant '$($restaurant.restaurant_id)' data-quality issues must be scalar strings."
            }
        }
        if ([string]::IsNullOrWhiteSpace([string]$restaurant.restaurant_id) -or
            [string]$restaurant.status -notin @('Ready', 'InsufficientData', 'DataQualityReview')) {
            throw 'GSS commenter-lens contract violation: restaurant identity or status is invalid.'
        }
        if ([string]$restaurant.restaurant_id -in $restaurantIds) {
            throw "GSS commenter-lens contract violation: restaurant '$($restaurant.restaurant_id)' is duplicated."
        }
        $restaurantIds += [string]$restaurant.restaurant_id
        foreach ($propertyName in @('commenter_response_count')) {
            $value = $restaurant.$propertyName
            if (-not (Test-GssNativeFiniteNumber $value) -or
                [double]$value -lt 0 -or
                [double]$value -ne [math]::Truncate([double]$value)) {
                throw "GSS commenter-lens contract violation: restaurant '$($restaurant.restaurant_id)' property '$propertyName' must be a nonnegative native integer."
            }
        }
        if ([string]$restaurant.comment_coverage_status -notin @(
            'Ready',
            'SuppressedUnverifiedPartitionAlignment',
            'SuppressedInvalidCoverage'
        )) {
            throw "GSS commenter-lens contract violation: restaurant '$($restaurant.restaurant_id)' has an invalid comment-coverage status."
        }
        if (-not $exactPartitionAlignmentVerified -and
            (
                $null -ne $restaurant.comment_coverage_pct -or
                [string]$restaurant.comment_coverage_status -eq 'Ready'
            )) {
            throw "GSS commenter-lens contract violation: restaurant '$($restaurant.restaurant_id)' exposes comment coverage without verified partition alignment."
        }
        if ($null -ne $restaurant.population_response_count -and
            (
                -not (Test-GssNativeFiniteNumber $restaurant.population_response_count) -or
                [double]$restaurant.population_response_count -lt 0 -or
                [double]$restaurant.population_response_count -ne [math]::Truncate([double]$restaurant.population_response_count)
            )) {
            throw "GSS commenter-lens contract violation: restaurant '$($restaurant.restaurant_id)' has an invalid population response count."
        }
        if ($null -ne $restaurant.comment_coverage_pct -and
            (
                -not (Test-GssNativeFiniteNumber $restaurant.comment_coverage_pct) -or
                [double]$restaurant.comment_coverage_pct -lt 0 -or
                [double]$restaurant.comment_coverage_pct -gt 100
            )) {
            throw "GSS commenter-lens contract violation: restaurant '$($restaurant.restaurant_id)' has an invalid comment-coverage percentage."
        }
        $metricIds = @()
        foreach ($metric in @($restaurant.metrics)) {
            Assert-GssExactObjectSchema `
                -Value $metric `
                -AllowedProperties @(
                    'metric_id',
                    'response_field',
                    'population_metric',
                    'event_minimum',
                    'event_maximum',
                    'commenter_scored_response_count',
                    'commenter_missing_score_count',
                    'commenter_event_count',
                    'commenter_event_rate_pct',
                    'commenter_rate_denominator_label',
                    'population_event_rate_pct',
                    'population_rate_denominator_label',
                    'denominator_alignment_status',
                    'commenter_minus_population_percentage_points',
                    'material_gap',
                    'negative_experience_overrepresented_among_commenters',
                    'population_event_count_reconstruction_error',
                    'reconstructed_population_event_count',
                    'derived_non_comment_response_count',
                    'derived_non_comment_event_count',
                    'derived_non_comment_event_rate_pct',
                    'commenter_minus_non_comment_percentage_points',
                    'comparison_status',
                    'derived_non_comment_status'
                ) `
                -Label "GSS commenter-lens metric for restaurant '$($restaurant.restaurant_id)'"
            foreach ($propertyName in @(
                'metric_id',
                'response_field',
                'population_metric',
                'event_minimum',
                'event_maximum',
                'commenter_scored_response_count',
                'commenter_missing_score_count',
                'commenter_event_count',
                'commenter_event_rate_pct',
                'commenter_rate_denominator_label',
                'population_event_rate_pct',
                'population_rate_denominator_label',
                'denominator_alignment_status',
                'commenter_minus_population_percentage_points',
                'material_gap',
                'negative_experience_overrepresented_among_commenters',
                'population_event_count_reconstruction_error',
                'reconstructed_population_event_count',
                'derived_non_comment_response_count',
                'derived_non_comment_event_count',
                'derived_non_comment_event_rate_pct',
                'commenter_minus_non_comment_percentage_points',
                'comparison_status',
                'derived_non_comment_status'
            )) {
                Assert-GssScalarValue `
                    -Value $metric.$propertyName `
                    -AllowNull:($propertyName -in @(
                        'population_metric',
                        'commenter_event_rate_pct',
                        'population_event_rate_pct',
                        'commenter_minus_population_percentage_points',
                        'material_gap',
                        'negative_experience_overrepresented_among_commenters',
                        'population_event_count_reconstruction_error',
                        'reconstructed_population_event_count',
                        'derived_non_comment_response_count',
                        'derived_non_comment_event_count',
                        'derived_non_comment_event_rate_pct',
                        'commenter_minus_non_comment_percentage_points'
                    )) `
                    -Label "GSS commenter-lens metric '$($metric.metric_id)' property '$propertyName'"
            }
            $metricPolicyMatches = @(
                $script:GssAnalysisPolicy.commenter_lens.metrics |
                    Where-Object { [string]$_.id -ceq [string]$metric.metric_id }
            )
            if ($metricPolicyMatches.Count -ne 1) {
                throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' is not defined exactly once by policy."
            }
            if ([string]$metric.metric_id -in $metricIds) {
                throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' is duplicated for restaurant '$($restaurant.restaurant_id)'."
            }
            $metricIds += [string]$metric.metric_id
            $metricPolicy = $metricPolicyMatches[0]
            $expectedPopulationMetric = if ($null -eq $metricPolicy.population_metric) { $null } else { [string]$metricPolicy.population_metric }
            if ([string]$metric.response_field -cne [string]$metricPolicy.response_field -or
                $metric.population_metric -cne $expectedPopulationMetric -or
                [double]$metric.event_minimum -ne [double]$metricPolicy.event_minimum -or
                [double]$metric.event_maximum -ne [double]$metricPolicy.event_maximum) {
                throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' does not match its policy definition."
            }
            foreach ($countName in @('commenter_scored_response_count', 'commenter_missing_score_count', 'commenter_event_count')) {
                $value = $metric.$countName
                if (-not (Test-GssNativeFiniteNumber $value) -or
                    [double]$value -lt 0 -or
                    [double]$value -ne [math]::Truncate([double]$value)) {
                    throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' property '$countName' must be a nonnegative native integer."
                }
            }
            if ([int]$metric.commenter_event_count -gt [int]$metric.commenter_scored_response_count -or
                ([int]$metric.commenter_scored_response_count + [int]$metric.commenter_missing_score_count) -ne [int]$restaurant.commenter_response_count) {
                throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' denominators are inconsistent."
            }
            if ([int]$metric.commenter_missing_score_count -gt 0 -and
                (
                    $null -ne $metric.commenter_minus_population_percentage_points -or
                    [string]$metric.comparison_status -eq 'Ready'
                )) {
                throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' exposes a population gap with missing commenter scores."
            }
            if (-not $exactPartitionAlignmentVerified -and
                (
                    $null -ne $metric.commenter_minus_population_percentage_points -or
                    $null -ne $metric.material_gap -or
                    $null -ne $metric.negative_experience_overrepresented_among_commenters -or
                    $null -ne $metric.reconstructed_population_event_count -or
                    $null -ne $metric.population_event_count_reconstruction_error -or
                    [string]$metric.comparison_status -eq 'Ready'
                )) {
                throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' exposes a cross-source comparison without verified partition alignment."
            }
            if (-not $exactPartitionAlignmentVerified -and
                (
                    $null -ne $metric.derived_non_comment_response_count -or
                    $null -ne $metric.derived_non_comment_event_count -or
                    $null -ne $metric.derived_non_comment_event_rate_pct -or
                    $null -ne $metric.commenter_minus_non_comment_percentage_points -or
                    [string]$metric.derived_non_comment_status -eq 'Ready'
                )) {
                throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' exposes derived non-comment results without verified partition alignment."
            }
            foreach ($numericName in @(
                'commenter_event_rate_pct',
                'population_event_rate_pct',
                'commenter_minus_population_percentage_points',
                'population_event_count_reconstruction_error',
                'derived_non_comment_event_rate_pct',
                'commenter_minus_non_comment_percentage_points'
            )) {
                $value = $metric.$numericName
                if ($null -ne $value -and -not (Test-GssNativeFiniteNumber $value)) {
                    throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' property '$numericName' must be null or a native finite number."
                }
            }
            foreach ($nullableCountName in @(
                'reconstructed_population_event_count',
                'derived_non_comment_response_count',
                'derived_non_comment_event_count'
            )) {
                $value = $metric.$nullableCountName
                if ($null -ne $value -and
                    (
                        -not (Test-GssNativeFiniteNumber $value) -or
                        [double]$value -lt 0 -or
                        [double]$value -ne [math]::Truncate([double]$value)
                    )) {
                    throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' property '$nullableCountName' must be null or a nonnegative native integer."
                }
            }
            foreach ($nullableBooleanName in @('material_gap', 'negative_experience_overrepresented_among_commenters')) {
                $value = $metric.$nullableBooleanName
                if ($null -ne $value -and $value -isnot [bool]) {
                    throw "GSS commenter-lens contract violation: metric '$($metric.metric_id)' property '$nullableBooleanName' must be null or Boolean."
                }
            }
        }
        $expectedMetricIds = @($script:GssAnalysisPolicy.commenter_lens.metrics | ForEach-Object { [string]$_.id })
        if ($metricIds.Count -ne $expectedMetricIds.Count -or
            @($expectedMetricIds | Where-Object { $_ -notin $metricIds }).Count -gt 0) {
            throw "GSS commenter-lens contract violation: restaurant '$($restaurant.restaurant_id)' does not contain the exact policy metric set."
        }
    }
    $serialized = $CommenterLens | ConvertTo-Json -Depth 20 -Compress
    foreach ($forbiddenProperty in @(
        'response_hash',
        'guest_first_name',
        'guest_last_name',
        'sanitized_text',
        'source_path',
        'source_row',
        'reservation_time',
        'individual_prediction'
    )) {
        if ($serialized -match ('(?i)"' + [regex]::Escape($forbiddenProperty) + '"\s*:')) {
            throw "GSS commenter-lens contract violation: row-level property '$forbiddenProperty' is forbidden."
        }
    }
}

function Test-GssAnalysisEvidenceContract {
    param([Parameter(Mandatory)][object]$Analysis)

    foreach ($propertyName in @('metric_evidence', 'theme_evidence', 'evidence_cards', 'commenter_lens')) {
        if ($null -eq $Analysis.PSObject.Properties[$propertyName]) {
            throw "GSS analysis evidence contract violation: analysis is missing required property '$propertyName'."
        }
    }
    Assert-GssCommenterLensContract -CommenterLens $Analysis.commenter_lens

    $unifiedCards = @($Analysis.evidence_cards)
    foreach ($card in @($Analysis.metric_evidence)) {
        $evidenceId = [string](Get-GssRequiredEvidenceProperty -Card $card -PropertyName 'evidence_id' -CardLabel 'metric evidence card')
        Assert-GssMetricEvidenceCardContract -Card $card -CardLabel "metric evidence '$evidenceId'"
        $matches = @($unifiedCards | Where-Object { [string]$_.evidence_id -eq $evidenceId })
        if ($matches.Count -ne 1) {
            throw "GSS analysis evidence contract violation: metric evidence '$evidenceId' must appear exactly once in evidence_cards."
        }
        Assert-GssMetricEvidenceCardContract -Card $matches[0] -CardLabel "evidence_cards metric '$evidenceId'"
    }

    foreach ($card in @($Analysis.theme_evidence)) {
        $themeId = [string](Get-GssRequiredEvidenceProperty -Card $card -PropertyName 'theme_id' -CardLabel 'theme evidence card')
        Assert-GssThemeEvidenceCardContract -Card $card -CardLabel "theme evidence '$themeId'"
        $matches = @($unifiedCards | Where-Object { [string]$_.theme_id -eq $themeId })
        if ($matches.Count -ne 1) {
            throw "GSS analysis evidence contract violation: theme evidence '$themeId' must appear exactly once in evidence_cards."
        }
        Assert-GssThemeEvidenceCardContract -Card $matches[0] -CardLabel "evidence_cards theme '$themeId'"
    }

    return $true
}

function New-GssEvidencePreviewText {
    param([object]$Analysis, [object]$Feedback, [datetime]$ReportingDate)

    $lines = @()
    $lines += "Subject: GSS Survey Report - Week Ending $($ReportingDate.ToString('MMddyy'))"
    $lines += "CLASSIFICATION: $script:GssRestrictedClassification"
    $lines += 'Automatic sending is disabled. Before any manual send, confirm every recipient is authorized to receive restricted guest survey data.'
    $lines += ''
    $lines += 'All,'
    $lines += ''
    $reportingDateDisplay = $ReportingDate.ToString('MMMM d, yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    $lines += "The attached restricted package covers 13-week rolling results through $reportingDateDisplay. Adjacent 13-week windows overlap by 12 weeks; movements compare the current result with that adjacent window and the matching prior-year rolling window."
    foreach ($restaurant in @($Analysis.RestaurantFindings)) {
        $lines += ''
        $lines += "$($restaurant.Restaurant):"
        if (@($restaurant.Strengths).Count -eq 0 -and @($restaurant.Opportunities).Count -eq 0) {
            $lines += 'No movement met the reporting thresholds.'
        }
        foreach ($item in @($restaurant.Strengths)) {
            $volumeTier = if ($item.PSObject.Properties['ResponseVolumeTier']) { $item.ResponseVolumeTier } else { $item.ConfidenceTier }
            $lines += "Strength: $($item.Metric). Level: $(Format-GssEvidenceNumber $item.Current) ($volumeTier response-volume tier, n=$(Format-GssEvidenceNumber $item.CurrentCount)). Movement: $(Format-GssEvidenceNumber $item.ChangeVsPreviousRollingWindow -Signed) points versus the previous window and $(Format-GssEvidenceNumber $item.YoYImprovement -Signed) points versus prior year. Benchmark: $(Format-GssEvidenceNumber $item.VsAllFranchisees -Signed) points versus all franchisees."
        }
        foreach ($item in @($restaurant.Opportunities)) {
            $volumeTier = if ($item.PSObject.Properties['ResponseVolumeTier']) { $item.ResponseVolumeTier } else { $item.ConfidenceTier }
            $lines += "Opportunity: $($item.Metric). Level: $(Format-GssEvidenceNumber $item.Current) ($volumeTier response-volume tier, n=$(Format-GssEvidenceNumber $item.CurrentCount)). Movement: $(Format-GssEvidenceNumber $item.ChangeVsPreviousRollingWindow -Signed) points versus the previous window and $(Format-GssEvidenceNumber $item.YoYImprovement -Signed) points versus prior year. Benchmark: $(Format-GssEvidenceNumber $item.VsAllFranchisees -Signed) points versus all franchisees."
        }
    }
    $lines += ''
    $lines += "Population driver modeling: $($Analysis.Modeling.Status). $($Analysis.Modeling.Reason)"
    if ($null -ne $Analysis.CommenterLens) {
        $lines += "$($Analysis.CommenterLens.scope_label) - aggregate commenter lens:"
        foreach ($restaurant in @($Analysis.CommenterLens.restaurants)) {
            $coverageText = if ($null -eq $restaurant.comment_coverage_pct) {
                'coverage unavailable'
            }
            else {
                '{0:N1}% comment coverage' -f [double]$restaurant.comment_coverage_pct
            }
            $lines += "- $($restaurant.restaurant_id): $($restaurant.commenter_response_count) comment-bearing surveys of $($restaurant.population_response_count) population responses; $coverageText; status $($restaurant.status)."
            foreach ($metric in @($restaurant.metrics)) {
                $gapText = if ($null -eq $metric.commenter_minus_population_percentage_points) {
                    'population gap suppressed'
                }
                else {
                    '{0:+0.0;-0.0;0.0} percentage points versus the population aggregate' -f [double]$metric.commenter_minus_population_percentage_points
                }
                $lines += "  - $($metric.metric_id): $(Format-GssEvidenceNumber $metric.commenter_event_rate_pct)% among scored comment-providing guests (n=$($metric.commenter_scored_response_count)); $gapText; $($metric.comparison_status)."
            }
        }
        $lines += 'Commenter scores describe only guests who provided comments. Commenter rates use nonmissing commenter scores, while population rates retain the vendor aggregate denominator. The commenter window is aligned by guest-detail visit date, not by an asserted exact source-report-week row match. These results do not estimate population score prevalence, causal drivers, or individual outcomes.'
        if (-not [bool]$Analysis.CommenterLens.reporting_window.exact_partition_alignment_verified) {
            $lines += 'Derived non-comment counts and rates are suppressed because exact commenter/population reporting-partition alignment has not been verified.'
        }
    }
    $lines += ''
    if ($Feedback.ResponseCount -gt 0) {
        $lines += "New guest feedback among guests who provided comments: $($Feedback.ResponseCount) unique responses with visit dates from $($Feedback.VisitDateStart) through $($Feedback.VisitDateEnd). Themes are reported only when supported by at least $($script:GssAnalysisPolicy.guest_feedback.minimum_unique_responses_per_theme) unique comment-providing responses."
        foreach ($theme in @($Feedback.Themes)) {
            $lines += "- $($theme.restaurant_id) $($theme.category): $($theme.unique_response_count) of $($theme.denominator_response_count) unique comment-providing responses, visit dates $($theme.visit_date_start) through $($theme.visit_date_end). Theme categories are non-exclusive."
        }
    }
    else {
        $lines += 'New guest feedback among guests who provided comments: no previously unseen responses were found in the validated detail exports.'
    }
    $lines += ''
    $lines += 'Recommended follow-up: review the evidence-backed opportunities with the appropriate restaurant leaders and confirm local context before choosing an action.'
    $lines += ''
    $lines += 'Bottom line: this risk-reduced analysis requires human review. These directional comparisons identify items for review; they do not establish statistical significance or causation.'
    $lines += ''
    $lines += "Methodology: population score results are 13-week rolling aggregates, and adjacent windows overlap by 12 weeks. Raw row-level scores are supplied only for surveys with comments; commenter score distributions and themes are labeled among guests who provided comments and are not population-representative. Guest feedback is deduplicated across current and archived exports; published analysis text is risk-reduced through name, contact-pattern, unsafe-control, and do-not-contact handling. The raw detail attachment still contains personal data, so the entire package remains $script:GssRestrictedClassification. Feedback is described using its actual visit-date range rather than as an exact seven-day sample."
    return ($lines -join [Environment]::NewLine)
}

function ConvertTo-GssSimpleHtml {
    param([string]$Text)

    $encodedLines = @($Text -split '\r?\n' | ForEach-Object { [System.Net.WebUtility]::HtmlEncode($_) })
    $body = @()
    foreach ($line in $encodedLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { $body += '<br>' }
        elseif ($line.StartsWith('Subject:')) { $body += "<p><strong>$line</strong></p>" }
        elseif ($line.EndsWith(':')) { $body += "<h2>$line</h2>" }
        elseif ($line.StartsWith('- ')) { $body += "<p>&bull; $($line.Substring(2))</p>" }
        else { $body += "<p>$line</p>" }
    }
    return '<!doctype html><html><head><meta charset="utf-8"></head><body style="font-family:Arial,sans-serif;color:#222;line-height:1.4">' + ($body -join '') + '</body></html>'
}

function Test-GssExistingEmailPackage {
    param(
        [string]$PackagePath,
        [string]$PackageId,
        [object[]]$ExpectedSourceDescriptors,
        [Parameter(Mandatory)][string]$ExpectedFeedbackSelectionFingerprint
    )

    $readyPath = Join-Path $PackagePath 'READY'
    $manifestPath = Join-Path $PackagePath 'email_manifest.json'
    if (-not (Test-Path -LiteralPath $readyPath -PathType Leaf) -or -not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "Existing deterministic package is incomplete and was not overwritten: $PackageId"
    }
    $readyLines = @(Get-Content -LiteralPath $readyPath)
    if ($readyLines.Count -ne 2 -or $readyLines[0] -ne $script:GssEmailPackageSchemaVersion -or $readyLines[1] -ne $PackageId) {
        throw "Existing deterministic package READY marker does not match package ID: $PackageId"
    }
    $manifest = Read-GssUtf8NoBomFile $manifestPath | ConvertFrom-Json
    if ($manifest.schema_version -ne $script:GssEmailPackageSchemaVersion -or $manifest.package_id -ne $PackageId) {
        throw "Existing deterministic package manifest does not match package ID: $PackageId"
    }
    if ([string]$manifest.classification -ne $script:GssRestrictedClassification -or -not [bool]$manifest.package_contains_personal_data) {
        throw "Existing deterministic package is missing the restricted personal-data classification: $PackageId"
    }
    if ([bool]$manifest.distribution_controls.automatic_sending_enabled -or
        -not [bool]$manifest.distribution_controls.human_review_required -or
        -not [bool]$manifest.distribution_controls.restricted_recipient_review.required_before_manual_send) {
        throw "Existing deterministic package is missing required manual restricted-recipient review controls: $PackageId"
    }
    if ([string]$manifest.feedback_selection_sha256 -notmatch '^[a-f0-9]{64}$' -or [string]$manifest.feedback_selection_sha256 -ne $ExpectedFeedbackSelectionFingerprint) {
        throw "Existing deterministic package feedback selection does not match current package inputs: $PackageId"
    }
    $recomputedPackageId = Get-GssDeterministicPackageId `
        -ReportingDate ([string]$manifest.reporting.reporting_date) `
        -SourceDescriptors @($manifest.sources) `
        -FeedbackSelectionFingerprint ([string]$manifest.feedback_selection_sha256) `
        -SchemaVersion ([string]$manifest.schema_version) `
        -PolicyVersion ([string]$manifest.policy_version)
    if ($recomputedPackageId -ne $PackageId) {
        throw "Existing deterministic package ID does not match its attested source inventory: $PackageId"
    }
    foreach ($requiredName in @('analysis.json', 'commenter_lens.json', 'commenter_lens.csv', 'email_preview.txt', 'email_preview.html', 'RESTRICTED.txt', 'READY')) {
        $requiredPath = Join-Path $PackagePath $requiredName
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf) -or (Get-Item -LiteralPath $requiredPath).Length -le 0) {
            throw "Existing deterministic package is missing required output: $requiredName"
        }
    }
    Assert-GssPortableArtifactInventory -PackagePath $PackagePath -Artifacts @($manifest.portable_artifacts)
    if ([string]$manifest.analysis_path -cne 'analysis.json' -or
        [string]$manifest.commenter_lens_json_path -cne 'commenter_lens.json' -or
        [string]$manifest.commenter_lens_csv_path -cne 'commenter_lens.csv' -or
        [string]$manifest.text_preview_path -cne 'email_preview.txt' -or
        [string]$manifest.html_preview_path -cne 'email_preview.html' -or
        [string]$manifest.classification_notice_path -cne 'RESTRICTED.txt') {
        throw "Existing deterministic package portable output paths are invalid: $PackageId"
    }
    $analysis = Read-GssUtf8NoBomFile (Join-Path $PackagePath 'analysis.json') | ConvertFrom-Json
    if ($analysis.schema_version -ne $script:GssEmailPackageSchemaVersion -or $analysis.package_id -ne $PackageId) {
        throw "Existing deterministic package analysis does not match package ID: $PackageId"
    }
    if ([string]$analysis.classification -ne $script:GssRestrictedClassification -or -not [bool]$analysis.package_contains_personal_data) {
        throw "Existing deterministic package analysis is missing the restricted personal-data classification: $PackageId"
    }
    if ([string]$analysis.feedback_selection_sha256 -ne [string]$manifest.feedback_selection_sha256) {
        throw "Existing deterministic package analysis feedback selection does not match its manifest: $PackageId"
    }
    $commenterLens = Read-GssUtf8NoBomFile (Join-Path $PackagePath 'commenter_lens.json') | ConvertFrom-Json
    Assert-GssCommenterLensContract -CommenterLens $commenterLens
    $embeddedCommenterLens = $analysis.commenter_lens | ConvertTo-Json -Depth 20 -Compress
    $standaloneCommenterLens = $commenterLens | ConvertTo-Json -Depth 20 -Compress
    if ($embeddedCommenterLens -cne $standaloneCommenterLens) {
        throw "Existing deterministic package analysis and commenter-lens output disagree: $PackageId"
    }
    $commenterCsv = Read-GssUtf8NoBomFile (Join-Path $PackagePath 'commenter_lens.csv')
    Assert-GssCommenterLensCsvContract -CsvText $commenterCsv
    $null = Test-GssAnalysisEvidenceContract -Analysis $analysis
    $availableEvidenceIds = @(
        @($analysis.portfolio_evidence) + @($analysis.metric_evidence) + @($analysis.sanitized_feedback) |
            ForEach-Object { [string]$_.evidence_id } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )
    foreach ($evidenceId in @($manifest.evidence_ids)) {
        if ($evidenceId -notin $availableEvidenceIds) { throw "Existing package references unknown evidence ID: $evidenceId" }
    }
    $availableThemeIds = @($analysis.feedback_themes.theme_id | Where-Object { $_ } | Sort-Object -Unique)
    foreach ($themeId in @($manifest.theme_ids)) {
        if ($themeId -notin $availableThemeIds) { throw "Existing package references unknown theme ID: $themeId" }
    }
    $expectedSources = @($ExpectedSourceDescriptors | ForEach-Object {
        [pscustomobject]@{
            role = [string]$_.role
            path = ([string]$_.source_path).Replace('\', '/')
            byte_size = [long]$_.byte_size
            sha256 = ([string]$_.sha256).ToLowerInvariant()
        }
    })
    $actualSources = @($manifest.sources)
    if ($actualSources.Count -ne $expectedSources.Count) {
        throw "Existing package source inventory does not match current package inputs: $PackageId"
    }
    foreach ($expectedSource in $expectedSources) {
        $matches = @($actualSources | Where-Object {
            ([string]$_.role) -eq $expectedSource.role -and ([string]$_.path).Replace('\', '/') -eq $expectedSource.path
        })
        if ($matches.Count -ne 1) { throw "Existing package source inventory is missing or duplicated: $($expectedSource.role) $($expectedSource.path)" }
        $actualSource = $matches[0]
        if ([long]$actualSource.byte_size -ne $expectedSource.byte_size -or ([string]$actualSource.sha256).ToLowerInvariant() -ne $expectedSource.sha256) {
            throw "Existing package source evidence does not match current input: $($expectedSource.role) $($expectedSource.path)"
        }
    }
    $runLogSource = @($expectedSources | Where-Object role -eq 'run_log')
    if ($runLogSource.Count -ne 1 -or ([string]$manifest.source_log_path).Replace('\', '/') -ne $runLogSource[0].path) {
        throw "Existing package source_log_path does not match its run_log source: $PackageId"
    }
    $requiredAttachmentRoles = @('comparison_pdf', 'rolling_workbook', 'detail_workbook')
    if (@($manifest.attachments).Count -ne 3) {
        throw "Existing package must contain exactly three attachments: $PackageId"
    }
    foreach ($requiredRole in $requiredAttachmentRoles) {
        if (@($manifest.attachments | Where-Object role -eq $requiredRole).Count -ne 1) {
            throw "Existing package must contain exactly one attachment for role '$requiredRole': $PackageId"
        }
    }
    foreach ($attachment in @($manifest.attachments)) {
        $path = Join-Path $PackagePath ([string]$attachment.path).Replace('/', '\')
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Existing package attachment is missing: $($attachment.role)" }
        $item = Get-Item -LiteralPath $path
        if ([long]$item.Length -ne [long]$attachment.byte_size -or (Get-GssSha256 $path) -ne [string]$attachment.sha256) {
            throw "Existing package attachment does not match its manifest: $($attachment.role)"
        }
    }
    $detailAttachments = @($manifest.attachments | Where-Object role -eq 'detail_workbook')
    if ($detailAttachments.Count -ne 1 -or
        -not [bool]$detailAttachments[0].contains_personal_data -or
        [string]$detailAttachments[0].classification -ne $script:GssRestrictedClassification -or
        [string]$detailAttachments[0].path -notlike '*RESTRICTED*') {
        throw "Existing package raw guest attachment is missing its restricted personal-data label: $PackageId"
    }
    return $manifest
}

function New-GssEmailPackage {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][object]$RunLog,
        [Parameter(Mandatory)][object]$AnalysisResult,
        [string]$LedgerPath,
        [AllowNull()][object]$ExpectedPackageInputEvidence
    )

    $packageMutex = New-Object System.Threading.Mutex($false, $script:GssTransactionMutexName)
    $ownsPackageMutex = $false
    try {
        try {
            $ownsPackageMutex = $packageMutex.WaitOne(0)
        }
        catch [System.Threading.AbandonedMutexException] {
            $ownsPackageMutex = $true
        }
        if (-not $ownsPackageMutex) {
            throw 'Another GSS workbook transaction is already active on this workstation.'
        }

    if ($RunLog.Mode -ne 'ApplyToMainWorkbook') {
        throw 'Email packages may be published only from a successful live apply log.'
    }
    if ($AnalysisResult.WorkbookStatus -eq 'Blocked') {
        throw 'Email package publication is blocked because workbook validation failed.'
    }
    if ($AnalysisResult.AnalysisStatus -eq 'Blocked') {
        throw 'Email package publication is blocked because analysis validation failed.'
    }

    $reportingDate = [datetime]::Parse([string]$RunLog.CurrentWeekEnding).Date
    $inventory = Get-GssDetailInventory -FolderPath $FolderPath -ReportingDate $reportingDate
    $commenterLens = Get-GssCommenterLens `
        -Inventory $inventory `
        -MetricDetail @($AnalysisResult.MetricDetail) `
        -ReportingDate $reportingDate
    Assert-GssCommenterLensContract -CommenterLens $commenterLens
    $AnalysisResult | Add-Member -NotePropertyName CommenterLens -NotePropertyValue $commenterLens -Force
    if ($null -eq $AnalysisResult.PSObject.Properties['Modeling'] -or $null -eq $AnalysisResult.Modeling) {
        $AnalysisResult | Add-Member `
            -NotePropertyName Modeling `
            -NotePropertyValue (Get-GssPopulationModelingAvailability) `
            -Force
    }
    $sourceDescriptors = @(
        (Get-GssSourceDescriptor -Role 'comparison_pdf' -Path ([string]$RunLog.EmailComparisonPdf) -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'rolling_workbook' -Path ([string]$RunLog.CurrentSourceWorkbook) -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'prior_year_rolling_workbook' -Path ([string]$RunLog.PriorYearSourceWorkbook) -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'live_workbook' -Path ([string]$RunLog.TargetWorkbook) -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'detail_workbook' -Path $inventory.CurrentWorkbook.Path -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'run_log' -Path ([string]$AnalysisResult.LogPath) -FolderPath $FolderPath)
    )
    $sourceDescriptors += @($inventory.Workbooks |
        Where-Object { $_.PortablePath -ne $inventory.CurrentWorkbook.PortablePath } |
        Sort-Object PortablePath |
        ForEach-Object { Get-GssSourceDescriptor -Role 'detail_archive_workbook' -Path $_.Path -FolderPath $FolderPath })
    Assert-GssLoggedFileEvidence -RunLog $RunLog -Descriptors $sourceDescriptors
    if ([string]::IsNullOrWhiteSpace($LedgerPath)) {
        $LedgerPath = Join-Path (Join-Path (Join-Path $FolderPath '_automation_runs') 'state') 'gss_feedback_first_seen.json'
    }
    if ($null -ne $ExpectedPackageInputEvidence) {
        $canonicalLedgerPath = Join-Path (Join-Path (Join-Path $FolderPath '_automation_runs') 'state') 'gss_feedback_first_seen.json'
        if (-not [string]::Equals(
            [System.IO.Path]::GetFullPath($LedgerPath),
            [System.IO.Path]::GetFullPath($canonicalLedgerPath),
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            throw 'Expected package-input evidence requires the canonical feedback ledger path.'
        }
    }
    $ledgerRollbackState = Get-GssFeedbackLedgerRollbackState -Path $LedgerPath
    $null = Assert-GssExpectedPackageInputEvidence `
        -ExpectedEvidence $ExpectedPackageInputEvidence `
        -SourceDescriptors $sourceDescriptors `
        -LedgerPath $LedgerPath `
        -LedgerState $ledgerRollbackState
    $ledger = Read-GssFeedbackLedger $LedgerPath
    $feedbackSelection = Get-GssFeedbackSelection -Inventory $inventory -Ledger $ledger -ReportingDate $reportingDate
    $packageId = Get-GssDeterministicPackageId `
        -ReportingDate $reportingDate.ToString('yyyy-MM-dd') `
        -SourceDescriptors $sourceDescriptors `
        -FeedbackSelectionFingerprint $feedbackSelection.Fingerprint
    $feedback = New-GssSanitizedFeedback `
        -Inventory $inventory `
        -Ledger $ledger `
        -PackageId $packageId `
        -ReportingDate $reportingDate `
        -FeedbackSelection $feedbackSelection
    if (-not $feedback.Privacy.pii_scan_passed) {
        throw "PII remained after guest-feedback sanitization: $($feedback.Privacy.remaining_pii_types -join ', ')"
    }

    if (Get-Command Select-GssRestaurantFindings -ErrorAction SilentlyContinue) {
        $AnalysisResult.RestaurantFindings = @(Select-GssRestaurantFindings -MetricDetail $AnalysisResult.MetricDetail -GuestThemes $feedback.Themes)
    }

    $rollingStart = $reportingDate.AddDays(-90)
    $previousWindowEnd = $reportingDate.AddDays(-7)
    $previousWindowStart = $rollingStart.AddDays(-7)
    $priorYearEnd = [datetime]::Parse([string]$RunLog.PriorYearWeekEnding).Date
    $reporting = [pscustomobject][ordered]@{
        reporting_date = $reportingDate.ToString('yyyy-MM-dd')
        rolling_start_date = $rollingStart.ToString('yyyy-MM-dd')
        rolling_end_date = $reportingDate.ToString('yyyy-MM-dd')
        previous_window_start_date = $previousWindowStart.ToString('yyyy-MM-dd')
        previous_window_end_date = $previousWindowEnd.ToString('yyyy-MM-dd')
        prior_year_end_date = $priorYearEnd.ToString('yyyy-MM-dd')
        result_label = "13-week rolling through $($reportingDate.ToString('yyyy-MM-dd'))"
        comparison_label = 'change versus previous rolling window'
        adjacent_window_overlap_weeks = [int]$script:GssAnalysisPolicy.rolling_windows.adjacent_overlap_weeks
    }
    $statuses = [pscustomobject][ordered]@{
        workbook_status = [string]$AnalysisResult.WorkbookStatus
        analysis_status = [string]$AnalysisResult.AnalysisStatus
        email_readiness = 'Ready'
        population_modeling_status = [string]$AnalysisResult.Modeling.Status
        commenter_lens_status = [string]$commenterLens.status
    }

    $metricEvidence = @()
    foreach ($restaurant in @($AnalysisResult.RestaurantFindings)) {
        foreach ($item in @($restaurant.Strengths)) { $metricEvidence += ConvertTo-GssMetricEvidenceCard $item 'strength' }
        foreach ($item in @($restaurant.Opportunities)) { $metricEvidence += ConvertTo-GssMetricEvidenceCard $item 'opportunity' }
    }
    $themeEvidence = @()
    foreach ($theme in @($feedback.Themes)) {
        $restaurantFinding = $AnalysisResult.RestaurantFindings | Where-Object { $_.RestaurantId -eq $theme.restaurant_id } | Select-Object -First 1
        $sourceEntity = if ($restaurantFinding) { $restaurantFinding.Restaurant } else { [string]$theme.restaurant_id }
        $restaurantName = if ($restaurantFinding -and $restaurantFinding.Name) { $restaurantFinding.Name } else { Get-GssRestaurantDisplayName $sourceEntity }
        $findingType = if ([int]$theme.concern_count -gt [int]$theme.positive_count) { 'opportunity' } elseif ([int]$theme.positive_count -gt [int]$theme.concern_count) { 'strength' } else { 'mixed' }
        $dncNote = if ([int]$theme.do_not_contact_count -gt 0) { "; $($theme.do_not_contact_count) do-not-contact response(s) included only in aggregate counts" } else { '' }
        $displayText = "${restaurantName}: among guests who provided comments, new guest feedback $($theme.category) theme in $($theme.unique_response_count) of $($theme.denominator_response_count) unique comment-providing responses with visit dates $($theme.visit_date_start) through $($theme.visit_date_end); $($theme.concern_count) concern and $($theme.positive_count) positive$dncNote. Theme categories are non-exclusive."
        $themeEvidence += [pscustomobject][ordered]@{
            theme_id = $theme.theme_id
            restaurant_id = $theme.restaurant_id
            restaurant = $restaurantName
            restaurant_name = $restaurantName
            source_entity = $sourceEntity
            kind = 'guest_feedback_theme'
            finding_type = $findingType
            category = $theme.category
            unique_response_count = [int]$theme.unique_response_count
            denominator_response_count = [int]$theme.denominator_response_count
            visit_date_start = [string]$theme.visit_date_start
            visit_date_end = [string]$theme.visit_date_end
            categories_are_non_exclusive = [bool]$theme.categories_are_non_exclusive
            concern_count = [int]$theme.concern_count
            positive_count = [int]$theme.positive_count
            do_not_contact_count = [int]$theme.do_not_contact_count
            quote_allowed = $false
            outreach_allowed = $false
            display_text = $displayText
        }
    }
    $feedbackBasis = if ($feedback.ResponseCount -gt 0) {
        "$($feedback.ResponseCount) unique new guest-feedback responses with visit dates $($feedback.VisitDateStart) through $($feedback.VisitDateEnd)"
    }
    else {
        'no previously unseen guest-feedback responses in the validated exports'
    }
    $portfolioDisplayText = "Portfolio reporting basis: 13-week rolling population aggregates through $($reporting.reporting_date); adjacent 13-week windows overlap by 12 weeks; changes compare with the previous rolling window and matching prior-year rolling window; $feedbackBasis. Raw row-level scores describe only guests who provided comments and are not population-representative. This risk-reduced analysis requires human review. These comparisons are directional and do not establish statistical significance or causation."
    $portfolioEvidence = [pscustomobject][ordered]@{
        evidence_id = 'portfolio-' + (Get-GssStringSha256 ("$packageId|reporting-basis")).Substring(0, 16)
        kind = 'reporting_basis'
        finding_type = 'context'
        restaurant_id = 'portfolio'
        restaurant = 'Portfolio'
        restaurant_name = 'Portfolio'
        display_text = $portfolioDisplayText
    }
    $restaurantSections = @()
    foreach ($restaurant in @($AnalysisResult.RestaurantFindings)) {
        $restaurantThemes = @($feedback.Themes | Where-Object { $_.restaurant_id -eq $restaurant.RestaurantId })
        $restaurantFeedbackCards = @($feedback.InternalCards | Where-Object { $_.RestaurantId -eq $restaurant.RestaurantId })
        $restaurantVisitStart = if ($restaurantFeedbackCards.Count -gt 0) { ($restaurantFeedbackCards.VisitDate | Sort-Object | Select-Object -First 1) } else { $null }
        $restaurantVisitEnd = if ($restaurantFeedbackCards.Count -gt 0) { ($restaurantFeedbackCards.VisitDate | Sort-Object -Descending | Select-Object -First 1) } else { $null }
        $restaurantSections += [pscustomobject][ordered]@{
            restaurant_id = $restaurant.RestaurantId
            restaurant = $restaurant.Name
            name = $restaurant.Name
            source_entity = $restaurant.Restaurant
            strengths = @($metricEvidence | Where-Object { $_.restaurant_id -eq $restaurant.RestaurantId -and $_.kind -eq 'strength' })
            opportunities = @($metricEvidence | Where-Object { $_.restaurant_id -eq $restaurant.RestaurantId -and $_.kind -eq 'opportunity' })
            guest_feedback = [pscustomobject][ordered]@{
                response_count = $restaurantFeedbackCards.Count
                visit_date_start = $restaurantVisitStart
                visit_date_end = $restaurantVisitEnd
                themes = $restaurantThemes
                evidence = @($themeEvidence | Where-Object { $_.restaurant_id -eq $restaurant.RestaurantId })
            }
            commenter_lens = (
                $commenterLens.restaurants |
                    Where-Object { $_.restaurant_id -eq $restaurant.RestaurantId } |
                    Select-Object -First 1
            )
        }
    }

    $analysisDocument = [pscustomobject][ordered]@{
        schema_version = $script:GssEmailPackageSchemaVersion
        policy_version = $script:GssAnalysisPolicyVersion
        package_id = $packageId
        classification = $script:GssRestrictedClassification
        package_contains_personal_data = $true
        feedback_selection_sha256 = $feedback.SelectionFingerprint
        reporting = $reporting
        statuses = $statuses
        distribution_controls = [pscustomobject][ordered]@{
            automatic_sending_enabled = $false
            human_review_required = $true
            restricted_recipient_review = [pscustomobject][ordered]@{
                required_before_manual_send = $true
                status = 'pending_manual_confirmation'
                reviewed_recipient_count = 0
                confirmation_source = 'manual'
            }
            ready_marker_meaning = 'Integrity-validated and ready for manual content/recipient review; never PII-free or send-approved.'
        }
        methodology = [pscustomobject][ordered]@{
            score_basis = 'Each score is a 13-week rolling result; rolling rows are not averaged together.'
            comparison_basis = 'Direction-adjusted change versus the previous 13-week rolling window, which overlaps the current window by 12 weeks, and the matching prior-year rolling window.'
            evidence_dimensions = @('level', 'movement', 'benchmark')
            confidence_policy = [pscustomobject][ordered]@{
                high = '100 or more responses'
                developing = '50-99 responses; threshold and corroboration required for top findings'
                low = '1-49 responses; never eligible for top findings'
                not_scored = 'zero or missing responses'
            }
            significance_claimed = $false
            causation_claimed = $false
            human_review_required = $true
            analysis_description = 'risk-reduced'
            population_modeling_status = [string]$AnalysisResult.Modeling.Status
            commenter_lens_scope = [string]$commenterLens.scope_label
            commenter_rows_population_representative = $false
            commenter_alignment_basis = [string]$commenterLens.reporting_window.source_alignment
            commenter_exact_partition_alignment_verified = [bool]$commenterLens.reporting_window.exact_partition_alignment_verified
            commenter_rate_denominator = 'comment-providing responses with a nonmissing metric score'
            population_rate_denominator = 'vendor rolling population aggregate metric denominator'
            feedback_label = 'new guest feedback among guests who provided comments'
            feedback_response_count = $feedback.ResponseCount
            feedback_visit_date_start = $feedback.VisitDateStart
            feedback_visit_date_end = $feedback.VisitDateEnd
            theme_denominator_label = 'N of M unique comment-providing responses over the stated visit-date range'
            theme_categories_are_non_exclusive = $true
        }
        commenter_lens = $commenterLens
        portfolio = [pscustomobject][ordered]@{
            evidence = @($portfolioEvidence)
        }
        restaurants = $restaurantSections
        portfolio_evidence = @($portfolioEvidence)
        metric_evidence = $metricEvidence
        theme_evidence = $themeEvidence
        feedback_themes = $feedback.Themes
        sanitized_feedback = $feedback.Cards
        evidence_cards = @($portfolioEvidence) + @($metricEvidence) + @($themeEvidence) + @($feedback.Cards)
        privacy = $feedback.Privacy
    }
    $null = Test-GssAnalysisEvidenceContract -Analysis $analysisDocument

    $outbox = Join-Path (Join-Path $FolderPath '_automation_runs') 'email_outbox'
    New-Item -ItemType Directory -Path $outbox -Force | Out-Null
    $packagePath = Join-Path $outbox $packageId
    if (Test-Path -LiteralPath $packagePath -PathType Container) {
        $existingManifest = Test-GssExistingEmailPackage `
            -PackagePath $packagePath `
            -PackageId $packageId `
            -ExpectedSourceDescriptors $sourceDescriptors `
            -ExpectedFeedbackSelectionFingerprint $feedback.SelectionFingerprint
        $null = Assert-GssExpectedPackageInputEvidence `
            -ExpectedEvidence $ExpectedPackageInputEvidence `
            -SourceDescriptors $sourceDescriptors `
            -LedgerPath $LedgerPath
        Write-GssFeedbackLedgerWithRollbackOnFailure `
            -Ledger $feedback.NextLedger `
            -Path $LedgerPath `
            -RollbackState $ledgerRollbackState
        return [pscustomobject]@{
            EmailReadiness = 'Ready'
            DataClassification = $script:GssRestrictedClassification
            AutomaticSendingEnabled = $false
            RestrictedRecipientReviewStatus = 'pending_manual_confirmation'
            PackageId = $packageId
            PackagePath = $packagePath
            ManifestPath = Join-Path $packagePath 'email_manifest.json'
            CommenterLensJsonPath = Join-Path $packagePath 'commenter_lens.json'
            CommenterLensCsvPath = Join-Path $packagePath 'commenter_lens.csv'
            ReadyMarkerPath = Join-Path $packagePath 'READY'
            ExistingPackage = $true
        }
    }

    $stagingPath = Join-Path $outbox ('.staging-' + $packageId + '-' + [guid]::NewGuid().ToString('N'))
    $attachmentDirectory = Join-Path $stagingPath 'attachments'
    New-Item -ItemType Directory -Path $attachmentDirectory -Force | Out-Null
    $ledgerMutationAttempted = $false
    try {
        $dateToken = $reportingDate.ToString('MMddyy')
        $attachmentNames = @{
            comparison_pdf = "GSS Email Comparison $dateToken.pdf"
            rolling_workbook = "GSS 13-Week Rolling $dateToken.xlsx"
            detail_workbook = "RESTRICTED GSS Detail $dateToken.xlsx"
        }
        $attachments = @()
        foreach ($source in @($sourceDescriptors | Where-Object { $_.role -in @('comparison_pdf', 'rolling_workbook', 'detail_workbook') })) {
            $fileName = $attachmentNames[$source.role]
            $destination = Join-Path $attachmentDirectory $fileName
            Copy-Item -LiteralPath $source.source_full_path -Destination $destination
            $copiedItem = Get-Item -LiteralPath $destination
            $copiedHash = Get-GssSha256 $destination
            if ([long]$copiedItem.Length -ne [long]$source.byte_size -or $copiedHash -ne $source.sha256) {
                throw "Attachment changed while the package was being created: $($source.role)"
            }
            $attachments += [pscustomobject][ordered]@{
                role = $source.role
                path = 'attachments/' + $fileName
                source_path = $source.source_path
                byte_size = [long]$copiedItem.Length
                sha256 = $copiedHash
                contains_personal_data = ($source.role -eq 'detail_workbook')
                classification = if ($source.role -eq 'detail_workbook') { $script:GssRestrictedClassification } else { 'INTERNAL' }
            }
        }

        $previewText = New-GssEvidencePreviewText -Analysis $AnalysisResult -Feedback $feedback -ReportingDate $reportingDate
        $previewHtml = ConvertTo-GssSimpleHtml $previewText
        $analysisPath = Join-Path $stagingPath 'analysis.json'
        $commenterLensJsonPath = Join-Path $stagingPath 'commenter_lens.json'
        $commenterLensCsvPath = Join-Path $stagingPath 'commenter_lens.csv'
        $textPath = Join-Path $stagingPath 'email_preview.txt'
        $htmlPath = Join-Path $stagingPath 'email_preview.html'
        $classificationPath = Join-Path $stagingPath 'RESTRICTED.txt'
        $manifestPath = Join-Path $stagingPath 'email_manifest.json'
        Write-GssUtf8NoBomFile -Path $analysisPath -Value ($analysisDocument | ConvertTo-Json -Depth 12)
        Write-GssUtf8NoBomFile -Path $commenterLensJsonPath -Value ($commenterLens | ConvertTo-Json -Depth 12)
        Write-GssUtf8NoBomFile -Path $commenterLensCsvPath -Value (ConvertTo-GssCommenterLensCsv -CommenterLens $commenterLens)
        Write-GssUtf8NoBomFile -Path $textPath -Value $previewText
        Write-GssUtf8NoBomFile -Path $htmlPath -Value $previewHtml
        Write-GssUtf8NoBomFile -Path $classificationPath -Value @"
$script:GssRestrictedClassification
This package includes a raw guest detail workbook. The analysis text is risk-reduced, but the package is not PII-free.
Automatic sending is disabled. Confirm every recipient is authorized before any manual send.
"@

        $sourceLogDescriptor = @($sourceDescriptors | Where-Object role -eq 'run_log')
        if ($sourceLogDescriptor.Count -ne 1) { throw 'Package source inventory must contain exactly one run_log descriptor.' }
        $sourceLogPath = [string]$sourceLogDescriptor[0].source_path
        $portableArtifacts = @(
            $script:GssPortableArtifactPaths.GetEnumerator() |
                ForEach-Object {
                    Get-GssPortableArtifactDescriptor `
                        -PackagePath $stagingPath `
                        -Role ([string]$_.Key) `
                        -RelativePath ([string]$_.Value)
                }
        )
        $manifest = [pscustomobject][ordered]@{
            schema_version = $script:GssEmailPackageSchemaVersion
            policy_version = $script:GssAnalysisPolicyVersion
            package_id = $packageId
            classification = $script:GssRestrictedClassification
            package_contains_personal_data = $true
            feedback_selection_sha256 = $feedback.SelectionFingerprint
            reporting = $reporting
            statuses = $statuses
            distribution_controls = $analysisDocument.distribution_controls
            source_log_path = $sourceLogPath
            analysis_path = 'analysis.json'
            commenter_lens_json_path = 'commenter_lens.json'
            commenter_lens_csv_path = 'commenter_lens.csv'
            text_preview_path = 'email_preview.txt'
            html_preview_path = 'email_preview.html'
            classification_notice_path = 'RESTRICTED.txt'
            ready_marker_path = 'READY'
            sources = @($sourceDescriptors | ForEach-Object {
                [pscustomobject][ordered]@{
                    role = $_.role
                    path = $_.source_path
                    byte_size = $_.byte_size
                    sha256 = $_.sha256
                }
            })
            attachments = $attachments
            portable_artifacts = $portableArtifacts
            evidence_ids = @(
                @($portfolioEvidence) + @($metricEvidence) + @($feedback.Cards) |
                    ForEach-Object { [string]$_.evidence_id } |
                    Where-Object { $_ } |
                    Sort-Object -Unique
            )
            theme_ids = @($feedback.Themes.theme_id | Sort-Object -Unique)
            privacy = $feedback.Privacy
        }
        Write-GssUtf8NoBomFile -Path $manifestPath -Value ($manifest | ConvertTo-Json -Depth 12)

        foreach ($requiredName in @('email_manifest.json', 'analysis.json', 'commenter_lens.json', 'commenter_lens.csv', 'email_preview.txt', 'email_preview.html', 'RESTRICTED.txt')) {
            $requiredPath = Join-Path $stagingPath $requiredName
            if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf) -or (Get-Item -LiteralPath $requiredPath).Length -le 0) {
                throw "Package output is empty or missing: $requiredName"
            }
            Assert-GssUtf8NoBomFile -Path $requiredPath -Label $requiredName
        }
        Assert-GssPortableArtifactInventory -PackagePath $stagingPath -Artifacts $portableArtifacts
        foreach ($attachment in $attachments) {
            $attachmentPath = Join-Path $stagingPath $attachment.path.Replace('/', '\')
            if ((Get-GssSha256 $attachmentPath) -ne $attachment.sha256) { throw "Final attachment validation failed: $($attachment.role)" }
        }
        foreach ($source in $sourceDescriptors) {
            $sourceItem = Get-Item -LiteralPath $source.source_full_path
            if ([long]$sourceItem.Length -ne [long]$source.byte_size -or (Get-GssSha256 $source.source_full_path) -ne $source.sha256) {
                throw "Source changed while the package was being created: $($source.role) $($source.source_path)"
            }
        }
        $analysisText = Read-GssUtf8NoBomFile $analysisPath
        $commenterLensJsonText = Read-GssUtf8NoBomFile $commenterLensJsonPath
        $commenterLensCsvText = Read-GssUtf8NoBomFile $commenterLensCsvPath
        $manifestText = Read-GssUtf8NoBomFile $manifestPath
        $classificationText = Read-GssUtf8NoBomFile $classificationPath
        $aiFacingText = $analysisText + $commenterLensJsonText + $commenterLensCsvText + $previewText + $previewHtml + $classificationText
        $nonRawText = $aiFacingText + $manifestText
        # Inspect decoded JSON values rather than serialized JSON syntax. A
        # quoted phrase ending in a letter and colon is serialized with \" and
        # must not be mistaken for a drive root solely because of that escape.
        Assert-GssPortableContentHasNoMachineSpecificPath `
            -StructuredValues @($analysisDocument, $commenterLens, $manifest) `
            -TextValues @($commenterLensCsvText, $previewText, $previewHtml, $classificationText)
        # Deterministic methodology and metadata can legitimately contain an
        # ordinary word that is also present in a guest-name field (for example,
        # "sample"). Scan all portable text for generic PII, but scope the
        # known-name check to guest-derived text that was actually published.
        $sanitizedFeedbackText = @($feedback.Cards | ForEach-Object {
            [string]$_.sanitized_text
            [string]$_.display_text
        }) -join "`n"
        $remainingPii = @(
            @(Get-GssRemainingPiiTypes -Text $nonRawText -KnownNames @()) +
            @(Get-GssRemainingPiiTypes -Text $sanitizedFeedbackText -KnownNames (Get-GssKnownGuestNames $inventory.AllResponseInstances)) |
                Sort-Object -Unique
        )
        if ($remainingPii.Count -gt 0) { throw "PII remained in portable package output: $($remainingPii -join ', ')" }

        $null = Assert-GssExpectedPackageInputEvidence `
            -ExpectedEvidence $ExpectedPackageInputEvidence `
            -SourceDescriptors $sourceDescriptors `
            -LedgerPath $LedgerPath
        # Persist first-seen state before making the ready package visible. The
        # exact prior ledger bytes are restored if any later publication step fails.
        $ledgerMutationAttempted = $true
        Write-GssFeedbackLedger -Ledger $feedback.NextLedger -Path $LedgerPath
        # This marker is the final package-file write; the directory is then promoted atomically.
        "$($script:GssEmailPackageSchemaVersion)`n$packageId" | Set-Content -LiteralPath (Join-Path $stagingPath 'READY') -Encoding ASCII
        $null = Publish-GssStagedEmailPackage `
            -StagingPath $stagingPath `
            -PackagePath $packagePath `
            -ValidationOperation {
                Test-GssExistingEmailPackage `
                    -PackagePath $packagePath `
                    -PackageId $packageId `
                    -ExpectedSourceDescriptors $sourceDescriptors `
                    -ExpectedFeedbackSelectionFingerprint $feedback.SelectionFingerprint
            }
        return [pscustomobject]@{
            EmailReadiness = 'Ready'
            DataClassification = $script:GssRestrictedClassification
            AutomaticSendingEnabled = $false
            RestrictedRecipientReviewStatus = 'pending_manual_confirmation'
            PackageId = $packageId
            PackagePath = $packagePath
            ManifestPath = Join-Path $packagePath 'email_manifest.json'
            CommenterLensJsonPath = Join-Path $packagePath 'commenter_lens.json'
            CommenterLensCsvPath = Join-Path $packagePath 'commenter_lens.csv'
            ReadyMarkerPath = Join-Path $packagePath 'READY'
            ExistingPackage = $false
        }
    }
    catch {
        $packageError = $_
        if ($ledgerMutationAttempted) {
            try {
                Restore-GssFeedbackLedgerState -Path $LedgerPath -State $ledgerRollbackState
            }
            catch {
                $packageError.Exception.Data['GssFeedbackLedgerRollbackError'] = $_.Exception.Message
            }
        }
        if (Test-Path -LiteralPath $stagingPath -PathType Container) {
            $resolvedOutbox = [System.IO.Path]::GetFullPath($outbox).TrimEnd('\') + '\'
            $resolvedStaging = [System.IO.Path]::GetFullPath($stagingPath)
            if ($resolvedStaging.StartsWith($resolvedOutbox, [System.StringComparison]::OrdinalIgnoreCase)) {
                try {
                    Invoke-GssFileSystemRetry -Operation {
                        Remove-Item -LiteralPath $resolvedStaging -Recurse -Force
                    }
                }
                catch {
                    $packageError.Exception.Data['GssStagingCleanupError'] = $_.Exception.Message
                }
            }
        }
        throw
    }
    }
    finally {
        try {
            if ($ownsPackageMutex) {
                [void]$packageMutex.ReleaseMutex()
            }
        }
        finally {
            $packageMutex.Dispose()
        }
    }
}
