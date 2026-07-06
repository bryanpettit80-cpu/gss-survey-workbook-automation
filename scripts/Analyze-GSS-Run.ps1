[CmdletBinding()]
param(
    [string]$Folder,
    [string]$LogPath,
    [int]$LookbackWeeks = 8,
    [string]$MainWorkbookName = 'Consolidated_Score_Trends_v6_ExecClean_YoY_WITH_QuickRead_WoW_YoY_FINAL_v14_UniformCF_DriversStyle_PATCHED_XMLSAFE.xlsx',
    [switch]$OutputObject
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1')

function Resolve-GssAnalysisFolder {
    param([string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        $projectRoot = Split-Path -Parent $scriptRoot
        $FolderPath = Split-Path -Parent $projectRoot
    }

    $FolderPath = $FolderPath.Trim('"')
    return (Resolve-Path -LiteralPath $FolderPath).Path
}

function Get-LatestGssRunLog {
    param([string]$FolderPath)

    $logDir = Join-Path (Join-Path $FolderPath '_automation_runs') 'logs'
    if (-not (Test-Path -LiteralPath $logDir -PathType Container)) {
        throw "GSS log folder not found: $logDir"
    }

    $latest = Get-ChildItem -LiteralPath $logDir -File -Filter 'gss_update_*.json' |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $latest) {
        throw "No gss_update_*.json files were found in $logDir"
    }

    return $latest.FullName
}

function Read-GssRunLog {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "GSS run log not found: $Path"
    }

    return Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json
}

function Get-DirectionAdjustedChange {
    param([object]$CurrentValue, [object]$PriorValue, [bool]$LowerIsBetter)

    if ($null -eq $CurrentValue -or $null -eq $PriorValue) { return $null }
    if ($LowerIsBetter) {
        return [double]$PriorValue - [double]$CurrentValue
    }
    return [double]$CurrentValue - [double]$PriorValue
}

function Get-Mean {
    param([object[]]$Values)

    $usable = @($Values | Where-Object { $null -ne $_ })
    if ($usable.Count -eq 0) { return $null }
    return ($usable | Measure-Object -Average).Average
}

function Convert-ToGssNumber {
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) { return $null }
    return [double]$Value
}

function Get-GssEntityKey {
    param([string]$Ownership, [string]$Restaurant)

    $restaurantName = $Restaurant
    if ([string]::IsNullOrWhiteSpace($restaurantName)) {
        $restaurantName = '(TOTAL)'
    }
    else {
        $restaurantName = $restaurantName.Trim()
    }

    return "$($Ownership.Trim())|$restaurantName"
}

function Get-GssEntityWeekKey {
    param([string]$EntityKey, [datetime]$WeekEnding)
    return "$EntityKey|$($WeekEnding.ToString('yyyyMMdd'))"
}

function Test-GssYoyPairing {
    param([datetime]$CurrentWeek, [datetime]$PriorYearWeek)
    return (($CurrentWeek.Date - $PriorYearWeek.Date).Days -eq 364)
}

function Get-RequiredGssEntities {
    return @(
        [pscustomobject]@{ EntityKey = 'All Franchisees|(TOTAL)'; Label = 'All Franchisees Total'; Ownership = 'All Franchisees'; Restaurant = $null; IncludeInInsights = $false },
        [pscustomobject]@{ EntityKey = 'Sorensen|(TOTAL)'; Label = 'Sorensen Total'; Ownership = 'Sorensen'; Restaurant = $null; IncludeInInsights = $false },
        [pscustomobject]@{ EntityKey = 'Sorensen|9354 Richmond'; Label = '9354 Richmond'; Ownership = 'Sorensen'; Restaurant = '9354 Richmond'; IncludeInInsights = $true },
        [pscustomobject]@{ EntityKey = 'Sorensen|9355 Virginia Beach'; Label = '9355 Virginia Beach'; Ownership = 'Sorensen'; Restaurant = '9355 Virginia Beach'; IncludeInInsights = $true }
    )
}

function ConvertTo-CellValue {
    param([object]$Value)

    if ($Value -is [System.Reflection.Missing]) { return $null }
    if ($Value -is [System.DBNull]) { return $null }
    return $Value
}

function ConvertTo-ArrayTable {
    param([object[,]]$Values, [int]$RowCount, [int]$ColumnCount)

    $headers = @()
    for ($col = 1; $col -le $ColumnCount; $col++) {
        $headers += [string](ConvertTo-CellValue $Values[1, $col])
    }

    $rows = @()
    for ($row = 2; $row -le $RowCount; $row++) {
        $record = [ordered]@{}
        for ($col = 1; $col -le $ColumnCount; $col++) {
            $record[$headers[$col - 1]] = ConvertTo-CellValue $Values[$row, $col]
        }
        $rows += [pscustomobject]$record
    }

    return [pscustomobject]@{
        Headers = $headers
        Rows = $rows
    }
}

function Read-GssWorksheetTable {
    param([object]$Worksheet)

    $lastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, 1).End($xlUp).Row
    $lastCol = $Worksheet.Cells.Item(1, $Worksheet.Columns.Count).End($xlToLeft).Column
    $range = $null
    try {
        $range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($lastRow, $lastCol))
        $values = $range.Value2
        return ConvertTo-ArrayTable $values $lastRow $lastCol
    }
    finally {
        Release-ComObject $range
    }
}

function Read-GssWorkbookData {
    param([string]$WorkbookPath)

    $excel = $null
    $workbook = $null
    $rawWs = $null
    $metricWs = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false

        $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
        $rawWs = $workbook.Worksheets.Item('Raw_Data')
        $metricWs = $workbook.Worksheets.Item('Metric_Catalog')

        $rawTable = Read-GssWorksheetTable $rawWs
        $metricTable = Read-GssWorksheetTable $metricWs

        $rawRows = @()
        foreach ($row in $rawTable.Rows) {
            if ($null -eq $row.Week) { continue }
            $entityKey = Get-GssEntityKey $row.Ownership $row.Restaurant
            $week = Convert-ExcelDate $row.Week
            $rawRows += [pscustomobject]@{
                SourceFile = $row.Source_File
                Week = $week
                Restaurant = $row.Restaurant
                Ownership = $row.Ownership
                EntityKey = $entityKey
                RowKey = if ($row.RowKey) { $row.RowKey } else { Get-GssEntityWeekKey $entityKey $week }
                Values = $row
            }
        }

        $metrics = @()
        foreach ($metric in $metricTable.Rows) {
            $displayName = $metric.'Metric (Display)'
            $rawKey = $metric.'MetricKey (Raw Header)'
            if ([string]::IsNullOrWhiteSpace($displayName) -or [string]::IsNullOrWhiteSpace($rawKey)) {
                continue
            }

            $metrics += [pscustomobject]@{
                DisplayName = $displayName
                RawKey = $rawKey
                LowerIsBetter = Convert-ToBool $metric.Lower_Is_Better
                IncludeInMovers = Convert-ToBool $metric.Include_In_Movers
                Category = $metric.Category
                Owner = $metric.Owner
            }
        }

        return [pscustomobject]@{
            Headers = $rawTable.Headers
            RawRows = $rawRows
            Metrics = $metrics
        }
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        Release-ComObject $metricWs
        Release-ComObject $rawWs
        Release-ComObject $workbook
        if ($excel) {
            $excel.Quit()
            Release-ComObject $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-GssRow {
    param([hashtable]$RowsByKey, [string]$EntityKey, [datetime]$Week)

    $key = Get-GssEntityWeekKey $EntityKey $Week
    if ($RowsByKey.ContainsKey($key)) {
        return $RowsByKey[$key]
    }

    return $null
}

function Get-GssMetricValue {
    param([object]$Row, [string]$MetricKey)

    if ($null -eq $Row) { return $null }
    if (-not $Row.Values.PSObject.Properties.Name.Contains($MetricKey)) { return $null }
    return Convert-ToGssNumber $Row.Values.$MetricKey
}

function New-GssMetricDetail {
    param(
        [object[]]$RawRows,
        [object[]]$Metrics,
        [datetime]$CurrentWeek,
        [datetime]$PriorYearWeek,
        [int]$LookbackWeeks
    )

    $rowsByKey = @{}
    foreach ($row in $RawRows) {
        $rowsByKey[(Get-GssEntityWeekKey $row.EntityKey $row.Week)] = $row
    }

    $entities = @(Get-RequiredGssEntities | Where-Object { $_.IncludeInInsights })
    $sorensenKey = 'Sorensen|(TOTAL)'
    $allFranchiseesKey = 'All Franchisees|(TOTAL)'
    $details = @()

    foreach ($entity in $entities) {
        foreach ($metric in $Metrics) {
            if (-not $metric.IncludeInMovers) { continue }

            $currentRow = Get-GssRow $rowsByKey $entity.EntityKey $CurrentWeek
            $priorWeekRow = Get-GssRow $rowsByKey $entity.EntityKey $CurrentWeek.AddDays(-7)
            $priorYearRow = Get-GssRow $rowsByKey $entity.EntityKey $PriorYearWeek
            $sorensenRow = Get-GssRow $rowsByKey $sorensenKey $CurrentWeek
            $allRow = Get-GssRow $rowsByKey $allFranchiseesKey $CurrentWeek
            $currentValue = Get-GssMetricValue $currentRow $metric.RawKey
            $priorWeekValue = Get-GssMetricValue $priorWeekRow $metric.RawKey
            $priorYearValue = Get-GssMetricValue $priorYearRow $metric.RawKey
            $sorensenValue = Get-GssMetricValue $sorensenRow $metric.RawKey
            $allValue = Get-GssMetricValue $allRow $metric.RawKey

            $lookbackValues = @()
            for ($i = 0; $i -lt $LookbackWeeks; $i++) {
                $lookbackRow = Get-GssRow $rowsByKey $entity.EntityKey $CurrentWeek.AddDays(-7 * $i)
                $lookbackValues += Get-GssMetricValue $lookbackRow $metric.RawKey
            }
            $rollingAverage = Get-Mean $lookbackValues
            $currentCount = Get-GssMetricValue $currentRow 'Count'

            $wow = Get-DirectionAdjustedChange $currentValue $priorWeekValue $metric.LowerIsBetter
            $yoy = Get-DirectionAdjustedChange $currentValue $priorYearValue $metric.LowerIsBetter
            $vsRolling = Get-DirectionAdjustedChange $currentValue $rollingAverage $metric.LowerIsBetter
            $vsSorensen = Get-DirectionAdjustedChange $currentValue $sorensenValue $metric.LowerIsBetter
            $vsAll = Get-DirectionAdjustedChange $currentValue $allValue $metric.LowerIsBetter
            $movementValues = @($wow, $yoy) | Where-Object { $null -ne $_ }
            $worstMovement = if ($movementValues.Count -gt 0) { ($movementValues | Measure-Object -Minimum).Minimum } else { $null }
            $bestMovement = if ($movementValues.Count -gt 0) { ($movementValues | Measure-Object -Maximum).Maximum } else { $null }

            $details += [pscustomobject]@{
                Entity = $entity.Label
                EntityKey = $entity.EntityKey
                WeekEnding = $CurrentWeek.ToString('yyyy-MM-dd')
                Metric = $metric.DisplayName
                RawMetric = $metric.RawKey
                Category = $metric.Category
                Owner = $metric.Owner
                LowerIsBetter = $metric.LowerIsBetter
                Current = $currentValue
                PriorWeek = $priorWeekValue
                PriorYear = $priorYearValue
                RollingAverage = $rollingAverage
                CurrentCount = $currentCount
                WoWImprovement = $wow
                YoYImprovement = $yoy
                VsRollingAverage = $vsRolling
                VsSorensenTotal = $vsSorensen
                VsAllFranchisees = $vsAll
                WorstMovement = $worstMovement
                BestMovement = $bestMovement
            }
        }
    }

    return $details
}

function Get-GssRunQa {
    param(
        [object]$RunLog,
        [string]$FolderPath,
        [string]$WorkbookPath,
        [object[]]$RawRows,
        [object[]]$Metrics,
        [string[]]$Headers,
        [object[]]$MetricDetail,
        [datetime]$CurrentWeek,
        [datetime]$PriorYearWeek
    )

    $blockers = @()
    $warnings = @()
    $rowsByKey = @{}
    foreach ($row in $RawRows) {
        $key = Get-GssEntityWeekKey $row.EntityKey $row.Week
        $rowsByKey[$key] = $row
    }

    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
        $blockers += "Main workbook is missing: $WorkbookPath"
    }
    if ($RunLog.EmailComparisonPdf -and -not (Test-Path -LiteralPath $RunLog.EmailComparisonPdf -PathType Leaf)) {
        $blockers += "Email comparison PDF is missing: $($RunLog.EmailComparisonPdf)"
    }
    if ($RunLog.Mode -eq 'ApplyToMainWorkbook' -and (-not $RunLog.BackupWorkbook -or -not (Test-Path -LiteralPath $RunLog.BackupWorkbook -PathType Leaf))) {
        $blockers += "Live apply backup workbook is missing."
    }
    if (-not (Test-GssYoyPairing $CurrentWeek $PriorYearWeek)) {
        $blockers += "Current week and prior-year week are not 364 days apart."
    }

    $expectedFolder = Join-Path $FolderPath '02 Weekly Rolling Source Workbooks'
    foreach ($sourcePath in @($RunLog.CurrentSourceWorkbook, $RunLog.PriorYearSourceWorkbook)) {
        if ($sourcePath -and -not ([System.IO.Path]::GetFullPath($sourcePath).StartsWith([System.IO.Path]::GetFullPath($expectedFolder), [System.StringComparison]::OrdinalIgnoreCase))) {
            $warnings += "Source workbook is outside 02 Weekly Rolling Source Workbooks: $sourcePath"
        }
    }

    $requiredEntities = @(Get-RequiredGssEntities)
    foreach ($week in @($CurrentWeek, $PriorYearWeek)) {
        foreach ($entity in $requiredEntities) {
            $key = Get-GssEntityWeekKey $entity.EntityKey $week
            if (-not $rowsByKey.ContainsKey($key)) {
                $blockers += "Missing Raw_Data row for $($entity.Label) on $($week.ToString('yyyy-MM-dd'))."
            }
        }
    }

    $duplicateRowKeys = @($RawRows | Group-Object RowKey | Where-Object { $_.Name -and $_.Count -gt 1 })
    foreach ($duplicate in $duplicateRowKeys) {
        $blockers += "Duplicate Raw_Data RowKey: $($duplicate.Name)"
    }

    foreach ($metric in $Metrics) {
        if (-not $Headers.Contains($metric.RawKey)) {
            $blockers += "Raw_Data is missing Metric_Catalog column: $($metric.RawKey)"
        }
    }

    if ($RunLog.RowsAppended -notin @(0, 4, 8)) {
        $warnings += "Unexpected RowsAppended value: $($RunLog.RowsAppended). Expected 0, 4, or 8."
    }

    foreach ($detail in $MetricDetail) {
        if ($detail.Metric -eq 'Count') { continue }
        if ($detail.CurrentCount -ne $null -and $detail.CurrentCount -lt 100) {
            $warnings += "Low sample count for $($detail.Entity): $($detail.CurrentCount)."
        }
        if ($detail.WorstMovement -ne $null -and $detail.WorstMovement -le -5) {
            $warnings += ("Large decline: {0} {1} moved {2:N1} points." -f $detail.Entity, $detail.Metric, $detail.WorstMovement)
        }
    }

    $status = if ($blockers.Count -gt 0) {
        'Blocked'
    }
    elseif ($warnings.Count -gt 0) {
        'Review'
    }
    else {
        'Ready'
    }

    return [pscustomobject]@{
        Status = $status
        Blockers = $blockers
        Warnings = @($warnings | Select-Object -Unique)
    }
}

function Select-GssAttentionItems {
    param([object[]]$MetricDetail, [int]$Count = 5)

    return @($MetricDetail |
        Where-Object { $_.WorstMovement -ne $null } |
        Sort-Object @{ Expression = { $_.WorstMovement }; Ascending = $true }, @{ Expression = { $_.CurrentCount }; Descending = $true } |
        Select-Object -First $Count)
}

function Select-GssStrengthItems {
    param([object[]]$MetricDetail, [int]$Count = 5)

    return @($MetricDetail |
        Where-Object { $_.BestMovement -ne $null -and $_.BestMovement -gt 0 -and ($_.WorstMovement -eq $null -or $_.WorstMovement -ge 0) } |
        Sort-Object @{ Expression = { $_.BestMovement }; Descending = $true }, @{ Expression = { $_.CurrentCount }; Descending = $true } |
        Select-Object -First $Count)
}

function Format-GssNumber {
    param([object]$Value, [int]$Digits = 1)

    if ($null -eq $Value) { return 'n/a' }
    return ('{0:N' + $Digits + '}') -f ([double]$Value)
}

function New-GssReviewMarkdown {
    param(
        [object]$Result,
        [object[]]$AttentionItems,
        [object[]]$StrengthItems
    )

    $lines = @()
    $lines += '# GSS Run Review'
    $lines += ''
    $lines += "Status: $($Result.OverallStatus)"
    $lines += "Run mode: $($Result.Run.Mode)"
    $lines += "Current week: $($Result.Run.CurrentWeekEnding)"
    $lines += "Prior-year week: $($Result.Run.PriorYearWeekEnding)"
    $lines += "Rows appended: $($Result.Run.RowsAppended)"
    $lines += ''
    $lines += '## Output Check'
    $lines += "- Workbook: $($Result.Run.TargetWorkbook)"
    $lines += "- Backup: $($Result.Run.BackupWorkbook)"
    $lines += "- Email PDF: $($Result.Run.EmailComparisonPdf)"
    $lines += "- Source: $($Result.Run.CurrentSourceWorkbook)"
    $lines += "- Prior-year source: $($Result.Run.PriorYearSourceWorkbook)"
    $lines += ''
    if ($Result.Qa.Blockers.Count -gt 0) {
        $lines += '## Blockers'
        foreach ($item in $Result.Qa.Blockers) { $lines += "- $item" }
        $lines += ''
    }
    if ($Result.Qa.Warnings.Count -gt 0) {
        $lines += '## Review Items'
        foreach ($item in $Result.Qa.Warnings) { $lines += "- $item" }
        $lines += ''
    }
    $lines += '## Needs Attention'
    if ($AttentionItems.Count -eq 0) {
        $lines += '- No material metric declines were detected.'
    }
    else {
        foreach ($item in $AttentionItems) {
            $lines += ('- {0}: {1} current {2}; worst movement {3}; WoW {4}; YoY {5}; vs all franchisees {6}.' -f `
                $item.Entity, $item.Metric, (Format-GssNumber $item.Current), (Format-GssNumber $item.WorstMovement), (Format-GssNumber $item.WoWImprovement), (Format-GssNumber $item.YoYImprovement), (Format-GssNumber $item.VsAllFranchisees))
        }
    }
    $lines += ''
    $lines += '## Strengths'
    if ($StrengthItems.Count -eq 0) {
        $lines += '- No positive movement items ranked above neutral.'
    }
    else {
        foreach ($item in $StrengthItems) {
            $lines += ('- {0}: {1} current {2}; best movement {3}; vs Sorensen total {4}.' -f `
                $item.Entity, $item.Metric, (Format-GssNumber $item.Current), (Format-GssNumber $item.BestMovement), (Format-GssNumber $item.VsSorensenTotal))
        }
    }
    $lines += ''
    $lines += '## Files'
    $lines += "- Detail CSV: $($Result.DetailCsvPath)"
    $lines += "- JSON: $($Result.JsonPath)"
    $lines += "- Review folder: $($Result.ReviewFolder)"
    $lines += ''

    return ($lines -join [Environment]::NewLine)
}

function Write-GssAnalysisSummary {
    param([object]$Result)

    Write-Host ''
    Write-Host 'GSS analytics review summary'
    Write-Host ('  Status: {0}' -f $Result.OverallStatus)
    Write-Host ('  Current week: {0}' -f $Result.Run.CurrentWeekEnding)
    Write-Host ('  Rows appended: {0}' -f $Result.Run.RowsAppended)
    Write-Host ('  Review: {0}' -f $Result.MarkdownPath)
    if ($Result.Qa.Blockers.Count -gt 0) {
        Write-Host ('  Blockers: {0}' -f $Result.Qa.Blockers.Count)
    }
    if ($Result.Qa.Warnings.Count -gt 0) {
        Write-Host ('  Review items: {0}' -f $Result.Qa.Warnings.Count)
    }
    Write-Host ''
}

function Invoke-GssRunAnalysis {
    param(
        [string]$FolderPath,
        [string]$RunLogPath,
        [int]$Weeks,
        [string]$WorkbookName,
        [switch]$ReturnObject
    )

    $FolderPath = Resolve-GssAnalysisFolder $FolderPath
    if ([string]::IsNullOrWhiteSpace($RunLogPath)) {
        $RunLogPath = Get-LatestGssRunLog $FolderPath
    }
    else {
        $RunLogPath = (Resolve-Path -LiteralPath $RunLogPath.Trim('"')).Path
    }

    $runLog = Read-GssRunLog $RunLogPath
    $workbookPath = Resolve-MainWorkbookPath $FolderPath $WorkbookName
    $workbookData = Read-GssWorkbookData $workbookPath
    $currentWeek = [datetime]::Parse($runLog.CurrentWeekEnding).Date
    $priorYearWeek = [datetime]::Parse($runLog.PriorYearWeekEnding).Date
    $metricDetail = @(New-GssMetricDetail $workbookData.RawRows $workbookData.Metrics $currentWeek $priorYearWeek $Weeks)
    $qa = Get-GssRunQa $runLog $FolderPath $workbookPath $workbookData.RawRows $workbookData.Metrics $workbookData.Headers $metricDetail $currentWeek $priorYearWeek
    $attentionItems = Select-GssAttentionItems $metricDetail
    $strengthItems = Select-GssStrengthItems $metricDetail

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reviewFolder = Join-Path (Join-Path (Join-Path $FolderPath '_automation_runs') 'qa') "run_review_$timestamp"
    New-Item -ItemType Directory -Path $reviewFolder -Force | Out-Null
    $markdownPath = Join-Path $reviewFolder 'review.md'
    $jsonPath = Join-Path $reviewFolder 'review.json'
    $detailCsvPath = Join-Path $reviewFolder 'metric_detail.csv'

    $result = [pscustomobject]@{
        OverallStatus = $qa.Status
        GeneratedAt = (Get-Date).ToString('s')
        Folder = $FolderPath
        ReviewFolder = $reviewFolder
        MarkdownPath = $markdownPath
        JsonPath = $jsonPath
        DetailCsvPath = $detailCsvPath
        LogPath = $RunLogPath
        Run = $runLog
        Qa = $qa
        TopAttention = $attentionItems
        TopStrengths = $strengthItems
        RawDataSummary = [pscustomobject]@{
            RowCount = $workbookData.RawRows.Count
            WeekCount = @($workbookData.RawRows | Select-Object -ExpandProperty Week -Unique).Count
            MetricCount = $workbookData.Metrics.Count
        }
    }

    $markdown = New-GssReviewMarkdown $result $attentionItems $strengthItems
    $markdown | Set-Content -LiteralPath $markdownPath -Encoding UTF8
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    $metricDetail | Export-Csv -LiteralPath $detailCsvPath -NoTypeInformation -Encoding UTF8

    Write-GssAnalysisSummary $result
    if ($ReturnObject) {
        return $result
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-GssRunAnalysis -FolderPath $Folder -RunLogPath $LogPath -Weeks $LookbackWeeks -WorkbookName $MainWorkbookName -ReturnObject:$OutputObject
}
