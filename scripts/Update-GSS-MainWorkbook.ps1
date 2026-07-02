[CmdletBinding()]
param(
    [string]$Folder,
    [string]$MainWorkbookName = 'Consolidated_Score_Trends_v6_ExecClean_YoY_WITH_QuickRead_WoW_YoY_FINAL_v14_UniformCF_DriversStyle_PATCHED_XMLSAFE.xlsx',
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

$xlUp = -4162
$xlToLeft = -4159
$xlPasteFormats = -4122
$xlCalculationAutomatic = -4105
$xlTypePDF = 0

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Normalize-Header {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return ([regex]::Replace(([string]$Value).ToLowerInvariant(), '[^a-z0-9]+', '')).Trim()
}

function Get-CellString {
    param([object]$Worksheet, [int]$Row, [int]$Column)
    $value = $Worksheet.Cells.Item($Row, $Column).Value2
    if ($null -eq $value) { return '' }
    return ([string]$value).Trim()
}

function Set-CellValue {
    param([object]$Worksheet, [int]$Row, [int]$Column, [object]$Value)
    $cell = $null
    try {
        $cell = $Worksheet.Cells.Item($Row, $Column)
        if ($null -eq $Value) {
            $cell.Value2 = ''
            return
        }
        if ($Value -is [bool]) {
            $Value = if ($Value) { 'TRUE' } else { 'FALSE' }
        }

        try {
            $cell.Value2 = $Value
        }
        catch [System.InvalidCastException] {
            $cell.Value = $Value
        }
    }
    finally {
        Release-ComObject $cell
    }
}

function Convert-ExcelDate {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.Date }
    if ($Value -is [double] -or $Value -is [int]) { return ([datetime]::FromOADate([double]$Value)).Date }
    return ([datetime]::Parse([string]$Value)).Date
}

function Convert-ToNullableDouble {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) { return $null }
    return [double]$Value
}

function Convert-ToBool {
    param([object]$Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    if ($Value -is [double] -or $Value -is [int]) { return ([double]$Value) -ne 0 }
    return ([string]$Value).Trim() -match '^(true|yes|1)$'
}

function Test-ExcludedGssPath {
    param([string]$Path, [string]$FolderPath)

    $resolvedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    $excludedFolders = @(
        (Join-Path $FolderPath 'GSS Survey Workbook Automation'),
        (Join-Path $FolderPath '_automation_runs')
    )

    foreach ($excludedFolder in $excludedFolders) {
        $resolvedExcluded = [System.IO.Path]::GetFullPath($excludedFolder).TrimEnd('\')
        if ($resolvedPath.Equals($resolvedExcluded, [System.StringComparison]::OrdinalIgnoreCase) -or
            $resolvedPath.StartsWith("$resolvedExcluded\", [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Get-GssFiles {
    param([string]$FolderPath, [string]$Filter)

    Get-ChildItem -LiteralPath $FolderPath -File -Filter $Filter -Recurse |
        Where-Object { -not (Test-ExcludedGssPath $_.FullName $FolderPath) }
}

function Resolve-MainWorkbookPath {
    param([string]$FolderPath, [string]$WorkbookName)

    $directPath = Join-Path $FolderPath $WorkbookName
    if (Test-Path -LiteralPath $directPath) {
        return (Resolve-Path -LiteralPath $directPath).Path
    }

    $matches = @(Get-GssFiles $FolderPath $WorkbookName)
    if ($matches.Count -eq 1) {
        return $matches[0].FullName
    }
    if ($matches.Count -gt 1) {
        $paths = ($matches | ForEach-Object { $_.FullName }) -join '; '
        throw "Multiple main workbook matches were found. Move old copies out of the GSS Surveys folder or pass -MainWorkbookName. Matches: $paths"
    }

    throw "Main workbook not found under ${FolderPath}: $WorkbookName"
}

function Get-HeaderMap {
    param([object]$Worksheet, [int]$HeaderRow)

    $lastCol = $Worksheet.Cells.Item($HeaderRow, $Worksheet.Columns.Count).End($xlToLeft).Column
    $map = @{}
    for ($col = 1; $col -le $lastCol; $col++) {
        $key = Normalize-Header (Get-CellString $Worksheet $HeaderRow $col)
        if ($key) { $map[$key] = $col }
    }
    return $map
}

function Get-NormalizedSourceName {
    param([System.IO.FileInfo]$File)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    $baseName = [regex]::Replace($baseName, '\s+FY\d{2,4}$', '')
    return "$baseName.xlsx"
}

function Get-RollingSources {
    param([object]$Excel, [string]$FolderPath)

    $candidates = Get-GssFiles $FolderPath 'Sorensen FW*.xlsx' |
        Where-Object { $_.Name -notmatch '(?i)CLEANED|Tidy|YY Trends' } |
        Sort-Object LastWriteTime -Descending

    $sources = @()
    foreach ($candidate in $candidates) {
        $sourceWb = $null
        $sourceWs = $null
        try {
            $sourceWb = $Excel.Workbooks.Open($candidate.FullName, 0, $true)
            $sourceWs = $sourceWb.Worksheets.Item(1)
            $a1 = Get-CellString $sourceWs 1 1

            $headerRow = $null
            for ($row = 1; $row -le 10; $row++) {
                $restaurantHeader = Normalize-Header (Get-CellString $sourceWs $row 1)
                $ownershipHeader = Normalize-Header (Get-CellString $sourceWs $row 2)
                if ($restaurantHeader -eq 'restaurant' -and $ownershipHeader -eq 'ownership') {
                    $headerRow = $row
                    break
                }
            }

            if ($null -eq $headerRow) {
                continue
            }

            if ($a1 -match 'Date Range:\s*(?<start>\d{1,2}/\d{1,2}/\d{4})\s+to\s+(?<end>\d{1,2}/\d{1,2}/\d{4})') {
                $endDate = [datetime]::Parse(
                    $Matches.end,
                    [Globalization.CultureInfo]::InvariantCulture
                ).Date

                $sources += [pscustomobject]@{
                    File = $candidate
                    WeekEnding = $endDate
                    SourceFileName = Get-NormalizedSourceName $candidate
                    HeaderRow = $headerRow
                }
            }
        }
        finally {
            if ($sourceWb) { $sourceWb.Close($false) }
            Release-ComObject $sourceWs
            Release-ComObject $sourceWb
        }
    }

    if ($sources.Count -eq 0) {
        throw "No rolling Sorensen source workbook with a Date Range header was found in $FolderPath."
    }

    return $sources
}

function Select-RollingSourceByWeek {
    param(
        [object[]]$Sources,
        [datetime]$WeekEnding
    )

    return $Sources |
        Where-Object { $_.WeekEnding.Date -eq $WeekEnding.Date } |
        Sort-Object @{ Expression = { $_.File.LastWriteTime }; Descending = $true } |
        Select-Object -First 1
}

function Select-LatestRollingSource {
    param([object[]]$Sources)

    return $Sources |
        Sort-Object @{ Expression = { $_.WeekEnding }; Descending = $true }, @{ Expression = { $_.File.LastWriteTime }; Descending = $true } |
        Select-Object -First 1
}

function Get-SourceRows {
    param(
        [object]$SourceWorksheet,
        [hashtable]$HeaderMap,
        [int]$HeaderRow,
        [datetime]$WeekEnding,
        [string]$SourceFileName
    )

    $requiredSourceHeaders = @(
        'Restaurant',
        'Ownership',
        'Count',
        'Overall Experience',
        'Service',
        'Culinary',
        'Value',
        'Pace of Meal',
        'Steak Cooked Properly (Y)',
        'Mgr Asked About Exp (Y)',
        'Overall Dissat (lower is better)',
        'Service Dissat (lower is better)',
        'Culinary Dissat (lower is better)',
        'Value Dissat (lower is better)',
        'Pace Of Meal Dissat (lower is better)'
    )

    foreach ($header in $requiredSourceHeaders) {
        $key = Normalize-Header $header
        if (-not $HeaderMap.ContainsKey($key)) {
            throw "Source workbook is missing required header '$header'."
        }
    }

    $restaurantCol = $HeaderMap[(Normalize-Header 'Restaurant')]
    $ownershipCol = $HeaderMap[(Normalize-Header 'Ownership')]
    $lastRow = $SourceWorksheet.Cells.Item($SourceWorksheet.Rows.Count, $ownershipCol).End($xlUp).Row

    $rows = @()
    for ($row = ($HeaderRow + 1); $row -le $lastRow; $row++) {
        $restaurant = Get-CellString $SourceWorksheet $row $restaurantCol
        $ownership = Get-CellString $SourceWorksheet $row $ownershipCol

        $wanted =
            ($ownership -eq 'All Franchisees' -and [string]::IsNullOrWhiteSpace($restaurant)) -or
            ($ownership -eq 'Sorensen' -and [string]::IsNullOrWhiteSpace($restaurant)) -or
            ($ownership -eq 'Sorensen' -and $restaurant -in @('9354 Richmond', '9355 Virginia Beach'))

        if (-not $wanted) { continue }

        $metricValues = @()
        foreach ($header in $requiredSourceHeaders[2..($requiredSourceHeaders.Count - 1)]) {
            $metricValues += $SourceWorksheet.Cells.Item($row, $HeaderMap[(Normalize-Header $header)]).Value2
        }

        $rows += [pscustomobject]@{
            SourceFile = $SourceFileName
            Week = $WeekEnding
            Restaurant = if ([string]::IsNullOrWhiteSpace($restaurant)) { $null } else { $restaurant }
            Ownership = $ownership
            Metrics = $metricValues
        }
    }

    $expectedKeys = @(
        'All Franchisees|(TOTAL)',
        'Sorensen|(TOTAL)',
        'Sorensen|9354 Richmond',
        'Sorensen|9355 Virginia Beach'
    )
    $actualKeys = $rows | ForEach-Object {
        if ($null -eq $_.Restaurant) { "$($_.Ownership)|(TOTAL)" } else { "$($_.Ownership)|$($_.Restaurant)" }
    }
    $missingKeys = $expectedKeys | Where-Object { $_ -notin $actualKeys }
    if ($missingKeys.Count -gt 0) {
        throw "Source workbook did not contain all required entity rows. Missing: $($missingKeys -join ', ')"
    }

    return $rows
}

function Get-SourceRowsFromFile {
    param(
        [object]$Excel,
        [object]$Source
    )

    $sourceWb = $null
    $sourceWs = $null
    try {
        $sourceWb = $Excel.Workbooks.Open($Source.File.FullName, 0, $true)
        $sourceWs = $sourceWb.Worksheets.Item(1)
        $sourceHeaders = Get-HeaderMap $sourceWs $Source.HeaderRow
        return Get-SourceRows $sourceWs $sourceHeaders $Source.HeaderRow $Source.WeekEnding $Source.SourceFileName
    }
    finally {
        if ($sourceWb) { $sourceWb.Close($false) }
        Release-ComObject $sourceWs
        Release-ComObject $sourceWb
    }
}

function Get-WeekRowCount {
    param([object]$RawDataWorksheet, [datetime]$WeekEnding)

    $lastRow = $RawDataWorksheet.Cells.Item($RawDataWorksheet.Rows.Count, 1).End($xlUp).Row
    $count = 0
    for ($row = 2; $row -le $lastRow; $row++) {
        $week = Convert-ExcelDate $RawDataWorksheet.Cells.Item($row, 2).Value2
        if ($week -and $week.Date -eq $WeekEnding.Date) {
            $count++
        }
    }
    return $count
}

function Test-WeekExists {
    param([object]$RawDataWorksheet, [datetime]$WeekEnding)
    return (Get-WeekRowCount $RawDataWorksheet $WeekEnding) -gt 0
}

function Get-InsertRowForWeek {
    param([object]$RawDataWorksheet, [datetime]$WeekEnding)

    $lastRow = $RawDataWorksheet.Cells.Item($RawDataWorksheet.Rows.Count, 1).End($xlUp).Row
    for ($row = 2; $row -le $lastRow; $row++) {
        $week = Convert-ExcelDate $RawDataWorksheet.Cells.Item($row, 2).Value2
        if ($week -and $week.Date -gt $WeekEnding.Date) {
            return $row
        }
    }
    return ($lastRow + 1)
}

function Add-RawDataRows {
    param(
        [object]$RawDataWorksheet,
        [object[]]$RowsToAppend
    )

    $rows = @($RowsToAppend)
    if ($rows.Count -eq 0) { return 0 }

    $weekEnding = $rows[0].Week
    $insertRow = Get-InsertRowForWeek $RawDataWorksheet $weekEnding
    $lastRawRow = $RawDataWorksheet.Cells.Item($RawDataWorksheet.Rows.Count, 1).End($xlUp).Row
    $rowCount = $rows.Count
    $endInsertRow = $insertRow + $rowCount - 1

    if ($insertRow -le $lastRawRow) {
        $RawDataWorksheet.Range("A$insertRow`:A$endInsertRow").EntireRow.Insert() | Out-Null
    }

    $templateRow = if ($insertRow -gt 2) { $insertRow - 1 } else { $endInsertRow + 1 }
    $rowsWritten = 0
    foreach ($sourceRow in $rows) {
        $targetRow = $insertRow + $rowsWritten

        $RawDataWorksheet.Range("A$templateRow`:S$templateRow").Copy() | Out-Null
        $RawDataWorksheet.Range("A$targetRow`:S$targetRow").PasteSpecial($xlPasteFormats) | Out-Null

        $values = @(
            $sourceRow.SourceFile,
            $sourceRow.Week.ToOADate(),
            $sourceRow.Restaurant,
            $sourceRow.Ownership
        ) + $sourceRow.Metrics

        $rowMatrix = New-Object 'object[,]' 1, 17
        for ($col = 1; $col -le 17; $col++) {
            $rowMatrix[0, ($col - 1)] = $values[$col - 1]
        }
        $RawDataWorksheet.Range("A$targetRow`:Q$targetRow").Value2 = $rowMatrix

        $RawDataWorksheet.Cells.Item($targetRow, 2).NumberFormat = 'm/d/yyyy'
        $RawDataWorksheet.Cells.Item($targetRow, 18).Formula = '=D' + $targetRow + '&"|"&IF(C' + $targetRow + '="","(TOTAL)",TRIM(C' + $targetRow + '))'
        $RawDataWorksheet.Cells.Item($targetRow, 19).Formula = '=R' + $targetRow + '&"|"&TEXT(B' + $targetRow + ',"yyyymmdd")'

        $rowsWritten++
    }

    return $rowsWritten
}

function Get-EntityKey {
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

function Get-EntityWeekKey {
    param([string]$EntityKey, [datetime]$WeekEnding)
    return "$EntityKey|$($WeekEnding.ToString('yyyyMMdd'))"
}

function Get-MetricCatalogRows {
    param([object]$MetricCatalogWorksheet)

    $lastRow = $MetricCatalogWorksheet.Cells.Item($MetricCatalogWorksheet.Rows.Count, 1).End($xlUp).Row
    $metrics = @()
    for ($row = 2; $row -le $lastRow; $row++) {
        $displayName = Get-CellString $MetricCatalogWorksheet $row 1
        $rawKey = Get-CellString $MetricCatalogWorksheet $row 2
        if ([string]::IsNullOrWhiteSpace($displayName) -or [string]::IsNullOrWhiteSpace($rawKey)) {
            continue
        }

        $metrics += [pscustomobject]@{
            DisplayName = $displayName
            RawKey = $rawKey
            LowerIsBetter = Convert-ToBool $MetricCatalogWorksheet.Cells.Item($row, 3).Value2
            IncludeInMovers = Convert-ToBool $MetricCatalogWorksheet.Cells.Item($row, 4).Value2
            Category = Get-CellString $MetricCatalogWorksheet $row 5
            Owner = Get-CellString $MetricCatalogWorksheet $row 6
        }
    }

    return $metrics
}

function Get-RawDataIndex {
    param([object]$RawDataWorksheet)

    $headers = Get-HeaderMap $RawDataWorksheet 1
    $requiredHeaders = @('Week', 'Restaurant', 'Ownership', 'Count')
    foreach ($header in $requiredHeaders) {
        $key = Normalize-Header $header
        if (-not $headers.ContainsKey($key)) {
            throw "Raw_Data is missing required header '$header'."
        }
    }

    $weekCol = $headers[(Normalize-Header 'Week')]
    $restaurantCol = $headers[(Normalize-Header 'Restaurant')]
    $ownershipCol = $headers[(Normalize-Header 'Ownership')]
    $lastRow = $RawDataWorksheet.Cells.Item($RawDataWorksheet.Rows.Count, 1).End($xlUp).Row
    $rowsByKey = @{}

    for ($row = 2; $row -le $lastRow; $row++) {
        $week = Convert-ExcelDate $RawDataWorksheet.Cells.Item($row, $weekCol).Value2
        if ($null -eq $week) { continue }

        $ownership = Get-CellString $RawDataWorksheet $row $ownershipCol
        if ([string]::IsNullOrWhiteSpace($ownership)) { continue }

        $restaurant = Get-CellString $RawDataWorksheet $row $restaurantCol
        $entityKey = Get-EntityKey $ownership $restaurant
        $rowsByKey[(Get-EntityWeekKey $entityKey $week)] = $row
    }

    return [pscustomobject]@{
        Headers = $headers
        RowsByKey = $rowsByKey
    }
}

function Get-RawMetricValue {
    param(
        [object]$RawDataWorksheet,
        [object]$RawIndex,
        [string]$EntityKey,
        [datetime]$WeekEnding,
        [string]$MetricKey
    )

    $rowKey = Get-EntityWeekKey $EntityKey $WeekEnding
    if (-not $RawIndex.RowsByKey.ContainsKey($rowKey)) {
        return $null
    }

    $headerKey = Normalize-Header $MetricKey
    if (-not $RawIndex.Headers.ContainsKey($headerKey)) {
        throw "Raw_Data is missing metric column '$MetricKey'."
    }

    $row = $RawIndex.RowsByKey[$rowKey]
    $col = $RawIndex.Headers[$headerKey]
    return Convert-ToNullableDouble $RawDataWorksheet.Cells.Item($row, $col).Value2
}

function Get-AverageValue {
    param([object[]]$Values)

    $usable = @($Values | Where-Object { $null -ne $_ })
    if ($usable.Count -eq 0) { return $null }
    return ($usable | Measure-Object -Average).Average
}

function Get-Improvement {
    param([object]$CurrentValue, [object]$PriorValue, [bool]$LowerIsBetter)

    if ($null -eq $CurrentValue -or $null -eq $PriorValue) { return $null }
    if ($LowerIsBetter) {
        return ([double]$PriorValue - [double]$CurrentValue)
    }
    return ([double]$CurrentValue - [double]$PriorValue)
}

function Get-ThreeWeekTrend {
    param(
        [object]$RawDataWorksheet,
        [object]$RawIndex,
        [string]$EntityKey,
        [datetime]$LatestWeek,
        [string]$MetricKey,
        [bool]$LowerIsBetter,
        [bool]$IncludeInMovers
    )

    if (-not $IncludeInMovers) {
        return [pscustomobject]@{ Arrow = ''; Average = $null; Persistent = '' }
    }

    $improvements = @()
    foreach ($weekOffset in @(0, -7, -14)) {
        $week = $LatestWeek.AddDays($weekOffset)
        $current = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $week $MetricKey
        $prior = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $week.AddDays(-7) $MetricKey
        $improvements += Get-Improvement $current $prior $LowerIsBetter
    }

    $average = Get-AverageValue $improvements
    if ($null -eq $average -or [math]::Abs($average) -lt 0.0000001) {
        return [pscustomobject]@{ Arrow = ''; Average = $average; Persistent = '' }
    }

    $arrow = if ($average -gt 0) { [string][char]0x25B2 } else { [string][char]0x25BC }
    $usable = @($improvements | Where-Object { $null -ne $_ })
    $persistent = ''
    if ($usable.Count -eq 3) {
        $positiveCount = @($usable | Where-Object { $_ -gt 0 }).Count
        $negativeCount = @($usable | Where-Object { $_ -lt 0 }).Count
        if ($positiveCount -eq 3 -or $negativeCount -eq 3) {
            $persistent = $average
        }
    }

    return [pscustomobject]@{ Arrow = $arrow; Average = $average; Persistent = $persistent }
}

function Update-EmailComparisonSection {
    param(
        [object]$Worksheet,
        [object]$RawDataWorksheet,
        [object]$RawIndex,
        [object[]]$Metrics,
        [datetime]$LatestWeek,
        [string]$EntityKey,
        [string]$EntityName,
        [int]$HeaderRow,
        [int]$FirstMetricRow,
        [string]$CompareEntityKey,
        [string]$CompareEntityName
    )

    Set-CellValue $Worksheet $HeaderRow 1 'Entity'
    Set-CellValue $Worksheet $HeaderRow 2 $EntityName
    Set-CellValue $Worksheet $HeaderRow 4 'As Of Week'
    Set-CellValue $Worksheet $HeaderRow 5 $LatestWeek.ToOADate()
    $Worksheet.Cells.Item($HeaderRow, 5).NumberFormat = 'm/d/yyyy'
    Set-CellValue $Worksheet $HeaderRow 8 $CompareEntityKey
    Set-CellValue $Worksheet $HeaderRow 18 $LatestWeek.ToOADate()
    $Worksheet.Cells.Item($HeaderRow, 18).NumberFormat = 'm/d/yyyy'
    Set-CellValue $Worksheet ($HeaderRow + 1) 1 'Compare To'
    Set-CellValue $Worksheet ($HeaderRow + 1) 2 $CompareEntityName

    $latestMinus7 = $LatestWeek.AddDays(-7)
    $latestMinus14 = $LatestWeek.AddDays(-14)
    $latestMinus21 = $LatestWeek.AddDays(-21)
    $priorYearWeek = $LatestWeek.AddDays(-364)
    $currentCount = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $LatestWeek 'Count'

    for ($i = 0; $i -lt $Metrics.Count; $i++) {
        $metric = $Metrics[$i]
        $row = $FirstMetricRow + $i
        $current = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $LatestWeek $metric.RawKey
        $priorWeek = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $latestMinus7 $metric.RawKey
        $priorYear = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $priorYearWeek $metric.RawKey
        $twoWeeksAgo = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $latestMinus14 $metric.RawKey
        $threeWeeksAgo = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $latestMinus21 $metric.RawKey
        $rollingAverage = Get-AverageValue @($current, $priorWeek, $twoWeeksAgo, $threeWeeksAgo)
        $wowImprovement = if ($metric.DisplayName -eq 'Count') { $null } else { Get-Improvement $current $priorWeek $metric.LowerIsBetter }
        $yoyImprovement = Get-Improvement $current $priorYear $metric.LowerIsBetter

        $compareCurrent = Get-RawMetricValue $RawDataWorksheet $RawIndex $CompareEntityKey $LatestWeek $metric.RawKey
        $comparePriorYear = Get-RawMetricValue $RawDataWorksheet $RawIndex $CompareEntityKey $priorYearWeek $metric.RawKey
        $compareYoyImprovement = Get-Improvement $compareCurrent $comparePriorYear $metric.LowerIsBetter
        $vsCompareCurrent = Get-Improvement $current $compareCurrent $metric.LowerIsBetter
        $vsCompareYoy = if ($null -eq $yoyImprovement -or $null -eq $compareYoyImprovement) { $null } else { $yoyImprovement - $compareYoyImprovement }
        $trend = Get-ThreeWeekTrend $RawDataWorksheet $RawIndex $EntityKey $LatestWeek $metric.RawKey $metric.LowerIsBetter $metric.IncludeInMovers
        $currentBadness = if ($null -eq $current -or $metric.DisplayName -eq 'Count') { $null } elseif ($metric.LowerIsBetter) { $current } else { 100 - $current }
        $yoyBadness = if ($null -eq $yoyImprovement -or $yoyImprovement -ge 0) { $null } else { -1 * $yoyImprovement }

        Set-CellValue $Worksheet $row 1 $metric.DisplayName
        Set-CellValue $Worksheet $row 2 $current
        Set-CellValue $Worksheet $row 3 $wowImprovement
        Set-CellValue $Worksheet $row 4 $yoyImprovement
        Set-CellValue $Worksheet $row 5 $rollingAverage
        Set-CellValue $Worksheet $row 6 $currentCount
        Set-CellValue $Worksheet $row 7 $metric.RawKey
        Set-CellValue $Worksheet $row 8 $metric.LowerIsBetter
        Set-CellValue $Worksheet $row 9 $metric.IncludeInMovers
        Set-CellValue $Worksheet $row 10 $priorWeek
        Set-CellValue $Worksheet $row 11 $priorYear
        Set-CellValue $Worksheet $row 12 $twoWeeksAgo
        Set-CellValue $Worksheet $row 13 $threeWeeksAgo
        Set-CellValue $Worksheet $row 14 $metric.Category
        Set-CellValue $Worksheet $row 15 $metric.Owner
        Set-CellValue $Worksheet $row 16 ''
        Set-CellValue $Worksheet $row 17 ''
        Set-CellValue $Worksheet $row 18 $trend.Arrow
        Set-CellValue $Worksheet $row 19 $yoyBadness
        Set-CellValue $Worksheet $row 20 $currentBadness
        Set-CellValue $Worksheet $row 21 ''
        Set-CellValue $Worksheet $row 22 ''
        Set-CellValue $Worksheet $row 23 $trend.Average
        Set-CellValue $Worksheet $row 24 $trend.Persistent

        if ($metric.DisplayName -eq 'Count') {
            Set-CellValue $Worksheet $row 25 ''
            Set-CellValue $Worksheet $row 26 ''
            Set-CellValue $Worksheet $row 27 ''
            Set-CellValue $Worksheet $row 28 ''
        }
        else {
            Set-CellValue $Worksheet $row 25 $compareCurrent
            Set-CellValue $Worksheet $row 26 $compareYoyImprovement
            Set-CellValue $Worksheet $row 27 $vsCompareCurrent
            Set-CellValue $Worksheet $row 28 $vsCompareYoy
        }
    }
}

function Update-EmailComparisonAndExport {
    param(
        [object]$Workbook,
        [object]$RawDataWorksheet,
        [object]$LatestSource,
        [string]$FolderPath,
        [string]$TestDir,
        [bool]$ApplyMode
    )

    $emailWs = $null
    $metricWs = $null
    try {
        $emailWs = $Workbook.Worksheets.Item('Email Comparison')
        $metricWs = $Workbook.Worksheets.Item('Metric_Catalog')
        $metrics = @(Get-MetricCatalogRows $metricWs)
        if ($metrics.Count -eq 0) {
            throw 'Metric_Catalog did not contain any metrics for the email comparison sheet.'
        }

        $rawIndex = Get-RawDataIndex $RawDataWorksheet
        $compareEntityKey = 'All Franchisees|(TOTAL)'
        $compareEntityName = 'TOTAL - All Franchisees'
        $latestWeek = $LatestSource.WeekEnding

        Update-EmailComparisonSection $emailWs $RawDataWorksheet $rawIndex $metrics $latestWeek 'Sorensen|9355 Virginia Beach' '9355 Virginia Beach (Sorensen)' 3 6 $compareEntityKey $compareEntityName
        Update-EmailComparisonSection $emailWs $RawDataWorksheet $rawIndex $metrics $latestWeek 'Sorensen|9354 Richmond' '9354 Richmond (Sorensen)' 23 26 $compareEntityKey $compareEntityName

        $emailWs.PageSetup.PrintArea = '$A$1:$AB$38'
        $emailWs.PageSetup.Orientation = 2
        $emailWs.PageSetup.Zoom = $false
        $emailWs.PageSetup.FitToPagesWide = 1
        $emailWs.PageSetup.FitToPagesTall = 1

        $pdfDir = if ($ApplyMode) { Join-Path $FolderPath '04 Email Comparison PDFs' } else { $TestDir }
        New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
        $pdfName = 'GSS Email Comparison {0}.pdf' -f $LatestSource.File.LastWriteTime.ToString('MMddyy')
        $pdfPath = Join-Path $pdfDir $pdfName
        $emailWs.ExportAsFixedFormat($xlTypePDF, $pdfPath)
        return $pdfPath
    }
    finally {
        Release-ComObject $metricWs
        Release-ComObject $emailWs
    }
}

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $scriptRoot = Split-Path -Parent $PSCommandPath
    $projectRoot = Split-Path -Parent $scriptRoot
    $Folder = Split-Path -Parent $projectRoot
}

$Folder = (Resolve-Path -LiteralPath $Folder).Path
$mainPath = Resolve-MainWorkbookPath $Folder $MainWorkbookName

$automationDir = Join-Path $Folder '_automation_runs'
$logDir = Join-Path $automationDir 'logs'
$backupDir = Join-Path $automationDir 'backups'
$testDir = Join-Path $automationDir 'test-output'
New-Item -ItemType Directory -Path $logDir, $backupDir, $testDir -Force | Out-Null

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$excel = $null
$sourceWb = $null
$sourceWs = $null
$targetWb = $null
$rawWs = $null
$targetPath = $mainPath
$backupPath = $null
$emailComparisonPdfPath = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    try {
        $excel.Calculation = $xlCalculationAutomatic
    }
    catch {
        Write-Warning "Excel did not accept the automatic calculation setting before opening a workbook. Continuing; the updater will still calculate before saving."
    }

    $allSources = @(Get-RollingSources $excel $Folder)
    $latestSource = Select-LatestRollingSource $allSources
    if ($null -eq $latestSource) {
        throw "No latest rolling Sorensen source workbook could be selected from $Folder."
    }

    $priorYearWeek = $latestSource.WeekEnding.AddDays(-364)
    $priorYearSource = Select-RollingSourceByWeek $allSources $priorYearWeek
    if ($null -eq $priorYearSource) {
        throw "No matching prior-year source workbook found for week ending $($priorYearWeek.ToString('yyyy-MM-dd'))."
    }

    $requestedSources = @($priorYearSource, $latestSource) |
        Sort-Object WeekEnding |
        ForEach-Object {
            [pscustomobject]@{
                Role = if ($_.WeekEnding.Date -eq $latestSource.WeekEnding.Date) { 'CurrentWeek' } else { 'PriorYearWeek' }
                Source = $_
            }
        }

    if (-not $Apply) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($MainWorkbookName)
        $targetPath = Join-Path $testDir "$base`_TEST_$timestamp.xlsx"
        Copy-Item -LiteralPath $mainPath -Destination $targetPath -Force
    }

    $targetWb = $excel.Workbooks.Open($targetPath, 0, $false)
    if ($Apply -and $targetWb.ReadOnly) {
        throw "Main workbook opened read-only. Close it in Excel/Dropbox and run again: $mainPath"
    }

    $rawWs = $targetWb.Worksheets.Item('Raw_Data')

    $status = 'Updated'
    $rowsAppended = 0
    $weeksAppended = @()
    $weeksSkipped = @()
    $sourceWork = @()

    foreach ($requested in $requestedSources) {
        $source = $requested.Source
        $existingCount = Get-WeekRowCount $rawWs $source.WeekEnding
        if ($existingCount -gt 0) {
            $weeksSkipped += [pscustomobject]@{
                Role = $requested.Role
                WeekEnding = $source.WeekEnding.ToString('yyyy-MM-dd')
                ExistingRows = $existingCount
                SourceWorkbook = $source.File.FullName
            }
            continue
        }

        $sourceWork += $requested
    }

    if ($Apply) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($MainWorkbookName)
        $backupPath = Join-Path $backupDir "$base`_BACKUP_$timestamp.xlsx"
        $targetWb.SaveCopyAs($backupPath)
    }

    if ($sourceWork.Count -eq 0) {
        $status = 'WorkbookAlreadyCurrentEmailPdfRefreshed'
    }
    else {
        foreach ($workItem in $sourceWork) {
            $source = $workItem.Source
            $rowsForSource = @(Get-SourceRowsFromFile $excel $source)
            $written = Add-RawDataRows $rawWs $rowsForSource
            $rowsAppended += $written
            $weeksAppended += [pscustomobject]@{
                Role = $workItem.Role
                WeekEnding = $source.WeekEnding.ToString('yyyy-MM-dd')
                SourceWorkbook = $source.File.FullName
                SourceFileNameInRawData = $source.SourceFileName
                RowsAppended = $written
            }
        }
    }

    $excel.CalculateFullRebuild()
    $emailComparisonPdfPath = Update-EmailComparisonAndExport $targetWb $rawWs $latestSource $Folder $testDir ([bool]$Apply)
    $targetWb.Save()

    $summary = [pscustomobject]@{
        Timestamp = (Get-Date).ToString('s')
        Mode = if ($Apply) { 'ApplyToMainWorkbook' } else { 'CopyTestOnly' }
        Status = $status
        Folder = $Folder
        CurrentWeekEnding = $latestSource.WeekEnding.ToString('yyyy-MM-dd')
        PriorYearWeekEnding = $priorYearWeek.ToString('yyyy-MM-dd')
        CurrentSourceWorkbook = $latestSource.File.FullName
        PriorYearSourceWorkbook = $priorYearSource.File.FullName
        TargetWorkbook = $targetPath
        BackupWorkbook = $backupPath
        EmailComparisonPdf = $emailComparisonPdfPath
        RowsAppended = $rowsAppended
        WeeksAppended = $weeksAppended
        WeeksSkipped = $weeksSkipped
    }

    $logPath = Join-Path $logDir "gss_update_$timestamp.json"
    $summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $logPath -Encoding UTF8
    $summary | Add-Member -NotePropertyName LogPath -NotePropertyValue $logPath -PassThru
}
finally {
    if ($targetWb) { $targetWb.Close($false) }
    if ($sourceWb) { $sourceWb.Close($false) }
    Release-ComObject $rawWs
    Release-ComObject $targetWb
    Release-ComObject $sourceWs
    Release-ComObject $sourceWb
    if ($excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
