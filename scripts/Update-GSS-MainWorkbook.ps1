[CmdletBinding()]
param(
    [string]$Folder,
    [string]$MainWorkbookName = 'GSS Score Trends - Main.xlsx',
    [switch]$Apply,
    [switch]$OutputObject
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Gss-Common.ps1')

$xlUp = -4162
$xlToLeft = -4159
$xlPasteFormats = -4122
$xlCalculationAutomatic = -4105
$xlTypePDF = 0
$xlPortrait = 1
$xlLandscape = 2
$xlSheetVisible = -1
$xlValidateList = 3
$xlValidAlertStop = 1
$xlBetween = 1
$xlUnlockedCells = 1
$gssRawDataLastFormulaRow = 5096

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
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

function Set-CellFormula {
    param([object]$Worksheet, [int]$Row, [int]$Column, [string]$Formula)

    $cell = $null
    try {
        $cell = $Worksheet.Cells.Item($Row, $Column)
        try {
            $cell.Formula2 = $Formula
        }
        catch {
            $cell.Formula = $Formula
        }
    }
    finally {
        Release-ComObject $cell
    }
}

function Get-GssExcelColor {
    param([int]$Red, [int]$Green, [int]$Blue)
    return ($Red + (256 * $Green) + (65536 * $Blue))
}

function Get-GssQaCheckPlan {
    param(
        [int]$EntityLastRow = 5,
        [int]$MetricLastRow = 14,
        [int]$RawDataLastRow = $gssRawDataLastFormulaRow
    )

    $activeCount = 'SUMPRODUCT(--(UPPER(''Entities''!$E$2:$E${ENTITY_LAST}&"")="TRUE"))'.Replace('${ENTITY_LAST}', [string]$EntityLastRow)
    $entityRows = 'SUMPRODUCT(--(UPPER(''Entities''!$E$2:$E${ENTITY_LAST}&"")="TRUE"),--(COUNTIFS(''Raw_Data''!$R$2:$R${RAW_LAST},''Entities''!$D$2:$D${ENTITY_LAST},''Raw_Data''!$B$2:$B${RAW_LAST},${WEEK})=1))'.Replace('${ENTITY_LAST}', [string]$EntityLastRow).Replace('${RAW_LAST}', [string]$RawDataLastRow)
    $blankMetrics = 'SUMPRODUCT(--(''Raw_Data''!$B$2:$B${RAW_LAST}=${WEEK}),MMULT(--(''Raw_Data''!$E$2:$Q${RAW_LAST}=""),TRANSPOSE(COLUMN(''Raw_Data''!$E$1:$Q$1)^0)))'.Replace('${RAW_LAST}', [string]$RawDataLastRow)

    $checks = @()
    $checks += [pscustomobject]@{
        Row = 6
        Name = 'Duplicate Raw_Data RowKeys'
        Formula = '=IF(SUMPRODUCT(--(''Raw_Data''!$S$2:$S${RAW_LAST}<>""),--(COUNTIF(''Raw_Data''!$S$2:$S${RAW_LAST},''Raw_Data''!$S$2:$S${RAW_LAST})>1))=0,"READY","ATTENTION")'.Replace('${RAW_LAST}', [string]$RawDataLastRow)
        DetailFormula = '=SUMPRODUCT(--(''Raw_Data''!$S$2:$S${RAW_LAST}<>""),--(COUNTIF(''Raw_Data''!$S$2:$S${RAW_LAST},''Raw_Data''!$S$2:$S${RAW_LAST})>1))&" duplicate data row(s)"'.Replace('${RAW_LAST}', [string]$RawDataLastRow)
    }

    $periods = @(
        [pscustomobject]@{ Row = 7; Name = 'Current-week entity and metric completeness'; Week = '$B$3' },
        [pscustomobject]@{ Row = 8; Name = 'Prior-week entity and metric completeness'; Week = '$B$3-7' },
        [pscustomobject]@{ Row = 9; Name = 'Prior-year entity and metric completeness'; Week = '$B$3-364' }
    )
    foreach ($period in $periods) {
        $rowCheck = $entityRows.Replace('${WEEK}', $period.Week)
        $blankCheck = $blankMetrics.Replace('${WEEK}', $period.Week)
        $checks += [pscustomobject]@{
            Row = $period.Row
            Name = $period.Name
            Formula = '=IF(AND(${ROW_CHECK}=${ACTIVE_COUNT},${BLANK_CHECK}=0),"READY","ATTENTION")'.Replace('${ROW_CHECK}', $rowCheck).Replace('${ACTIVE_COUNT}', $activeCount).Replace('${BLANK_CHECK}', $blankCheck)
            DetailFormula = '=${ROW_CHECK}&" of "&${ACTIVE_COUNT}&" active entities; "&${BLANK_CHECK}&" blank required metric value(s)"'.Replace('${ROW_CHECK}', $rowCheck).Replace('${ACTIVE_COUNT}', $activeCount).Replace('${BLANK_CHECK}', $blankCheck)
        }
    }

    $checks += [pscustomobject]@{
        Row = 10
        Name = 'Required Raw_Data metric headers'
        Formula = '=IF(SUMPRODUCT(--(''Metric_Catalog''!$B$2:$B${METRIC_LAST}<>""),--ISNUMBER(MATCH(''Metric_Catalog''!$B$2:$B${METRIC_LAST},''Raw_Data''!$A$1:$S$1,0)))=COUNTA(''Metric_Catalog''!$B$2:$B${METRIC_LAST}),"READY","ATTENTION")'.Replace('${METRIC_LAST}', [string]$MetricLastRow)
        DetailFormula = '=(COUNTA(''Metric_Catalog''!$B$2:$B${METRIC_LAST})-SUMPRODUCT(--(''Metric_Catalog''!$B$2:$B${METRIC_LAST}<>""),--ISNUMBER(MATCH(''Metric_Catalog''!$B$2:$B${METRIC_LAST},''Raw_Data''!$A$1:$S$1,0))))&" missing required header(s)"'.Replace('${METRIC_LAST}', [string]$MetricLastRow)
    }
    $checks += [pscustomobject]@{
        Row = 11
        Name = 'Dashboard selector validity'
        Formula = '=IF(AND(SUMPRODUCT(--(''Entities''!$A$2:$A${ENTITY_LAST}=''Exec_Dashboard''!$C$3),--(UPPER(''Entities''!$E$2:$E${ENTITY_LAST}&"")="TRUE"))=1,COUNTIF(''Metric_Catalog''!$A$2:$A${METRIC_LAST},''Exec_Dashboard''!$F$3)=1),"READY","ATTENTION")'.Replace('${ENTITY_LAST}', [string]$EntityLastRow).Replace('${METRIC_LAST}', [string]$MetricLastRow)
        DetailFormula = '="Entity: "&''Exec_Dashboard''!$C$3&"; Metric: "&''Exec_Dashboard''!$F$3'
    }
    $checks += [pscustomobject]@{
        Row = 12
        Name = 'Remaining Raw_Data capacity'
        Formula = '=IF(ROWS(''Raw_Data''!$B$2:$B${RAW_LAST})-COUNT(''Raw_Data''!$B$2:$B${RAW_LAST})>=8,"READY","ATTENTION")'.Replace('${RAW_LAST}', [string]$RawDataLastRow)
        DetailFormula = '=(ROWS(''Raw_Data''!$B$2:$B${RAW_LAST})-COUNT(''Raw_Data''!$B$2:$B${RAW_LAST}))&" rows remain before the formula bound"'.Replace('${RAW_LAST}', [string]$RawDataLastRow)
    }

    return $checks
}

function Get-GssWorkbookLayoutPlan {
    param([int]$ActiveEntityCount, [int]$MetricCount)

    return @(
        [pscustomobject]@{ Sheet = 'README'; PrintArea = '$A$1:$A$34'; Orientation = $xlPortrait; FitWide = 1; FitTall = 2 },
        [pscustomobject]@{ Sheet = 'QA Checks'; PrintArea = '$A$1:$C$12'; Orientation = $xlPortrait; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Quick_Read_WoW'; PrintArea = ('$A$1:$Q${0}' -f (5 + $ActiveEntityCount)); Orientation = $xlLandscape; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Quick_Read_YoY'; PrintArea = ('$A$1:$Q${0}' -f (5 + $ActiveEntityCount)); Orientation = $xlLandscape; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Exec_Dashboard'; PrintArea = '$A$1:$AD$57'; Orientation = $xlLandscape; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Drivers_Detail'; PrintArea = ('$A$1:$AB${0}' -f (5 + $MetricCount)); Orientation = $xlLandscape; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Email Comparison'; PrintArea = '$A$1:$AB$38'; Orientation = $xlLandscape; FitWide = 1; FitTall = 1 }
    )
}

function Get-GssGuardrailPlan {
    return [pscustomobject]@{
        ProtectWorkbookStructure = $true
        ProtectAllWorksheets = $true
        UnlockedCells = @('Exec_Dashboard!C3', 'Exec_Dashboard!F3')
        ValidationShowError = $true
        IntentionalGapIsBlocker = $false
    }
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

    $sourceFolder = Resolve-GssChildFolder $FolderPath '02 Weekly Rolling Source Workbooks'
    $searchRoot = if ($sourceFolder) { $sourceFolder } else { $FolderPath }
    $candidates = @(Get-GssFiles $searchRoot 'Sorensen FW*.xlsx' |
        Where-Object { $_.Name -notmatch '(?i)CLEANED|Tidy|YY Trends' } |
        Sort-Object LastWriteTime -Descending)

    if ($candidates.Count -eq 0 -and $sourceFolder) {
        Write-Warning "No rolling Sorensen source workbooks were found in '02 Weekly Rolling Source Workbooks'. Falling back to a recursive compatibility search under $FolderPath."
        $candidates = @(Get-GssFiles $FolderPath 'Sorensen FW*.xlsx' |
            Where-Object { $_.Name -notmatch '(?i)CLEANED|Tidy|YY Trends' } |
            Sort-Object LastWriteTime -Descending)
    }

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

function Get-GssConfigurationCounts {
    param([object]$Workbook)

    $entityWs = $null
    $metricWs = $null
    try {
        $entityWs = $Workbook.Worksheets.Item('Entities')
        $metricWs = $Workbook.Worksheets.Item('Metric_Catalog')
        $entityLastRow = $entityWs.Cells.Item($entityWs.Rows.Count, 1).End($xlUp).Row
        $metricLastRow = $metricWs.Cells.Item($metricWs.Rows.Count, 1).End($xlUp).Row
        $activeEntityCount = 0
        for ($row = 2; $row -le $entityLastRow; $row++) {
            if ((Get-CellString $entityWs $row 1) -and (Convert-ToBool $entityWs.Cells.Item($row, 5).Value2)) {
                $activeEntityCount++
            }
        }

        $metricCount = 0
        for ($row = 2; $row -le $metricLastRow; $row++) {
            if ((Get-CellString $metricWs $row 1) -and (Get-CellString $metricWs $row 2)) {
                $metricCount++
            }
        }

        return [pscustomobject]@{
            EntityLastRow = $entityLastRow
            ActiveEntityCount = $activeEntityCount
            MetricLastRow = $metricLastRow
            MetricCount = $metricCount
        }
    }
    finally {
        Release-ComObject $metricWs
        Release-ComObject $entityWs
    }
}

function Remove-GssWorkbookGuardrails {
    param([object]$Workbook)

    if ($Workbook.ProtectStructure) {
        $Workbook.Unprotect('')
    }

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($index)
            if ($worksheet.ProtectContents -or $worksheet.ProtectDrawingObjects -or $worksheet.ProtectScenarios) {
                $worksheet.Unprotect('')
            }
        }
        finally {
            Release-ComObject $worksheet
        }
    }
}

function Set-GssReadmeContent {
    param([object]$Workbook)

    $worksheet = $null
    $sectionRange = $null
    $contentRange = $null
    try {
        $worksheet = $Workbook.Worksheets.Item('README')
        Set-CellValue $worksheet 5 1 '1) Save the newest Sorensen rolling survey workbook in 02 Weekly Rolling Source Workbooks.'
        Set-CellValue $worksheet 6 1 '2) Close this workbook, then double-click Run GSS Update After Upload.cmd in the GSS Surveys folder.'
        Set-CellValue $worksheet 7 1 '3) Wait for the copy-test to finish. Type APPLY only when the copy-test passes and you want to update the live workbook.'
        Set-CellValue $worksheet 8 1 '4) Reopen this workbook and check QA Checks. Use the reports when the overall status is READY.'
        Set-CellValue $worksheet 9 1 'The automation recalculates the workbook, refreshes the email PDF, protects formulas, and creates a backup before live changes. Do not paste into Raw_Data manually.'
        Set-CellValue $worksheet 12 1 '1) Quick_Read_WoW: review the latest week-over-week scorecard across all active entities.'
        Set-CellValue $worksheet 13 1 '2) Quick_Read_YoY: review the same view against the matching prior-year week.'
        Set-CellValue $worksheet 14 1 '3) Exec_Dashboard: use the Entity and Metric dropdowns to drill into one story.'
        Set-CellValue $worksheet 15 1 '4) Drivers_Detail: review every metric for the entity selected on Exec_Dashboard.'
        Set-CellValue $worksheet 18 1 '• Quick_Read_WoW: positive values mean improvement. A blank is surfaced by QA Checks when required source data is missing.'
        Set-CellValue $worksheet 19 1 '• Quick_Read_YoY: compares CurrentWeek to CurrentWeek-364. The confirmed 2025 history break is informational, not an error.'
        Set-CellValue $worksheet 20 1 '• Exec_Dashboard: selector-driven view. The same selections drive Drivers_Detail and the movers table.'
        Set-CellValue $worksheet 21 1 '• QA Checks: the visible READY/ATTENTION gate checks row keys, period completeness, headers, selectors, and remaining capacity.'
        Set-CellValue $worksheet 26 1 'Data hygiene rules (checked automatically):'
        Set-CellValue $worksheet 29 1 '• The break between the 7/6/2025 and 10/5/2025 data points is intentional and does not block READY status.'
        Set-CellValue $worksheet 31 1 '• ATTENTION status: stop using the reports and share the newest file from _automation_runs\logs with the workbook maintainer.'
        Set-CellValue $worksheet 32 1 '• Dropdown rejects a choice: select an active entity or a metric shown in the dropdown list.'
        Set-CellValue $worksheet 34 1 '• Current model capacity is Raw_Data rows 2:5096. QA Checks warns before fewer than two weekly loads (8 rows) remain.'

        Set-CellFormula $worksheet 3 1 '="Workbook status: "&''QA Checks''!$B$2&" | Latest data week: "&TEXT(''QA Checks''!$B$3,"m/d/yyyy")'

        $contentRange = $worksheet.Range('A1:A34')
        $contentRange.WrapText = $true
        $contentRange.VerticalAlignment = -4160
        $contentRange.Font.Name = 'Aptos'
        $contentRange.Font.Size = 10
        $worksheet.Columns.Item(1).ColumnWidth = 112
        $contentRange.Rows.AutoFit() | Out-Null

        $worksheet.Range('A1').Font.Size = 16
        $worksheet.Range('A1').Font.Bold = $true
        $worksheet.Range('A1').Font.Color = (Get-GssExcelColor 255 255 255)
        $worksheet.Range('A1').Interior.Color = (Get-GssExcelColor 31 78 121)
        foreach ($row in @(4, 11, 17, 23, 26, 30)) {
            $sectionRange = $worksheet.Range("A$row")
            $sectionRange.Font.Bold = $true
            $sectionRange.Interior.Color = (Get-GssExcelColor 221 235 247)
            Release-ComObject $sectionRange
            $sectionRange = $null
        }
    }
    finally {
        Release-ComObject $sectionRange
        Release-ComObject $contentRange
        Release-ComObject $worksheet
    }
}

function Set-GssBanner {
    param(
        [object]$Worksheet,
        [string]$Address,
        [string]$Formula,
        [switch]$Merge
    )

    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        if ($Merge) {
            if ($range.MergeCells) { $range.UnMerge() }
            $range.ClearContents() | Out-Null
            $range.Merge() | Out-Null
        }
        try {
            $range.Cells.Item(1, 1).Formula2 = $Formula
        }
        catch {
            $range.Cells.Item(1, 1).Formula = $Formula
        }
        $range.Font.Name = 'Aptos'
        $range.Font.Bold = $true
        $range.Font.Size = 10
        $range.HorizontalAlignment = -4131
        $range.VerticalAlignment = -4108
        $range.WrapText = $true
        $range.RowHeight = 24
    }
    finally {
        Release-ComObject $range
    }
}

function Set-GssWorkbookBanners {
    param([object]$Workbook)

    $worksheet = $null
    try {
        $worksheet = $Workbook.Worksheets.Item('Quick_Read_WoW')
        Set-GssBanner $worksheet 'A4:Q4' '="Workbook status: "&''QA Checks''!$B$2&" | Positive = improvement (direction-adjusted for dissatisfaction). Blank = missing data."' -Merge
        Release-ComObject $worksheet
        $worksheet = $Workbook.Worksheets.Item('Quick_Read_YoY')
        Set-GssBanner $worksheet 'A4:Q4' '="Workbook status: "&''QA Checks''!$B$2&" | Positive = improvement versus the matching prior-year week. Blank = missing data."' -Merge
        Release-ComObject $worksheet
        $worksheet = $Workbook.Worksheets.Item('Exec_Dashboard')
        Set-GssBanner $worksheet 'A2:F2' '="Workbook status: "&''QA Checks''!$B$2&" | Change the two dropdowns below to explore the reports."' -Merge
        Release-ComObject $worksheet
        $worksheet = $Workbook.Worksheets.Item('Drivers_Detail')
        Set-GssBanner $worksheet 'A2:D2' '="Workbook status: "&''QA Checks''!$B$2&" | This table follows the Executive Dashboard entity selection."' -Merge
        Release-ComObject $worksheet
        $worksheet = $Workbook.Worksheets.Item('Email Comparison')
        Set-GssBanner $worksheet 'A2:AB2' '="Workbook status: "&''QA Checks''!$B$2&" | Ready-to-share comparison output."' -Merge
    }
    finally {
        Release-ComObject $worksheet
    }
}

function Update-GssQaSheet {
    param(
        [object]$Workbook,
        [object]$Configuration
    )

    $qaWs = $null
    $readmeWs = $null
    $range = $null
    try {
        $readmeWs = $Workbook.Worksheets.Item('README')
        try {
            $qaWs = $Workbook.Worksheets.Item('QA Checks')
        }
        catch {
            $qaWs = $null
        }

        if ($qaWs -and $qaWs.Index -ne ($readmeWs.Index + 1)) {
            $qaWs.Delete()
            Release-ComObject $qaWs
            $qaWs = $null
        }
        if (-not $qaWs) {
            $qaWs = $Workbook.Worksheets.Add([System.Type]::Missing, $readmeWs)
            $qaWs.Name = 'QA Checks'
        }
        $qaWs.Visible = $xlSheetVisible
        $qaWs.Cells.Clear() | Out-Null

        Set-CellValue $qaWs 1 1 'GSS Workbook QA Checks'
        Set-CellValue $qaWs 2 1 'Overall workbook status'
        Set-CellFormula $qaWs 2 2 '=IFERROR(IF(COUNTIF($B$6:$B$12,"READY")=ROWS($B$6:$B$12),"READY","ATTENTION"),"ATTENTION")'
        Set-CellValue $qaWs 3 1 'Latest data week'
        Set-CellFormula $qaWs 3 2 "=MAX('Raw_Data'!`$B`$2:`$B`$$gssRawDataLastFormulaRow)"
        $qaWs.Cells.Item(3, 2).NumberFormat = 'm/d/yyyy'
        Set-CellValue $qaWs 4 1 'Confirmed history note'
        Set-CellValue $qaWs 4 2 'The break between the 7/6/2025 and 10/5/2025 data points is intentional and is not a blocker.'
        $qaWs.Range('B4:C4').Merge() | Out-Null
        Set-CellValue $qaWs 5 1 'Check'
        Set-CellValue $qaWs 5 2 'Status'
        Set-CellValue $qaWs 5 3 'Detail'

        foreach ($check in @(Get-GssQaCheckPlan $Configuration.EntityLastRow $Configuration.MetricLastRow $gssRawDataLastFormulaRow)) {
            Set-CellValue $qaWs $check.Row 1 $check.Name
            Set-CellFormula $qaWs $check.Row 2 $check.Formula
            Set-CellFormula $qaWs $check.Row 3 $check.DetailFormula
        }

        $range = $qaWs.Range('A1:C12')
        $range.Font.Name = 'Aptos'
        $range.Font.Size = 10
        $range.VerticalAlignment = -4108
        $range.WrapText = $true
        Release-ComObject $range
        $range = $qaWs.Range('A1:C1')
        $range.Merge() | Out-Null
        $range.Font.Size = 16
        $range.Font.Bold = $true
        $range.Font.Color = (Get-GssExcelColor 255 255 255)
        $range.Interior.Color = (Get-GssExcelColor 31 78 121)
        Release-ComObject $range
        $range = $qaWs.Range('A5:C5')
        $range.Font.Bold = $true
        $range.Font.Color = (Get-GssExcelColor 255 255 255)
        $range.Interior.Color = (Get-GssExcelColor 68 114 196)
        Release-ComObject $range
        $range = $qaWs.Range('A2:B2')
        $range.Font.Bold = $true
        $range.Font.Size = 12
        Release-ComObject $range
        $range = $qaWs.Range('A4:C4')
        $range.Interior.Color = (Get-GssExcelColor 255 242 204)
        Release-ComObject $range
        $range = $qaWs.Range('A6:C12')
        $range.Borders.LineStyle = 1
        Release-ComObject $range
        $qaWs.Columns.Item(1).ColumnWidth = 42
        $qaWs.Columns.Item(2).ColumnWidth = 18
        $qaWs.Columns.Item(3).ColumnWidth = 76
        $qaWs.Rows.Item('1:12').AutoFit() | Out-Null
        $qaWs.Tab.Color = (Get-GssExcelColor 112 173 71)
    }
    finally {
        Release-ComObject $range
        Release-ComObject $readmeWs
        Release-ComObject $qaWs
    }
}

function Set-GssWorkbookPrintLayouts {
    param([object]$Workbook, [object]$Configuration)

    foreach ($layout in @(Get-GssWorkbookLayoutPlan $Configuration.ActiveEntityCount $Configuration.MetricCount)) {
        $worksheet = $null
        $pageSetup = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($layout.Sheet)
            $pageSetup = $worksheet.PageSetup
            $pageSetup.PrintArea = $layout.PrintArea
            $pageSetup.Orientation = $layout.Orientation
            $pageSetup.Zoom = $false
            $pageSetup.FitToPagesWide = $layout.FitWide
            $pageSetup.FitToPagesTall = $layout.FitTall
            $pageSetup.PrintGridlines = $false
            $pageSetup.CenterHorizontally = $true
            $pageSetup.CenterVertically = $false
        }
        finally {
            Release-ComObject $pageSetup
            Release-ComObject $worksheet
        }
    }
}

function Set-GssDashboardValidation {
    param([object]$Workbook, [object]$Configuration)

    $worksheet = $null
    $cell = $null
    $validation = $null
    try {
        $worksheet = $Workbook.Worksheets.Item('Exec_Dashboard')
        foreach ($definition in @(
            [pscustomobject]@{ Address = 'C3'; Formula = "=Entities!`$A`$2:`$A`$$($Configuration.EntityLastRow)"; Message = 'Choose an active entity from the dropdown.' },
            [pscustomobject]@{ Address = 'F3'; Formula = "=Metric_Catalog!`$A`$2:`$A`$$($Configuration.MetricLastRow)"; Message = 'Choose a metric from the dropdown.' }
        )) {
            $cell = $worksheet.Range($definition.Address)
            $cell.Locked = $false
            $validation = $cell.Validation
            try { $validation.Delete() } catch { }
            $validation.Add($xlValidateList, $xlValidAlertStop, $xlBetween, $definition.Formula)
            $validation.IgnoreBlank = $false
            $validation.InCellDropdown = $true
            $validation.ShowInput = $true
            $validation.ShowError = $true
            $validation.InputTitle = 'Dashboard selection'
            $validation.InputMessage = $definition.Message
            $validation.ErrorTitle = 'Invalid dashboard selection'
            $validation.ErrorMessage = $definition.Message
            Release-ComObject $validation
            $validation = $null
            Release-ComObject $cell
            $cell = $null
        }
    }
    finally {
        Release-ComObject $validation
        Release-ComObject $cell
        Release-ComObject $worksheet
    }
}

function Set-GssStatusStyles {
    param([object]$Workbook, [string]$WorkbookStatus)

    $isReady = $WorkbookStatus -eq 'READY'
    $fillColor = if ($isReady) { Get-GssExcelColor 226 239 218 } else { Get-GssExcelColor 252 228 214 }
    $fontColor = if ($isReady) { Get-GssExcelColor 0 97 0 } else { Get-GssExcelColor 156 0 6 }
    foreach ($definition in @(
        [pscustomobject]@{ Sheet = 'README'; Address = 'A3' },
        [pscustomobject]@{ Sheet = 'QA Checks'; Address = 'A2:B2' },
        [pscustomobject]@{ Sheet = 'Quick_Read_WoW'; Address = 'A4:Q4' },
        [pscustomobject]@{ Sheet = 'Quick_Read_YoY'; Address = 'A4:Q4' },
        [pscustomobject]@{ Sheet = 'Exec_Dashboard'; Address = 'A2:F2' },
        [pscustomobject]@{ Sheet = 'Drivers_Detail'; Address = 'A2:D2' },
        [pscustomobject]@{ Sheet = 'Email Comparison'; Address = 'A2:AB2' }
    )) {
        $worksheet = $null
        $range = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($definition.Sheet)
            $range = $worksheet.Range($definition.Address)
            $range.Interior.Color = $fillColor
            $range.Font.Color = $fontColor
            $range.Font.Bold = $true
        }
        finally {
            Release-ComObject $range
            Release-ComObject $worksheet
        }
    }
}

function Get-GssWorkbookStatus {
    param([object]$Workbook)

    $worksheet = $null
    $cell = $null
    try {
        $worksheet = $Workbook.Worksheets.Item('QA Checks')
        $cell = $worksheet.Range('B2')
        return ([string]$cell.Value2).Trim().ToUpperInvariant()
    }
    finally {
        Release-ComObject $cell
        Release-ComObject $worksheet
    }
}

function Set-GssWorkbookGuardrails {
    param([object]$Workbook, [object]$Configuration)

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $null
        $usedRange = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($index)
            $usedRange = $worksheet.UsedRange
            $usedRange.Locked = $true
        }
        finally {
            Release-ComObject $usedRange
            Release-ComObject $worksheet
        }
    }

    Set-GssDashboardValidation $Workbook $Configuration

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($index)
            $worksheet.Protect('', $true, $true, $true, $true, $false, $false, $false, $false, $false, $false, $false, $false, $false, $true, $false)
            $worksheet.EnableSelection = if ($worksheet.Name -eq 'Exec_Dashboard') { $xlUnlockedCells } else { 0 }
        }
        finally {
            Release-ComObject $worksheet
        }
    }

    $Workbook.Protect('', $true, $false)
}

function Test-GssWorkbookGuardrailsApplied {
    param([object]$Workbook)

    if (-not $Workbook.ProtectStructure) { return $false }
    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($index)
            if (-not $worksheet.ProtectContents) { return $false }
        }
        finally {
            Release-ComObject $worksheet
        }
    }

    $execWs = $null
    $entityCell = $null
    $metricCell = $null
    try {
        $execWs = $Workbook.Worksheets.Item('Exec_Dashboard')
        $entityCell = $execWs.Range('C3')
        $metricCell = $execWs.Range('F3')
        return (-not [bool]$entityCell.Locked) -and (-not [bool]$metricCell.Locked)
    }
    finally {
        Release-ComObject $metricCell
        Release-ComObject $entityCell
        Release-ComObject $execWs
    }
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

function Update-EmailComparisonData {
    param(
        [object]$Workbook,
        [object]$RawDataWorksheet,
        [object]$LatestSource
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
    }
    finally {
        Release-ComObject $metricWs
        Release-ComObject $emailWs
    }
}

function Export-EmailComparisonPdf {
    param(
        [object]$Workbook,
        [object]$LatestSource,
        [string]$FolderPath,
        [string]$TestDir,
        [bool]$ApplyMode
    )

    $emailWs = $null
    try {
        $emailWs = $Workbook.Worksheets.Item('Email Comparison')
        $pdfDir = if ($ApplyMode) { Join-Path $FolderPath '04 Email Comparison PDFs' } else { $TestDir }
        New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
        $pdfName = 'GSS Email Comparison {0}.pdf' -f $LatestSource.File.LastWriteTime.ToString('MMddyy')
        $pdfPath = Join-Path $pdfDir $pdfName
        $emailWs.ExportAsFixedFormat($xlTypePDF, $pdfPath)
        return $pdfPath
    }
    finally {
        Release-ComObject $emailWs
    }
}

function Write-RunSummary {
    param([object]$Summary)

    Write-Host ''
    Write-Host 'GSS workbook update summary'
    Write-Host ('  Mode: {0}' -f $Summary.Mode)
    Write-Host ('  Status: {0}' -f $Summary.Status)
    Write-Host ('  Workbook QA: {0}' -f $Summary.WorkbookStatus)
    Write-Host ('  Guardrails applied: {0}' -f $Summary.GuardrailsApplied)
    Write-Host ('  Current week: {0}' -f $Summary.CurrentWeekEnding)
    Write-Host ('  Prior-year week: {0}' -f $Summary.PriorYearWeekEnding)
    Write-Host ('  Rows appended: {0}' -f $Summary.RowsAppended)
    Write-Host ('  Current source: {0}' -f $Summary.CurrentSourceWorkbook)
    Write-Host ('  Prior-year source: {0}' -f $Summary.PriorYearSourceWorkbook)
    Write-Host ('  Target workbook: {0}' -f $Summary.TargetWorkbook)
    if ($Summary.BackupWorkbook) {
        Write-Host ('  Backup workbook: {0}' -f $Summary.BackupWorkbook)
    }
    Write-Host ('  Email PDF: {0}' -f $Summary.EmailComparisonPdf)
    Write-Host ('  Log: {0}' -f $Summary.LogPath)
    Write-Host ''
}

function Invoke-GssWorkbookUpdate {
    param(
        [string]$FolderPath,
        [string]$WorkbookName,
        [switch]$ApplyMode,
        [switch]$ReturnObject
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        $scriptRoot = Split-Path -Parent $PSCommandPath
        $projectRoot = Split-Path -Parent $scriptRoot
        $FolderPath = Split-Path -Parent $projectRoot
    }

    $FolderPath = $FolderPath.Trim('"')
    $FolderPath = (Resolve-Path -LiteralPath $FolderPath).Path
    $mainPath = Resolve-MainWorkbookPath $FolderPath $WorkbookName

    $automationDir = Join-Path $FolderPath '_automation_runs'
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
    $workbookStatus = $null
    $guardrailsApplied = $false

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

        $allSources = @(Get-RollingSources $excel $FolderPath)
        $latestSource = Select-LatestRollingSource $allSources
        if ($null -eq $latestSource) {
            throw "No latest rolling Sorensen source workbook could be selected from $FolderPath."
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

        if (-not $ApplyMode) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($WorkbookName)
            $targetPath = Join-Path $testDir "$base`_TEST_$timestamp.xlsx"
            Copy-Item -LiteralPath $mainPath -Destination $targetPath -Force
        }

        $targetWb = $excel.Workbooks.Open($targetPath, 0, $false)
        if ($ApplyMode -and $targetWb.ReadOnly) {
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

        if ($ApplyMode) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($WorkbookName)
            $backupPath = Join-Path $backupDir "$base`_BACKUP_$timestamp.xlsx"
            $targetWb.SaveCopyAs($backupPath)
        }

        Remove-GssWorkbookGuardrails $targetWb

        if ($sourceWork.Count -eq 0) {
            $status = 'WorkbookAlreadyCurrentGuardrailsRefreshed'
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

        $configuration = Get-GssConfigurationCounts $targetWb
        Update-EmailComparisonData $targetWb $rawWs $latestSource
        Update-GssQaSheet $targetWb $configuration
        Set-GssReadmeContent $targetWb
        Set-GssWorkbookBanners $targetWb
        Set-GssWorkbookPrintLayouts $targetWb $configuration
        $excel.CalculateFullRebuild()
        $workbookStatus = Get-GssWorkbookStatus $targetWb
        Set-GssStatusStyles $targetWb $workbookStatus
        $emailComparisonPdfPath = Export-EmailComparisonPdf $targetWb $latestSource $FolderPath $testDir ([bool]$ApplyMode)
        Set-GssWorkbookGuardrails $targetWb $configuration
        $guardrailsApplied = Test-GssWorkbookGuardrailsApplied $targetWb
        if (-not $guardrailsApplied) {
            throw 'Workbook guardrail verification failed before save.'
        }
        $targetWb.Save()

        $summary = [pscustomobject]@{
            Timestamp = (Get-Date).ToString('s')
            Mode = if ($ApplyMode) { 'ApplyToMainWorkbook' } else { 'CopyTestOnly' }
            Status = $status
            WorkbookStatus = $workbookStatus
            GuardrailsApplied = $guardrailsApplied
            Folder = $FolderPath
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
        $summary | Add-Member -NotePropertyName LogPath -NotePropertyValue $logPath
        Write-RunSummary $summary
        if ($ReturnObject) {
            return $summary
        }
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
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-GssWorkbookUpdate -FolderPath $Folder -WorkbookName $MainWorkbookName -ApplyMode:$Apply -ReturnObject:$OutputObject
}
