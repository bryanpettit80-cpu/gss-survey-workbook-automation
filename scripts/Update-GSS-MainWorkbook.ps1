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

function Convert-ExcelDate {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.Date }
    if ($Value -is [double] -or $Value -is [int]) { return ([datetime]::FromOADate([double]$Value)).Date }
    return ([datetime]::Parse([string]$Value)).Date
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

    $candidates = Get-ChildItem -LiteralPath $FolderPath -File -Filter 'Sorensen FW*.xlsx' |
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

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $scriptRoot = Split-Path -Parent $PSCommandPath
    $projectRoot = Split-Path -Parent $scriptRoot
    $Folder = Split-Path -Parent $projectRoot
}

$Folder = (Resolve-Path -LiteralPath $Folder).Path
$mainPath = Join-Path $Folder $MainWorkbookName
if (-not (Test-Path -LiteralPath $mainPath)) {
    throw "Main workbook not found: $mainPath"
}

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

    if ($sourceWork.Count -eq 0) {
        $status = 'SkippedAlreadyPresent'
    }
    else {
        if ($Apply) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($MainWorkbookName)
            $backupPath = Join-Path $backupDir "$base`_BACKUP_$timestamp.xlsx"
            $targetWb.SaveCopyAs($backupPath)
        }

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

        $excel.CalculateFullRebuild()
        $targetWb.Save()
    }

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
