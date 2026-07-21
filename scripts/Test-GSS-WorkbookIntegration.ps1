[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$WorkbookPath,
    [string]$BaselinePath,
    [string]$ExportVerificationFolder
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1')

$xlCellTypeFormulas = -4123
$xlCellTypeConstants = 2
$xlErrors = 16

function Assert-GssIntegration {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw "Workbook integration assertion failed: $Message"
    }
}

function Get-GssErrorCellCount {
    param([object]$Worksheet, [int]$CellType)

    $usedRange = $null
    $errorCells = $null
    try {
        $usedRange = $Worksheet.UsedRange
        try {
            $errorCells = $usedRange.SpecialCells($CellType, $xlErrors)
            return [int]$errorCells.Count
        }
        catch {
            return 0
        }
    }
    finally {
        Release-ComObject $errorCells
        Release-ComObject $usedRange
    }
}

function Assert-GssLockedRectangle {
    param(
        [object]$Worksheet,
        [int]$FirstRow,
        [int]$FirstColumn,
        [int]$LastRow,
        [int]$LastColumn
    )

    if ($LastRow -lt $FirstRow -or $LastColumn -lt $FirstColumn) { return }
    $range = $null
    try {
        $range = $Worksheet.Range(
            $Worksheet.Cells.Item($FirstRow, $FirstColumn),
            $Worksheet.Cells.Item($LastRow, $LastColumn)
        )
        Assert-GssIntegration ([bool]$range.Locked) "$($Worksheet.Name)!$($range.Address(0, 0)) must remain locked."
    }
    finally {
        Release-ComObject $range
    }
}

function Test-GssIntentionalCellChange {
    param([string]$SheetName, [int]$Row, [int]$Column)

    switch ($SheetName) {
        'README' { return $Row -le 34 }
        'QA Checks' { return $Row -le 12 -and $Column -le 3 }
        'Quick_Read_WoW' { return $Row -eq 4 -and $Column -le 17 }
        'Quick_Read_YoY' { return $Row -eq 4 -and $Column -le 17 }
        'Exec_Dashboard' { return $Row -eq 2 -and $Column -le 6 }
        'Drivers_Detail' { return $Row -eq 2 -and $Column -le 4 }
        'Email Comparison' {
            return ($Row -eq 2 -and $Column -le 28) -or
                ($Row -in @(3, 23) -and $Column -eq 4) -or
                ($Row -in @(5, 25) -and $Column -ge 2 -and $Column -le 5) -or
                ($Column -eq 5 -and (($Row -ge 6 -and $Row -le 18) -or ($Row -ge 26 -and $Row -le 38)))
        }
        default { return $false }
    }
}

function Compare-GssWorkbookCells {
    param([object]$BaselineWorkbook, [object]$TargetWorkbook)

    for ($sheetIndex = 1; $sheetIndex -le $BaselineWorkbook.Worksheets.Count; $sheetIndex++) {
        $baselineWs = $null
        $targetWs = $null
        $baselineRange = $null
        $targetRange = $null
        try {
            $baselineWs = $BaselineWorkbook.Worksheets.Item($sheetIndex)
            $sheetName = [string]$baselineWs.Name
            $targetWs = $TargetWorkbook.Worksheets.Item($sheetName)
            $baselineRange = $baselineWs.UsedRange
            $targetRange = $targetWs.Range($baselineRange.Address())
            $baselineValues = $baselineRange.FormulaR1C1
            $targetValues = $targetRange.FormulaR1C1
            $rowCount = [int]$baselineRange.Rows.Count
            $columnCount = [int]$baselineRange.Columns.Count

            for ($row = 1; $row -le $rowCount; $row++) {
                for ($column = 1; $column -le $columnCount; $column++) {
                    if (Test-GssIntentionalCellChange $sheetName $row $column) { continue }
                    $before = if ($null -eq $baselineValues[$row, $column]) { '' } else { [string]$baselineValues[$row, $column] }
                    $after = if ($null -eq $targetValues[$row, $column]) { '' } else { [string]$targetValues[$row, $column] }
                    if ($before -cne $after) {
                        throw "Unexpected cell change in $sheetName at row $row, column $column."
                    }
                }
            }
        }
        finally {
            Release-ComObject $targetRange
            Release-ComObject $baselineRange
            Release-ComObject $targetWs
            Release-ComObject $baselineWs
        }
    }
}

$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
if ($BaselinePath) {
    $BaselinePath = (Resolve-Path -LiteralPath $BaselinePath).Path
}
if ($ExportVerificationFolder) {
    New-Item -ItemType Directory -Path $ExportVerificationFolder -Force | Out-Null
    $ExportVerificationFolder = (Resolve-Path -LiteralPath $ExportVerificationFolder).Path
}

$excel = $null
$workbook = $null
$baselineWorkbook = $null
$verificationPdfs = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
    $excel.CalculateFullRebuild()

    $qaWs = $null
    $statusCell = $null
    try {
        $qaWs = $workbook.Worksheets.Item('QA Checks')
        $statusCell = $qaWs.Range('B2')
        Assert-GssIntegration ($qaWs.Index -eq 2) 'QA Checks must be immediately after README.'
        Assert-GssIntegration ($qaWs.Visible -eq $xlSheetVisible) 'QA Checks must be visible.'
        Assert-GssIntegration (([string]$statusCell.Value2).Trim().ToUpperInvariant() -eq 'READY') 'QA Checks B2 must report READY.'
    }
    finally {
        Release-ComObject $statusCell
        Release-ComObject $qaWs
    }

    Assert-GssIntegration ([bool]$workbook.ProtectStructure) 'Workbook structure must be protected.'
    $formulaErrorCount = 0
    $constantErrorCount = 0
    for ($sheetIndex = 1; $sheetIndex -le $workbook.Worksheets.Count; $sheetIndex++) {
        $worksheet = $null
        $usedRange = $null
        $outsideUsedRangeCell = $null
        try {
            $worksheet = $workbook.Worksheets.Item($sheetIndex)
            Assert-GssIntegration ([bool]$worksheet.ProtectContents) "Worksheet $($worksheet.Name) must be protected."
            $formulaErrorCount += Get-GssErrorCellCount $worksheet $xlCellTypeFormulas
            $constantErrorCount += Get-GssErrorCellCount $worksheet $xlCellTypeConstants

            $outsideUsedRangeCell = $worksheet.Cells.Item($worksheet.Rows.Count, $worksheet.Columns.Count)
            Assert-GssIntegration ([bool]$outsideUsedRangeCell.Locked) "$($worksheet.Name) cells outside UsedRange must remain locked."

            $usedRange = $worksheet.UsedRange
            if ($worksheet.Name -eq 'Exec_Dashboard') {
                $lastRow = $usedRange.Row + $usedRange.Rows.Count - 1
                $lastColumn = $usedRange.Column + $usedRange.Columns.Count - 1
                $entitySelector = $null
                $metricSelector = $null
                try {
                    $entitySelector = $worksheet.Range('C3')
                    $metricSelector = $worksheet.Range('F3')
                    Assert-GssIntegration (-not [bool]$entitySelector.Locked) 'Exec_Dashboard!C3 must be unlocked.'
                    Assert-GssIntegration (-not [bool]$metricSelector.Locked) 'Exec_Dashboard!F3 must be unlocked.'
                }
                finally {
                    Release-ComObject $metricSelector
                    Release-ComObject $entitySelector
                }
                Assert-GssLockedRectangle $worksheet 1 1 2 $lastColumn
                Assert-GssLockedRectangle $worksheet 3 1 3 2
                Assert-GssLockedRectangle $worksheet 3 4 3 5
                Assert-GssLockedRectangle $worksheet 3 7 3 $lastColumn
                Assert-GssLockedRectangle $worksheet 4 1 $lastRow $lastColumn
            }
            else {
                Assert-GssIntegration ([bool]$usedRange.Locked) "$($worksheet.Name) used range must remain locked."
            }
        }
        finally {
            Release-ComObject $outsideUsedRangeCell
            Release-ComObject $usedRange
            Release-ComObject $worksheet
        }
    }
    Assert-GssIntegration ($formulaErrorCount -eq 0) "Calculated workbook contains $formulaErrorCount formula error cell(s)."
    Assert-GssIntegration ($constantErrorCount -eq 0) "Workbook contains $constantErrorCount constant error cell(s)."

    foreach ($hiddenSheetName in @('Entities', 'Metric_Catalog', 'Driver_Weights')) {
        $hiddenWs = $null
        try {
            $hiddenWs = $workbook.Worksheets.Item($hiddenSheetName)
            Assert-GssIntegration ($hiddenWs.Visible -ne $xlSheetVisible) "$hiddenSheetName must remain hidden."
        }
        finally {
            Release-ComObject $hiddenWs
        }
    }

    $configuration = Get-GssConfigurationCounts $workbook
    $execWs = $null
    try {
        $execWs = $workbook.Worksheets.Item('Exec_Dashboard')
        foreach ($definition in @(
            [pscustomobject]@{ Address = 'C3'; Formula = "=Entities!`$A`$2:`$A`$$($configuration.EntityLastRow)" },
            [pscustomobject]@{ Address = 'F3'; Formula = "=Metric_Catalog!`$A`$2:`$A`$$($configuration.MetricLastRow)" }
        )) {
            $cell = $null
            $validation = $null
            try {
                $cell = $execWs.Range($definition.Address)
                $validation = $cell.Validation
                Assert-GssIntegration ($validation.Type -eq $xlValidateList) "$($definition.Address) must use list validation."
                Assert-GssIntegration ([bool]$validation.ShowError) "$($definition.Address) must show validation errors."
                Assert-GssIntegration (-not [bool]$validation.IgnoreBlank) "$($definition.Address) must reject blank selections."
                Assert-GssIntegration (([string]$validation.Formula1) -eq $definition.Formula) "$($definition.Address) validation range is incorrect."
            }
            finally {
                Release-ComObject $validation
                Release-ComObject $cell
            }
        }
    }
    finally {
        Release-ComObject $execWs
    }

    foreach ($layout in @(Get-GssWorkbookLayoutPlan $configuration.ActiveEntityCount $configuration.MetricCount)) {
        $worksheet = $null
        $pageSetup = $null
        try {
            $worksheet = $workbook.Worksheets.Item($layout.Sheet)
            $pageSetup = $worksheet.PageSetup
            Assert-GssIntegration ($pageSetup.PrintArea -eq $layout.PrintArea) "$($layout.Sheet) print area is incorrect."
            Assert-GssIntegration ($pageSetup.Orientation -eq $layout.Orientation) "$($layout.Sheet) orientation is incorrect."
            Assert-GssIntegration (-not [bool]$pageSetup.Zoom) "$($layout.Sheet) must use fit-to-pages scaling."
            Assert-GssIntegration ($pageSetup.FitToPagesWide -eq $layout.FitWide) "$($layout.Sheet) FitToPagesWide is incorrect."
            Assert-GssIntegration ($pageSetup.FitToPagesTall -eq $layout.FitTall) "$($layout.Sheet) FitToPagesTall is incorrect."
        }
        finally {
            Release-ComObject $pageSetup
            Release-ComObject $worksheet
        }
    }

    foreach ($banner in @(
        [pscustomobject]@{ Sheet = 'README'; Address = 'A3' },
        [pscustomobject]@{ Sheet = 'Quick_Read_WoW'; Address = 'A4' },
        [pscustomobject]@{ Sheet = 'Quick_Read_YoY'; Address = 'A4' },
        [pscustomobject]@{ Sheet = 'Exec_Dashboard'; Address = 'A2' },
        [pscustomobject]@{ Sheet = 'Drivers_Detail'; Address = 'A2' },
        [pscustomobject]@{ Sheet = 'Email Comparison'; Address = 'A2' }
    )) {
        $worksheet = $null
        $cell = $null
        try {
            $worksheet = $workbook.Worksheets.Item($banner.Sheet)
            $cell = $worksheet.Range($banner.Address)
            Assert-GssIntegration (([string]$cell.Text) -like '*READY*') "$($banner.Sheet) status banner must display READY."
        }
        finally {
            Release-ComObject $cell
            Release-ComObject $worksheet
        }
    }

    if ($BaselinePath) {
        $baselineWorkbook = $excel.Workbooks.Open($BaselinePath, 0, $true)
        Compare-GssWorkbookCells $baselineWorkbook $workbook
    }

    if ($ExportVerificationFolder) {
        foreach ($sheetName in @('README', 'QA Checks', 'Quick_Read_WoW', 'Quick_Read_YoY', 'Exec_Dashboard', 'Drivers_Detail', 'Email Comparison')) {
            $worksheet = $null
            try {
                $worksheet = $workbook.Worksheets.Item($sheetName)
                $safeName = $sheetName -replace '[^A-Za-z0-9_-]', '_'
                $pdfPath = Join-Path $ExportVerificationFolder "$safeName.pdf"
                $worksheet.ExportAsFixedFormat($xlTypePDF, $pdfPath)
                $verificationPdfs += $pdfPath
            }
            finally {
                Release-ComObject $worksheet
            }
        }
    }

    Write-Host 'GSS Excel workbook integration verification passed.'
    Write-Host "  Workbook: $WorkbookPath"
    Write-Host '  Workbook QA: READY'
    Write-Host '  Formula errors: 0'
    Write-Host '  Constant errors: 0'
    if ($BaselinePath) { Write-Host "  Preserved baseline cells: $BaselinePath" }
    foreach ($pdfPath in $verificationPdfs) { Write-Host "  Verification PDF: $pdfPath" }
}
finally {
    if ($baselineWorkbook) { $baselineWorkbook.Close($false) }
    if ($workbook) { $workbook.Close($false) }
    Release-ComObject $baselineWorkbook
    Release-ComObject $workbook
    if ($excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
