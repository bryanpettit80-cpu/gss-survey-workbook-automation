[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TemplatePath,
    [Parameter(Mandatory)][string]$EvidenceRoot,
    [switch]$FunctionsOnly
)

$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\Update-GSS-MainWorkbook.ps1')

function Assert-Scenario {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Reporting scenario failed: $Message" }
}

function Get-ScenarioRow {
    param([datetime]$Week, [double]$Overall, [double]$Dissat)
    @(
        [pscustomobject]@{ SourceFile = 'synthetic.xlsx'; Week = $Week; Restaurant = $null; Ownership = 'All Franchisees'; Metrics = @(1000,68,80,85,76,78,92,66,7,6,7,8,9) },
        [pscustomobject]@{ SourceFile = 'synthetic.xlsx'; Week = $Week; Restaurant = $null; Ownership = 'Sorensen'; Metrics = @(300,76,80,85,76,78,92,66,5,6,7,8,9) },
        [pscustomobject]@{ SourceFile = 'synthetic.xlsx'; Week = $Week; Restaurant = '9354 Richmond'; Ownership = 'Sorensen'; Metrics = @(150,$Overall,80,85,76,78,92,66,$Dissat,6,7,8,9) },
        [pscustomobject]@{ SourceFile = 'synthetic.xlsx'; Week = $Week; Restaurant = '9355 Virginia Beach'; Ownership = 'Sorensen'; Metrics = @(150,82,80,85,76,78,92,66,4,6,7,8,9) }
    )
}

function Write-ScenarioSource {
    param([object]$Excel, [string]$Path, [datetime]$Week, [double]$Overall, [double]$Dissat)
    $book = $null
    $sheet = $null
    try {
        $book = $Excel.Workbooks.Add()
        $sheet = $book.Worksheets.Item(1)
        Set-CellValue $sheet 1 1 ('Date Range: {0} to {1}' -f $Week.AddDays(-90).ToString('M/d/yyyy'), $Week.ToString('M/d/yyyy'))
        $headers = @(Get-GssSourceHeader)
        for ($col = 0; $col -lt $headers.Count; $col++) { Set-CellValue $sheet 3 ($col + 1) $headers[$col] }
        $row = 4
        foreach ($values in @(Get-ScenarioRow $Week $Overall $Dissat)) {
            Set-CellValue $sheet $row 1 $values.Restaurant
            Set-CellValue $sheet $row 2 $values.Ownership
            for ($col = 0; $col -lt 13; $col++) { Set-CellValue $sheet $row ($col + 3) $values.Metrics[$col] }
            $row++
        }
        $book.SaveAs($Path, 51)
    }
    finally {
        if ($book) { $book.Close($false) }
        Release-ComObject $sheet
        Release-ComObject $book
    }
}

function Get-ScenarioOutput {
    param([string]$Path)
    $excel = $null; $book = $null; $raw = $null; $email = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
        $book = $excel.Workbooks.Open($Path, 0, $true)
        $excel.CalculateFullRebuild()
        $raw = $book.Worksheets.Item('Raw_Data')
        $email = $book.Worksheets.Item('Email Comparison')
        $metrics = @{}
        for ($row = 26; $row -le 38; $row++) {
            $metrics[[string]$email.Cells.Item($row, 1).Value2] = [pscustomobject]@{
                Current = $email.Cells.Item($row, 2).Value2
                PreviousChange = $email.Cells.Item($row, 3).Value2
                PriorYearChange = $email.Cells.Item($row, 4).Value2
            }
        }
        [pscustomobject]@{ Metrics = $metrics; RawRows = @(Read-GssRawDataRow $raw) }
    }
    finally {
        if ($book) { $book.Close($false) }
        Release-ComObject $raw; Release-ComObject $email; Release-ComObject $book
        if ($excel) { $excel.Quit() }; Release-ComObject $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Get-ScenarioModelFingerprint {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        function Read-ScenarioXmlEntry {
            param([string]$EntryName)
            $reader = New-Object IO.StreamReader($archive.GetEntry($EntryName).Open())
            try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        $workbookXml = Read-ScenarioXmlEntry 'xl/workbook.xml'
        $relationsXml = Read-ScenarioXmlEntry 'xl/_rels/workbook.xml.rels'
        $material = @()
        foreach ($sheet in $workbookXml.workbook.sheets.sheet) {
            if ($sheet.name -notin @('Exec_Dashboard', 'Drivers_Detail', 'Weekly Trend')) { continue }
            $relationshipId = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $target = [string]($relationsXml.Relationships.Relationship | Where-Object Id -eq $relationshipId).Target
            $entry = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { 'xl/' + $target }
            $sheetXml = Read-ScenarioXmlEntry $entry
            foreach ($cell in $sheetXml.worksheet.sheetData.row.c) {
                if ($cell.f -and [string]$cell.f -match 'Raw_Data') { $material += '{0}|{1}|{2}' -f $sheet.name, $cell.r, [string]$cell.f }
            }
        }
        Assert-Scenario ($material.Count -gt 0) 'model reference evidence must contain formulas'
        return Get-GssTextSha256 ($material -join "`n")
    }
    finally { $archive.Dispose() }
}

if ($FunctionsOnly) { return }
$template = (Resolve-Path -LiteralPath $TemplatePath).Path
$evidence = [IO.Path]::GetFullPath($EvidenceRoot)
$folder = Join-Path $evidence ('synthetic-reporting-' + [guid]::NewGuid().ToString('N'))
$mainDirectory = Join-Path $folder '01 Main Workbook'
$sourceDirectory = Join-Path $folder '02 Rolling Reports'
New-Item -ItemType Directory -Path $mainDirectory, $sourceDirectory -Force | Out-Null
$main = Join-Path $mainDirectory 'Survey History.xlsx'
Copy-Item -LiteralPath $template -Destination $main
$excel = $null; $book = $null; $raw = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
    $book = $excel.Workbooks.Open($main, 0, $false)
    Remove-GssWorkbookGuardrails $book
    $raw = $book.Worksheets.Item('Raw_Data')
    $raw.Range('A2:S5096').ClearContents() | Out-Null
    $null = Add-RawDataRows $raw @(Get-ScenarioRow ([datetime]'2026-09-27') 77 10)
    $book.Save(); $book.Close($false)
    Release-ComObject $raw; $raw = $null; Release-ComObject $book; $book = $null
    Write-ScenarioSource $excel (Join-Path $sourceDirectory 'Latest guest survey.xlsx') ([datetime]'2026-10-11') 70 5
    Write-ScenarioSource $excel (Join-Path $sourceDirectory 'Missed last week.xlsx') ([datetime]'2026-10-04') 75 8
    Write-ScenarioSource $excel (Join-Path $sourceDirectory 'Prior year.xlsx') ([datetime]'2025-10-12') 80 9
    Write-ScenarioSource $excel (Join-Path $sourceDirectory 'Catch-up prior year.xlsx') ([datetime]'2025-10-05') 79 9
}
finally {
    if ($book) { $book.Close($false) }
    Release-ComObject $raw; Release-ComObject $book
    if ($excel) { $excel.Quit() }; Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Output "Synthetic fixture: $folder"
$modelBefore = Get-ScenarioModelFingerprint $main
$prepared = Invoke-GssWorkbookUpdate -FolderPath $folder -WorkbookName 'GSS Score Trends - Main.xlsx' -ReportingWeek '2026-10-11' -ReturnObject
Assert-Scenario ($prepared.RowsAppended -eq 16) 'two missed current periods and their prior-year periods must append exactly 16 rows'
Assert-Scenario ($prepared.SourceEvidence.Count -eq 4) 'all catch-up source evidence must be retained'
Assert-Scenario ($prepared.PostSaveValidation.Status -eq 'Ready') 'persisted workbook must pass validation'
$first = Get-ScenarioOutput $prepared.TargetWorkbook
Assert-Scenario ((Get-ScenarioModelFingerprint $prepared.TargetWorkbook) -eq $modelBefore) 'catch-up must preserve dashboard and detail Raw_Data formula references'
Assert-Scenario ($first.Metrics['Overall Experience'].Current -eq 70) 'overall must equal 70'
Assert-Scenario ($first.Metrics['Overall Experience'].PreviousChange -eq -5) '70 minus 75 must equal -5 percentage points'
Assert-Scenario ($first.Metrics['Overall Experience'].PriorYearChange -eq -10) '70 minus 80 must equal -10 percentage points'
Assert-Scenario ($first.Metrics['Overall Dissat'].Current -eq 5) 'dissat must equal 5'
Assert-Scenario ($first.Metrics['Overall Dissat'].PreviousChange -eq 3) '8 to 5 dissat must improve by 3 percentage points'
Assert-Scenario ($first.Metrics['Overall Dissat'].PriorYearChange -eq 4) '9 to 5 dissat must improve by 4 percentage points'
Assert-Scenario ($first.Metrics['Count'].Current -eq 150) 'response count must remain 150'
Assert-Scenario (@($first.RawRows | Group-Object { $_.Week.ToString('yyyy-MM-dd') } | Where-Object Count -ne 4).Count -eq 0) 'every imported week must have exactly four entities'
Write-Output 'Catch-up and independent workbook arithmetic: PASS'

# This promotion is confined to the newly created synthetic fixture directory.
$promotion = Invoke-GssPreparedRunPromotion -PreparedRun $prepared -ExpectedFingerprint $prepared.RunFingerprint -FolderPath $folder -MainWorkbookPath $main -BackupDirectory (Join-Path $folder '_automation_runs\backups') -LogDirectory (Join-Path $folder '_automation_runs\logs')
Assert-Scenario ($promotion.TransactionStatus -eq 'Committed') 'synthetic promotion must commit'
$currentSource = Join-Path $sourceDirectory 'Latest guest survey.xlsx'
$renamedSource = Join-Path $sourceDirectory 'Renamed guest survey.xlsx'
Move-Item -LiteralPath $currentSource -Destination $renamedSource
Copy-Item -LiteralPath $renamedSource -Destination (Join-Path $sourceDirectory 'Duplicate same contents.xlsx')
$unchanged = Invoke-GssWorkbookUpdate -FolderPath $folder -ReportingWeek '2026-10-11' -ReturnObject
Assert-Scenario ($unchanged.NoBusinessChanges -and $unchanged.RowsAppended -eq 0 -and $unchanged.RowsCorrected -eq 0) 'renamed identical duplicates must be unchanged'
Write-Output 'Identical renamed and duplicate inputs: PASS'

$correctedSource = Join-Path $sourceDirectory 'Corrected guest survey.xlsx'
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    Write-ScenarioSource $excel $correctedSource ([datetime]'2026-10-11') 72 4
}
finally { if ($excel) { $excel.Quit() }; Release-ComObject $excel; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
$priorMainHash = Get-GssSha256 $main
$choiceRequired = $false
try { $null = Invoke-GssWorkbookUpdate -FolderPath $folder -ReportingWeek '2026-10-11' }
catch { $choiceRequired = $_.Exception.Data['GssChoice'] -eq 'CorrectionSource'; if (-not $choiceRequired) { throw } }
Assert-Scenario $choiceRequired 'conflicting same-period inputs must require a choice'
Assert-Scenario ((Get-GssSha256 $main) -eq $priorMainHash) 'choice request must preserve live fixture workbook'
$correction = Invoke-GssWorkbookUpdate -FolderPath $folder -ReportingWeek '2026-10-11' -CorrectionSourcePaths @($correctedSource) -ReturnObject
Assert-Scenario ($correction.RowsCorrected -eq 4 -and $correction.RowsAppended -eq 0) 'chosen correction must replace exactly four rows'
$second = Get-ScenarioOutput $correction.TargetWorkbook
Assert-Scenario ((Get-ScenarioModelFingerprint $correction.TargetWorkbook) -eq $modelBefore) 'correction must preserve dashboard and detail Raw_Data formula references'
Assert-Scenario ($second.Metrics['Overall Experience'].Current -eq 72 -and $second.Metrics['Overall Experience'].PreviousChange -eq -3 -and $second.Metrics['Overall Experience'].PriorYearChange -eq -8) 'corrected overall arithmetic must equal 72/-3/-8'
Assert-Scenario ($second.Metrics['Overall Dissat'].Current -eq 4 -and $second.Metrics['Overall Dissat'].PreviousChange -eq 4 -and $second.Metrics['Overall Dissat'].PriorYearChange -eq 5) 'corrected dissat arithmetic must equal 4/+4/+5'
Assert-Scenario (@($second.RawRows | Where-Object { $_.Week.Date -eq [datetime]'2026-10-11' }).Count -eq 4) 'correction must not duplicate rows'
Write-Output 'Explicit correction and independent workbook arithmetic: PASS'
$correctedPromotion = Invoke-GssPreparedRunPromotion -PreparedRun $correction -ExpectedFingerprint $correction.RunFingerprint -FolderPath $folder -MainWorkbookPath $main -BackupDirectory (Join-Path $folder '_automation_runs\backups') -LogDirectory (Join-Path $folder '_automation_runs\logs')
$acceptedRerun = Invoke-GssWorkbookUpdate -FolderPath $folder -ReportingWeek '2026-10-11' -ReturnObject
Assert-Scenario ($acceptedRerun.NoBusinessChanges -and $acceptedRerun.RowsCorrected -eq 0) 'accepted correction must not require another choice on an identical rerun'
Assert-Scenario ($acceptedRerun.CurrentSourceSemanticSha256 -eq $correction.CurrentSourceSemanticSha256) 'accepted correction must retain the approved content'
Write-Output 'Committed correction choice reused without another prompt: PASS'
$receipt = [pscustomobject]@{
    Status = 'Passed'; FixtureFolder = $folder; CatchUpRun = $prepared.LogPath; SyntheticPromotion = $promotion.LogPath
    IdenticalRun = $unchanged.LogPath; CorrectionRun = $correction.LogPath
    CorrectedPromotion = $correctedPromotion.LogPath; AcceptedRerun = $acceptedRerun.LogPath; ModelFormulaReferencesUnchanged = $true
    FirstPdf = $prepared.EmailComparisonPdf; CorrectedPdf = $correction.EmailComparisonPdf
    InitialExpected = 'Overall 70/-5/-10; Dissat 5/+3/+4; Count 150'
    CorrectedExpected = 'Overall 72/-3/-8; Dissat 4/+4/+5; Count 150'
    CheckedUtc = [datetime]::UtcNow.ToString('o')
}
$receiptPath = Join-Path $folder 'scenario-results.json'
Write-GssAtomicJson -Path $receiptPath -InputObject $receipt
$receipt | ConvertTo-Json -Depth 6
