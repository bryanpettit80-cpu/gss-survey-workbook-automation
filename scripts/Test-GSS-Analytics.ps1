[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Analyze-GSS-Run.ps1')

function Assert-Equal {
    param([object]$Actual, [object]$Expected, [string]$Name)
    if ($Actual -ne $Expected) {
        throw "Assertion failed for $Name. Expected '$Expected'; actual '$Actual'."
    }
}

function Assert-True {
    param([bool]$Value, [string]$Name)
    if (-not $Value) {
        throw "Assertion failed for $Name."
    }
}

Assert-Equal (Get-DirectionAdjustedChange 90 80 $false) 10 'Higher-is-better positive improvement'
Assert-Equal (Get-DirectionAdjustedChange 5 8 $true) 3 'Lower-is-better positive improvement'
Assert-Equal (Test-GssYoyPairing ([datetime]'2026-07-05') ([datetime]'2025-07-06')) $true 'Valid YoY pairing'
Assert-Equal (Test-GssYoyPairing ([datetime]'2026-07-05') ([datetime]'2025-07-05')) $false 'Invalid YoY pairing'
Assert-Equal (Format-GssMovementNumber 2.25 1) '+2.3' 'Positive movement formatting'
Assert-Equal (Format-GssMovementNumber -2.25 1) '-2.3' 'Negative movement formatting'

$metrics = @(
    [pscustomobject]@{ DisplayName = 'Overall Experience'; RawKey = 'Overall Experience'; LowerIsBetter = $false; IncludeInMovers = $true; Category = 'Overall'; Owner = 'GM' },
    [pscustomobject]@{ DisplayName = 'Overall Dissat'; RawKey = 'Overall Dissat'; LowerIsBetter = $true; IncludeInMovers = $true; Category = 'Overall'; Owner = 'GM' },
    [pscustomobject]@{ DisplayName = 'Count'; RawKey = 'Count'; LowerIsBetter = $false; IncludeInMovers = $false; Category = 'Sampling'; Owner = 'Ops' }
)

function New-TestRawRow {
    param(
        [string]$EntityKey,
        [datetime]$Week,
        [double]$Overall,
        [double]$Dissat,
        [int]$Count
    )

    $values = [pscustomobject]@{
        'Overall Experience' = $Overall
        'Overall Dissat' = $Dissat
        'Count' = $Count
    }
    [pscustomobject]@{
        EntityKey = $EntityKey
        Week = $Week
        RowKey = Get-GssEntityWeekKey $EntityKey $Week
        Values = $values
    }
}

$current = [datetime]'2026-07-05'
$priorYear = [datetime]'2025-07-06'
$rawRows = @(
    (New-TestRawRow 'Sorensen|9354 Richmond' $current 70 5 150),
    (New-TestRawRow 'Sorensen|9354 Richmond' $current.AddDays(-7) 75 8 150),
    (New-TestRawRow 'Sorensen|9354 Richmond' $priorYear 80 9 140),
    (New-TestRawRow 'Sorensen|9354 Richmond' $current.AddDays(-14) 73 7 130),
    (New-TestRawRow 'Sorensen|9354 Richmond' $current.AddDays(-21) 74 6 120),
    (New-TestRawRow 'Sorensen|9355 Virginia Beach' $current 82 4 150),
    (New-TestRawRow 'Sorensen|9355 Virginia Beach' $current.AddDays(-7) 80 5 150),
    (New-TestRawRow 'Sorensen|9355 Virginia Beach' $priorYear 78 7 140),
    (New-TestRawRow 'Sorensen|(TOTAL)' $current 76 5 300),
    (New-TestRawRow 'All Franchisees|(TOTAL)' $current 68 7 1000)
)

$detail = @(New-GssMetricDetail $rawRows $metrics $current $priorYear 4)
$richmondOverall = $detail | Where-Object { $_.Entity -eq '9354 Richmond' -and $_.Metric -eq 'Overall Experience' } | Select-Object -First 1
$richmondDissat = $detail | Where-Object { $_.Entity -eq '9354 Richmond' -and $_.Metric -eq 'Overall Dissat' } | Select-Object -First 1
$vbOverall = $detail | Where-Object { $_.Entity -eq '9355 Virginia Beach' -and $_.Metric -eq 'Overall Experience' } | Select-Object -First 1

Assert-Equal $richmondOverall.WoWImprovement -5 'Metric detail WoW decline'
Assert-Equal $richmondOverall.YoYImprovement -10 'Metric detail YoY decline'
Assert-Equal $richmondOverall.WorstMovementLabel 'YoY' 'Weakest comparison label'
Assert-Equal $richmondOverall.BestMovementLabel 'WoW' 'Strongest comparison label when both are negative'
Assert-Equal $richmondDissat.WoWImprovement 3 'Lower-is-better metric detail'
Assert-Equal $vbOverall.BestMovement 4 'Strength ranking candidate'
Assert-Equal $vbOverall.BestMovementLabel 'YoY' 'Strength improvement label'

$attention = Select-GssAttentionItems $detail 1
Assert-Equal $attention[0].Entity '9354 Richmond' 'Attention ranking entity'
Assert-Equal $attention[0].Metric 'Overall Experience' 'Attention ranking metric'

$strength = Select-GssStrengthItems $detail 1
Assert-Equal $strength[0].Entity '9355 Virginia Beach' 'Strength ranking entity'

$headers = @('Overall Experience', 'Overall Dissat', 'Count')
$runLog = [pscustomobject]@{
    Mode = 'ApplyToMainWorkbook'
    EmailComparisonPdf = $null
    BackupWorkbook = $null
    CurrentSourceWorkbook = 'C:\GSS Surveys\02 Weekly Rolling Source Workbooks\current.xlsx'
    PriorYearSourceWorkbook = 'C:\GSS Surveys\02 Weekly Rolling Source Workbooks\prior.xlsx'
    RowsAppended = 8
}
$qa = Get-GssRunQa $runLog 'C:\GSS Surveys' 'C:\missing.xlsx' $rawRows $metrics $headers $detail $current $priorYear
Assert-Equal $qa.Status 'Blocked' 'QA blocker status'
Assert-True (@($qa.Blockers | Where-Object { $_ -like 'Main workbook is missing*' }).Count -eq 1) 'QA missing workbook blocker'
Assert-True (@($qa.Warnings | Where-Object { $_ -like 'QA Checks sheet is missing*' }).Count -eq 1) 'Legacy missing QA sheet warning'

$attentionQa = Get-GssRunQa $runLog 'C:\GSS Surveys' 'C:\missing.xlsx' $rawRows $metrics $headers $detail $current $priorYear $true 'ATTENTION'
Assert-True (@($attentionQa.Blockers | Where-Object { $_ -like 'Workbook QA Checks reports ATTENTION*' }).Count -eq 1) 'Workbook ATTENTION is a blocker'

$readyQa = Get-GssRunQa $runLog 'C:\GSS Surveys' 'C:\missing.xlsx' $rawRows $metrics $headers $detail $current $priorYear $true 'READY'
Assert-True (@($readyQa.Blockers | Where-Object { $_ -like 'Workbook QA*' }).Count -eq 0) 'Workbook READY adds no QA blocker'
Assert-True (@($readyQa.Warnings | Where-Object { $_ -like 'QA Checks*' }).Count -eq 0) 'Workbook READY adds no QA warning'

$invalidQa = Get-GssRunQa $runLog 'C:\GSS Surveys' 'C:\missing.xlsx' $rawRows $metrics $headers $detail $current $priorYear $true '#VALUE!'
Assert-True (@($invalidQa.Blockers | Where-Object { $_ -like 'Workbook QA Checks returned an invalid status*' }).Count -eq 1) 'Invalid workbook QA is a blocker'

$pathRunLog = [pscustomobject]@{ TargetWorkbook = 'C:\GSS Surveys\_automation_runs\test-output\guarded-copy.xlsx' }
Assert-Equal (Resolve-GssAnalysisWorkbookPath $pathRunLog 'C:\GSS Surveys' 'main.xlsx') 'C:\GSS Surveys\_automation_runs\test-output\guarded-copy.xlsx' 'Analyzer uses logged copy-test workbook'

Write-Host 'GSS analytics logic tests passed.'
