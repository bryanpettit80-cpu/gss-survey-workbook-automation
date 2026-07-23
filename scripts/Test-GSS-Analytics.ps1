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
Assert-Equal $script:GssAnalysisPolicyVersion 'gss-analysis-policy/v2' 'Versioned analysis policy'
Assert-Equal ([double]$script:GssAnalysisPolicy.thresholds.previous_window.candidate_points) 1 'Previous-window candidate threshold remains one point'
Assert-Equal ([double]$script:GssAnalysisPolicy.thresholds.previous_window.action_points) 2 'Previous-window action threshold remains two points'
Assert-Equal ([double]$script:GssAnalysisPolicy.thresholds.prior_year.candidate_points) 2 'Prior-year candidate threshold remains two points'
Assert-Equal ([double]$script:GssAnalysisPolicy.thresholds.prior_year.action_points) 5 'Prior-year action threshold remains five points'
Assert-Equal (Get-GssConfidenceTier 100) 'High' 'High confidence boundary'
Assert-Equal (Get-GssConfidenceTier 99) 'Developing' 'Developing confidence upper boundary'
Assert-Equal (Get-GssConfidenceTier 50) 'Developing' 'Developing confidence lower boundary'
Assert-Equal (Get-GssConfidenceTier 49) 'Low' 'Low confidence upper boundary'
Assert-Equal (Get-GssConfidenceTier 1) 'Low' 'Low confidence lower boundary'
Assert-Equal (Get-GssConfidenceTier 0) 'Not scored' 'Zero responses are not scored'
Assert-Equal (Get-GssConfidenceTier $null) 'Not scored' 'Missing responses are not scored'

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
Assert-Equal $richmondOverall.WorstMovementLabel 'PriorYearRollingWindow' 'Weakest comparison label'
Assert-Equal $richmondOverall.BestMovementLabel 'PreviousRollingWindow' 'Strongest comparison label when both are negative'
Assert-Equal $richmondOverall.ConfidenceTier 'High' 'Metric confidence tier'
Assert-Equal $richmondOverall.Level.RollingWeeks 13 'Structured level window'
Assert-Equal $richmondOverall.Movement.AdjacentWindowOverlapWeeks 12 'Structured movement adjacent overlap'
Assert-Equal $richmondOverall.Benchmark.VsAllFranchisees $richmondOverall.VsAllFranchisees 'Structured benchmark value'
Assert-Equal $richmondDissat.WoWImprovement 3 'Lower-is-better metric detail'
Assert-Equal $vbOverall.BestMovement 4 'Strength ranking candidate'
Assert-Equal $vbOverall.BestMovementLabel 'PriorYearRollingWindow' 'Strength improvement label'

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

$capacityRunLog = $runLog | Select-Object *
$capacityRunLog | Add-Member -NotePropertyName WorstCaseWeeklyLoadsRemaining -NotePropertyValue 103
$capacityQa = Get-GssRunQa $capacityRunLog 'C:\GSS Surveys' 'C:\missing.xlsx' $rawRows $metrics $headers $detail $current $priorYear $true 'READY'
Assert-True (@($capacityQa.Warnings | Where-Object { $_ -like 'Workbook capacity is below the redesign trigger*' }).Count -eq 1) 'Capacity below 104 weekly loads triggers redesign warning'

$healthyCapacityRunLog = $runLog | Select-Object *
$healthyCapacityRunLog | Add-Member -NotePropertyName WorstCaseWeeklyLoadsRemaining -NotePropertyValue 104
$healthyCapacityQa = Get-GssRunQa $healthyCapacityRunLog 'C:\GSS Surveys' 'C:\missing.xlsx' $rawRows $metrics $headers $detail $current $priorYear $true 'READY'
Assert-True (@($healthyCapacityQa.Warnings | Where-Object { $_ -like 'Workbook capacity is below the redesign trigger*' }).Count -eq 0) 'Capacity at 104 weekly loads does not trigger redesign warning'

$pathRunLog = [pscustomobject]@{ TargetWorkbook = 'C:\GSS Surveys\_automation_runs\test-output\guarded-copy.xlsx' }
Assert-Equal (Resolve-GssAnalysisWorkbookPath $pathRunLog 'C:\GSS Surveys' 'main.xlsx') 'C:\GSS Surveys\_automation_runs\test-output\guarded-copy.xlsx' 'Analyzer uses logged copy-test workbook'

function Get-BoundaryMetricDetail {
    param(
        [double]$CurrentValue,
        [double]$PreviousValue,
        [double]$PriorYearValue,
        [bool]$LowerIsBetter = $false
    )
    $metricSet = @([pscustomobject]@{ DisplayName = 'Service'; RawKey = 'Service'; LowerIsBetter = $LowerIsBetter; IncludeInMovers = $true; Category = 'Service'; Owner = 'FOH' })
    $rows = @(
        [pscustomobject]@{ EntityKey = 'Sorensen|9354 Richmond'; Week = $current; RowKey = 'current'; Values = [pscustomobject]@{ Service = $CurrentValue; Count = 120 } },
        [pscustomobject]@{ EntityKey = 'Sorensen|9354 Richmond'; Week = $current.AddDays(-7); RowKey = 'previous'; Values = [pscustomobject]@{ Service = $PreviousValue; Count = 120 } },
        [pscustomobject]@{ EntityKey = 'Sorensen|9354 Richmond'; Week = $priorYear; RowKey = 'prior-year'; Values = [pscustomobject]@{ Service = $PriorYearValue; Count = 120 } },
        [pscustomobject]@{ EntityKey = 'Sorensen|(TOTAL)'; Week = $current; RowKey = 'sorensen'; Values = [pscustomobject]@{ Service = $CurrentValue; Count = 240 } },
        [pscustomobject]@{ EntityKey = 'All Franchisees|(TOTAL)'; Week = $current; RowKey = 'all'; Values = [pscustomobject]@{ Service = $CurrentValue; Count = 1000 } }
    )
    return @(New-GssMetricDetail $rows $metricSet $current $priorYear 4)[0]
}

Assert-Equal (Get-BoundaryMetricDetail 80 79.001 80).IsCandidate $false 'Previous-window candidate below one point'
Assert-Equal (Get-BoundaryMetricDetail 80 79 80).IsCandidate $true 'Previous-window candidate at one point'
Assert-Equal (Get-BoundaryMetricDetail 80 78.001 80).BaseActionItem $false 'Previous-window action below two points'
Assert-Equal (Get-BoundaryMetricDetail 80 78 80).BaseActionItem $true 'Previous-window action at two points'
Assert-Equal (Get-BoundaryMetricDetail 80 80 78.001).IsCandidate $false 'Prior-year candidate below two points'
Assert-Equal (Get-BoundaryMetricDetail 80 80 78).IsCandidate $true 'Prior-year candidate at two points'
Assert-Equal (Get-BoundaryMetricDetail 80 80 75.001).BaseActionItem $false 'Prior-year action below five points'
Assert-Equal (Get-BoundaryMetricDetail 80 80 75).BaseActionItem $true 'Prior-year action at five points'
$lowerBoundary = Get-BoundaryMetricDetail 4 6 4 $true
Assert-Equal $lowerBoundary.ChangeVsPreviousRollingWindow 2 'Lower-is-better direction adjustment at action boundary'
Assert-Equal $lowerBoundary.CandidateDirection 'Improvement' 'Lower-is-better improvement direction'
$mixedDirection = Get-BoundaryMetricDetail 80 77 84
Assert-Equal $mixedDirection.CandidateDirection 'Improvement' 'Eligible previous-window action is not hidden by larger sub-threshold YoY decline'
Assert-Equal $mixedDirection.CandidateComparison 'PreviousRollingWindow' 'Mixed-direction eligible comparison selection'

function New-TestFindingItem {
    param(
        [string]$RestaurantId,
        [string]$MetricName,
        [string]$Direction,
        [double]$Score,
        [bool]$BaseAction = $true,
        [int]$ResponseCount = 120
    )
    return [pscustomobject]@{
        EvidenceId = "metric-$RestaurantId-$MetricName"
        RestaurantId = $RestaurantId
        Entity = if ($RestaurantId -eq '9354') { '9354 Richmond' } else { '9355 Virginia Beach' }
        Metric = $MetricName
        RawMetric = $MetricName
        Category = if ($MetricName -like 'Service*') { 'Service' } else { 'Overall' }
        LowerIsBetter = $false
        Current = 80
        CurrentCount = $ResponseCount
        ChangeVsPreviousRollingWindow = if ($Direction -eq 'Improvement') { $Score } else { -$Score }
        YoYImprovement = 0
        VsAllFranchisees = 0
        VsSorensenTotal = 0
        PersistentMovement = $false
        IsCandidate = $true
        CandidateDirection = $Direction
        CandidateMagnitude = $Score
        CandidateThresholdMet = $BaseAction
        Corroboration = @()
        BaseActionItem = $BaseAction
    }
}

$capDetail = @(
    (New-TestFindingItem '9354' 'Opp A' 'Opportunity' 6),
    (New-TestFindingItem '9354' 'Opp B' 'Opportunity' 5),
    (New-TestFindingItem '9354' 'Opp C' 'Opportunity' 4),
    (New-TestFindingItem '9354' 'Strength A' 'Improvement' 4),
    (New-TestFindingItem '9354' 'Strength B' 'Improvement' 3),
    (New-TestFindingItem '9355' 'Opp D' 'Opportunity' 3),
    (New-TestFindingItem '9355' 'Strength C' 'Improvement' 2)
)
$capped = @(Select-GssRestaurantFindings $capDetail)
$richmondFindings = $capped | Where-Object RestaurantId -eq '9354'
$beachFindings = $capped | Where-Object RestaurantId -eq '9355'
Assert-Equal @($richmondFindings.Opportunities).Count 2 'Richmond opportunity cap'
Assert-Equal @($richmondFindings.Strengths).Count 1 'Richmond strength cap'
Assert-Equal @($beachFindings.Opportunities).Count 1 'Virginia Beach independent opportunity ranking'
Assert-Equal @($beachFindings.Strengths).Count 1 'Virginia Beach independent strength ranking'

$guestCandidate = New-TestFindingItem '9354' 'Service Candidate' 'Opportunity' 1.2 $false
$withoutGuest = @(Select-GssRestaurantFindings @($guestCandidate)) | Where-Object RestaurantId -eq '9354'
Assert-Equal @($withoutGuest.Opportunities).Count 0 'Sub-action candidate is omitted without corroboration'
$guestTheme = [pscustomobject]@{ restaurant_id = '9354'; category = 'service'; concern_count = 2; positive_count = 0; theme_id = 'theme-9354-service' }
$withGuest = @(Select-GssRestaurantFindings @($guestCandidate) @($guestTheme)) | Where-Object RestaurantId -eq '9354'
Assert-Equal @($withGuest.Opportunities).Count 1 'Guest feedback elevates a candidate to action item'
Assert-True (@($withGuest.Opportunities[0].Corroboration | Where-Object { $_ -eq 'guest_feedback:theme-9354-service' }).Count -eq 1) 'Guest corroboration evidence ID retained'

$developingThresholdOnly = New-TestFindingItem '9354' 'Service Developing Threshold Only' 'Opportunity' 3 $true 75
$developingWithoutCorroboration = @(Select-GssRestaurantFindings @($developingThresholdOnly)) | Where-Object RestaurantId -eq '9354'
Assert-Equal @($developingWithoutCorroboration.Opportunities).Count 0 'Developing confidence requires corroboration in addition to the point threshold'

$developingWithGuest = @(Select-GssRestaurantFindings @($developingThresholdOnly) @($guestTheme)) | Where-Object RestaurantId -eq '9354'
Assert-Equal @($developingWithGuest.Opportunities).Count 1 'Developing confidence qualifies with threshold and corroboration'
Assert-Equal $developingWithGuest.Opportunities[0].ConfidenceTier 'Developing' 'Developing confidence retained in selected finding'

$developingBelowThreshold = New-TestFindingItem '9354' 'Service Developing Below Threshold' 'Opportunity' 1.2 $false 75
$developingBelowThresholdWithGuest = @(Select-GssRestaurantFindings @($developingBelowThreshold) @($guestTheme)) | Where-Object RestaurantId -eq '9354'
Assert-Equal @($developingBelowThresholdWithGuest.Opportunities).Count 0 'Developing corroboration cannot replace the point threshold'

$lowWithGuest = New-TestFindingItem '9354' 'Service Low Confidence' 'Opportunity' 6 $true 49
$lowFindings = @(Select-GssRestaurantFindings @($lowWithGuest) @($guestTheme)) | Where-Object RestaurantId -eq '9354'
Assert-Equal @($lowFindings.Opportunities).Count 0 'Low confidence is never a top Opportunity'
Assert-Equal @($lowFindings.Strengths).Count 0 'Low confidence is never a top Strength'

$notScoredWithGuest = New-TestFindingItem '9354' 'Service Not Scored' 'Improvement' 6 $true 0
$notScoredFindings = @(Select-GssRestaurantFindings @($notScoredWithGuest) @($guestTheme)) | Where-Object RestaurantId -eq '9354'
Assert-Equal @($notScoredFindings.Opportunities).Count 0 'Not scored is never a top Opportunity'
Assert-Equal @($notScoredFindings.Strengths).Count 0 'Not scored is never a top Strength'

Write-Host 'GSS analytics logic tests passed.'
