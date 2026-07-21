[CmdletBinding()]
param(
    [string]$Folder,
    [string]$LogPath,
    [int]$LookbackWeeks = 8,
    [string]$MainWorkbookName = 'GSS Score Trends - Main.xlsx',
    [switch]$OutputObject,
    [switch]$PublishEmailPackage
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Gss-Common.ps1')
. (Join-Path $scriptRoot 'Gss-EmailPackage.ps1')

$xlUp = -4162
$xlToLeft = -4159

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

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

function Resolve-GssAnalysisWorkbookPath {
    param(
        [object]$RunLog,
        [string]$FolderPath,
        [string]$WorkbookName
    )

    if ($RunLog.TargetWorkbook -and -not [string]::IsNullOrWhiteSpace([string]$RunLog.TargetWorkbook)) {
        return Resolve-GssDropboxPath -Path ([string]$RunLog.TargetWorkbook) -FolderPath $FolderPath
    }

    return Resolve-MainWorkbookPath $FolderPath $WorkbookName
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
    $qaWs = $null
    $qaStatusCell = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false

        $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
        $rawWs = $workbook.Worksheets.Item('Raw_Data')
        $metricWs = $workbook.Worksheets.Item('Metric_Catalog')
        $qaSheetPresent = $false
        $workbookQaStatus = $null
        try {
            $qaWs = $workbook.Worksheets.Item('QA Checks')
            $qaSheetPresent = $true
            $qaStatusCell = $qaWs.Range('B2')
            $workbookQaStatus = $qaStatusCell.Value2
        }
        catch {
            $qaSheetPresent = $false
            $workbookQaStatus = $null
        }

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
            QaSheetPresent = $qaSheetPresent
            WorkbookQaStatus = $workbookQaStatus
        }
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        Release-ComObject $qaStatusCell
        Release-ComObject $qaWs
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

            # Every Raw_Data value is already a 13-week rolling result. Never average
            # multiple rolling rows together; compare adjacent rolling windows directly.
            $twoWindowsAgoRow = Get-GssRow $rowsByKey $entity.EntityKey $CurrentWeek.AddDays(-14)
            $threeWindowsAgoRow = Get-GssRow $rowsByKey $entity.EntityKey $CurrentWeek.AddDays(-21)
            $twoWindowsAgoValue = Get-GssMetricValue $twoWindowsAgoRow $metric.RawKey
            $threeWindowsAgoValue = Get-GssMetricValue $threeWindowsAgoRow $metric.RawKey
            $currentCount = Get-GssMetricValue $currentRow 'Count'

            $wow = Get-DirectionAdjustedChange $currentValue $priorWeekValue $metric.LowerIsBetter
            $yoy = Get-DirectionAdjustedChange $currentValue $priorYearValue $metric.LowerIsBetter
            $vsSorensen = Get-DirectionAdjustedChange $currentValue $sorensenValue $metric.LowerIsBetter
            $vsAll = Get-DirectionAdjustedChange $currentValue $allValue $metric.LowerIsBetter
            $previousStep = Get-DirectionAdjustedChange $priorWeekValue $twoWindowsAgoValue $metric.LowerIsBetter
            $earlierStep = Get-DirectionAdjustedChange $twoWindowsAgoValue $threeWindowsAgoValue $metric.LowerIsBetter
            $movementComparisons = @(
                [pscustomobject]@{ Label = 'PreviousRollingWindow'; Value = $wow },
                [pscustomobject]@{ Label = 'PriorYearRollingWindow'; Value = $yoy }
            ) | Where-Object { $null -ne $_.Value }
            $worstMovement = $null
            $worstMovementLabel = ''
            $bestMovement = $null
            $bestMovementLabel = ''
            if ($movementComparisons.Count -gt 0) {
                $worstComparison = $movementComparisons | Sort-Object Value | Select-Object -First 1
                $bestComparison = $movementComparisons | Sort-Object @{ Expression = { $_.Value }; Descending = $true } | Select-Object -First 1
                $worstMovement = $worstComparison.Value
                $worstMovementLabel = $worstComparison.Label
                $bestMovement = $bestComparison.Value
                $bestMovementLabel = $bestComparison.Label
            }

            $candidateComparisons = @()
            if ($null -ne $wow -and [math]::Abs([double]$wow) -ge 1) {
                $candidateComparisons += [pscustomobject]@{ Label = 'PreviousRollingWindow'; Value = [double]$wow; ActionThresholdMet = ([math]::Abs([double]$wow) -ge 2); TieOrder = 0 }
            }
            if ($null -ne $yoy -and [math]::Abs([double]$yoy) -ge 2) {
                $candidateComparisons += [pscustomobject]@{ Label = 'PriorYearRollingWindow'; Value = [double]$yoy; ActionThresholdMet = ([math]::Abs([double]$yoy) -ge 5); TieOrder = 1 }
            }
            $evaluatedCandidates = @()
            foreach ($candidateComparison in $candidateComparisons) {
                $directionSign = if ($candidateComparison.Value -gt 0) { 1 } else { -1 }
                $recentSteps = @($wow, $previousStep, $earlierStep) | Where-Object { $null -ne $_ }
                $sameDirectionSteps = @($recentSteps | Where-Object { ([math]::Sign([double]$_)) -eq $directionSign })
                $isPersistent = ($null -ne $wow -and ([math]::Sign([double]$wow)) -eq $directionSign -and $sameDirectionSteps.Count -ge 2)
                $candidateCorroboration = @()
                if ($isPersistent) { $candidateCorroboration += 'persistence' }
                if ($null -ne $vsAll -and [math]::Abs([double]$vsAll) -ge 1 -and ([math]::Sign([double]$vsAll)) -eq $directionSign) {
                    $candidateCorroboration += 'franchise_gap'
                }
                $evaluatedCandidates += [pscustomobject]@{
                    Label = $candidateComparison.Label
                    Value = $candidateComparison.Value
                    ActionThresholdMet = $candidateComparison.ActionThresholdMet
                    TieOrder = $candidateComparison.TieOrder
                    Persistent = $isPersistent
                    Corroboration = $candidateCorroboration
                    IsAction = ([bool]$candidateComparison.ActionThresholdMet -or $candidateCorroboration.Count -gt 0)
                }
            }
            # Choose among eligible action comparisons first. This prevents a larger,
            # sub-threshold comparison in the opposite direction from hiding a valid action.
            $actionCandidates = @($evaluatedCandidates | Where-Object { $_.IsAction })
            $candidatePool = if ($actionCandidates.Count -gt 0) { $actionCandidates } else { $evaluatedCandidates }
            $candidate = $candidatePool |
                Sort-Object @{ Expression = { [bool]$_.ActionThresholdMet }; Descending = $true }, @{ Expression = { [math]::Abs($_.Value) }; Descending = $true }, TieOrder |
                Select-Object -First 1
            $candidateDirection = $null
            $persistentMovement = $false
            $corroboration = @()
            $baseAction = $false
            if ($candidate) {
                $candidateDirection = if ($candidate.Value -gt 0) { 'Improvement' } else { 'Opportunity' }
                $persistentMovement = [bool]$candidate.Persistent
                $corroboration = @($candidate.Corroboration)
                $baseAction = [bool]$candidate.IsAction
            }

            $restaurantIdMatch = [regex]::Match($entity.Label, '^\s*(\d{4})\b')
            $restaurantId = if ($restaurantIdMatch.Success) { $restaurantIdMatch.Groups[1].Value } else { Normalize-Header $entity.Label }
            $evidenceMaterial = "$($entity.EntityKey)|$($metric.RawKey)|$($CurrentWeek.ToString('yyyy-MM-dd'))"

            $details += [pscustomobject]@{
                EvidenceId = 'metric-' + (Get-GssStringSha256 $evidenceMaterial).Substring(0, 16)
                Entity = $entity.Label
                RestaurantId = $restaurantId
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
                RollingAverage = $null
                PreviousRollingWindow = $priorWeekValue
                CurrentCount = $currentCount
                WoWImprovement = $wow
                ChangeVsPreviousRollingWindow = $wow
                YoYImprovement = $yoy
                ChangeVsPriorYearRollingWindow = $yoy
                VsRollingAverage = $null
                VsSorensenTotal = $vsSorensen
                VsAllFranchisees = $vsAll
                PreviousStepImprovement = $previousStep
                EarlierStepImprovement = $earlierStep
                IsCandidate = ($null -ne $candidate)
                CandidateDirection = $candidateDirection
                CandidateComparison = if ($candidate) { $candidate.Label } else { $null }
                CandidateMagnitude = if ($candidate) { [math]::Abs([double]$candidate.Value) } else { $null }
                PersistentMovement = $persistentMovement
                Corroboration = $corroboration
                BaseActionItem = $baseAction
                WorstMovement = $worstMovement
                WorstMovementLabel = $worstMovementLabel
                BestMovement = $bestMovement
                BestMovementLabel = $bestMovementLabel
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
        [datetime]$PriorYearWeek,
        [bool]$QaSheetPresent = $false,
        [object]$WorkbookQaStatus = $null
    )

    $workbookBlockers = @()
    $analysisBlockers = @()
    $emailBlockers = @()
    $warnings = @()
    $rowsByKey = @{}
    foreach ($row in $RawRows) {
        $key = Get-GssEntityWeekKey $row.EntityKey $row.Week
        $rowsByKey[$key] = $row
    }

    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
        $workbookBlockers += "Main workbook is missing: $WorkbookPath"
    }
    if ($RunLog.EmailComparisonPdf) {
        try {
            $resolvedPdf = Resolve-GssDropboxPath -Path ([string]$RunLog.EmailComparisonPdf) -FolderPath $FolderPath -RequireFile
        }
        catch {
            $emailBlockers += $_.Exception.Message
        }
    }
    else {
        $emailBlockers += 'Email comparison PDF is not recorded in the run log.'
    }
    if ($RunLog.Mode -eq 'ApplyToMainWorkbook') {
        if (-not $RunLog.BackupWorkbook) {
            $workbookBlockers += 'Live apply backup workbook is missing.'
        }
        else {
            try { $null = Resolve-GssDropboxPath -Path ([string]$RunLog.BackupWorkbook) -FolderPath $FolderPath -RequireFile }
            catch { $workbookBlockers += 'Live apply backup workbook is missing.' }
        }
    }
    if (-not (Test-GssYoyPairing $CurrentWeek $PriorYearWeek)) {
        $workbookBlockers += "Reporting date and prior-year rolling date are not 364 days apart."
    }

    if (-not $QaSheetPresent) {
        $warnings += 'QA Checks sheet is missing. This is allowed only for a legacy workbook; run the current updater to add workbook guardrails.'
    }
    else {
        $normalizedWorkbookQaStatus = if ($null -eq $WorkbookQaStatus) { '' } else { ([string]$WorkbookQaStatus).Trim().ToUpperInvariant() }
        if ([string]::IsNullOrWhiteSpace($normalizedWorkbookQaStatus)) {
            $warnings += 'QA Checks status is blank. Run the current updater to recalculate workbook QA.'
        }
        elseif ($normalizedWorkbookQaStatus -eq 'ATTENTION') {
            $workbookBlockers += 'Workbook QA Checks reports ATTENTION. Open QA Checks and resolve every flagged item before using the reports.'
        }
        elseif ($normalizedWorkbookQaStatus -ne 'READY') {
            $workbookBlockers += "Workbook QA Checks returned an invalid status: $normalizedWorkbookQaStatus"
        }
    }

    foreach ($sourcePath in @($RunLog.CurrentSourceWorkbook, $RunLog.PriorYearSourceWorkbook)) {
        if (-not $sourcePath) {
            $emailBlockers += 'A rolling source workbook is not recorded in the run log.'
            continue
        }
        try {
            $resolvedSource = Resolve-GssDropboxPath -Path ([string]$sourcePath) -FolderPath $FolderPath -RequireFile
            $relativeSource = ConvertTo-GssDropboxRelativePath -Path $resolvedSource -FolderPath $FolderPath
            if (-not $relativeSource.StartsWith('02 Weekly Rolling Source Workbooks/', [System.StringComparison]::OrdinalIgnoreCase)) {
                $emailBlockers += "Rolling source workbook is outside 02 Weekly Rolling Source Workbooks: $relativeSource"
            }
        }
        catch {
            $emailBlockers += $_.Exception.Message
        }
    }

    $requiredEntities = @(Get-RequiredGssEntities)
    foreach ($week in @($CurrentWeek, $PriorYearWeek)) {
        foreach ($entity in $requiredEntities) {
            $key = Get-GssEntityWeekKey $entity.EntityKey $week
            if (-not $rowsByKey.ContainsKey($key)) {
                $workbookBlockers += "Missing Raw_Data row for $($entity.Label) on $($week.ToString('yyyy-MM-dd'))."
            }
        }
    }

    $duplicateRowKeys = @($RawRows | Group-Object RowKey | Where-Object { $_.Name -and $_.Count -gt 1 })
    foreach ($duplicate in $duplicateRowKeys) {
        $workbookBlockers += "Duplicate Raw_Data RowKey: $($duplicate.Name)"
    }

    foreach ($metric in $Metrics) {
        if (-not $Headers.Contains($metric.RawKey)) {
            $workbookBlockers += "Raw_Data is missing Metric_Catalog column: $($metric.RawKey)"
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

    $workbookStatus = if ($workbookBlockers.Count -gt 0) { 'Blocked' } else { 'Ready' }
    $analysisStatus = if ($analysisBlockers.Count -gt 0) {
        'Blocked'
    }
    elseif ($warnings.Count -gt 0) {
        'Review'
    }
    else {
        'Ready'
    }
    $emailReadiness = if ($emailBlockers.Count -gt 0) { 'Blocked' } else { 'Ready' }
    $status = if ($workbookStatus -eq 'Blocked') { 'Blocked' } elseif ($analysisStatus -ne 'Ready' -or $emailReadiness -ne 'Ready') { 'Review' } else { 'Ready' }

    return [pscustomobject]@{
        Status = $status
        WorkbookStatus = $workbookStatus
        AnalysisStatus = $analysisStatus
        EmailReadiness = $emailReadiness
        WorkbookBlockers = @($workbookBlockers | Select-Object -Unique)
        AnalysisBlockers = @($analysisBlockers | Select-Object -Unique)
        EmailBlockers = @($emailBlockers | Select-Object -Unique)
        Blockers = @($workbookBlockers + $analysisBlockers + $emailBlockers | Select-Object -Unique)
        Warnings = @($warnings | Select-Object -Unique)
    }
}

function Select-GssAttentionItems {
    param([object[]]$MetricDetail, [int]$Count = 5)

    return @($MetricDetail |
        Where-Object { $_.WorstMovement -ne $null -and $_.WorstMovement -lt 0 } |
        Sort-Object @{ Expression = { $_.WorstMovement }; Ascending = $true }, @{ Expression = { $_.CurrentCount }; Descending = $true } |
        Select-Object -First $Count)
}

function Select-GssStrengthItems {
    param([object[]]$MetricDetail, [int]$Count = 5)

    return @($MetricDetail |
        Where-Object { $_.BestMovement -ne $null -and $_.BestMovement -gt 0 -and ($_.WorstMovement -eq $null -or $_.WorstMovement -ge 0) } |
        Sort-Object @{ Expression = { $_.BestMovement }; Descending = $true }, @{ Expression = { $_.CurrentCount }; Descending = $true }, @{ Expression = { $_.Entity }; Descending = $true } |
        Select-Object -First $Count)
}

function Get-GssMetricFeedbackCategory {
    param([object]$MetricDetail)

    $name = (Normalize-Header "$($MetricDetail.Category) $($MetricDetail.Metric) $($MetricDetail.RawMetric)")
    if ($name -match 'pace') { return 'pace' }
    if ($name -match 'value') { return 'value' }
    if ($name -match 'culinary|steak|food') { return 'culinary' }
    if ($name -match 'service') { return 'service' }
    if ($name -match 'manager|overall|experience') { return 'hospitality/recovery' }
    return $null
}

function Select-GssRestaurantFindings {
    param(
        [object[]]$MetricDetail,
        [object[]]$GuestThemes = @()
    )

    $results = @()
    foreach ($entity in @(Get-RequiredGssEntities | Where-Object { $_.IncludeInInsights })) {
        $restaurantIdMatch = [regex]::Match($entity.Label, '^\s*(\d{4})\b')
        $restaurantId = if ($restaurantIdMatch.Success) { $restaurantIdMatch.Groups[1].Value } else { Normalize-Header $entity.Label }
        $actionItems = @()
        foreach ($detail in @($MetricDetail | Where-Object { $_.RestaurantId -eq $restaurantId -and $_.IsCandidate })) {
            $copy = $detail | Select-Object *
            $corroboration = @($copy.Corroboration)
            $feedbackCategory = Get-GssMetricFeedbackCategory $copy
            if ($feedbackCategory) {
                $matchingThemes = @($GuestThemes | Where-Object { $_.restaurant_id -eq $restaurantId -and $_.category -eq $feedbackCategory })
                foreach ($theme in $matchingThemes) {
                    $supportsDirection = if ($copy.CandidateDirection -eq 'Opportunity') {
                        [int]$theme.concern_count -ge 2
                    }
                    else {
                        [int]$theme.positive_count -ge 2
                    }
                    if ($supportsDirection) { $corroboration += "guest_feedback:$($theme.theme_id)" }
                }
            }
            $copy.Corroboration = @($corroboration | Select-Object -Unique)
            $isAction = [bool]$copy.BaseActionItem -or @($copy.Corroboration | Where-Object { $_ -like 'guest_feedback:*' }).Count -gt 0
            $copy | Add-Member -NotePropertyName IsActionItem -NotePropertyValue $isAction -Force
            $directionSign = if ($copy.CandidateDirection -eq 'Improvement') { 1 } else { -1 }
            $eligibleMagnitudes = @()
            foreach ($value in @($copy.ChangeVsPreviousRollingWindow, $copy.YoYImprovement)) {
                if ($null -ne $value -and [math]::Sign([double]$value) -eq $directionSign) { $eligibleMagnitudes += [math]::Abs([double]$value) }
            }
            $score = if ($eligibleMagnitudes.Count -gt 0) { ($eligibleMagnitudes | Measure-Object -Maximum).Maximum } else { 0 }
            $score += (0.01 * @($copy.Corroboration).Count)
            $copy | Add-Member -NotePropertyName ActionScore -NotePropertyValue $score -Force
            if ($isAction) { $actionItems += $copy }
        }

        $opportunities = @($actionItems |
            Where-Object { $_.CandidateDirection -eq 'Opportunity' } |
            Sort-Object @{ Expression = { $_.ActionScore }; Descending = $true }, @{ Expression = { $_.CurrentCount }; Descending = $true }, Metric |
            Select-Object -First 2)
        $opportunityIds = @($opportunities.EvidenceId)
        $strengths = @($actionItems |
            Where-Object { $_.CandidateDirection -eq 'Improvement' -and $_.EvidenceId -notin $opportunityIds } |
            Sort-Object @{ Expression = { $_.ActionScore }; Descending = $true }, @{ Expression = { $_.CurrentCount }; Descending = $true }, Metric |
            Select-Object -First 1)
        $results += [pscustomobject]@{
            RestaurantId = $restaurantId
            Restaurant = $entity.Label
            Name = ([regex]::Replace($entity.Label, '^\s*\d{4}\s+', '')).Trim()
            Strengths = $strengths
            Opportunities = $opportunities
        }
    }
    return $results
}

function Format-GssNumber {
    param([object]$Value, [int]$Digits = 1)

    if ($null -eq $Value) { return 'n/a' }
    $rounded = [math]::Round([double]$Value, $Digits, [System.MidpointRounding]::AwayFromZero)
    return ('{0:N' + $Digits + '}') -f $rounded
}

function Format-GssMovementNumber {
    param([object]$Value, [int]$Digits = 1)

    if ($null -eq $Value) { return 'n/a' }
    $number = [math]::Round([double]$Value, $Digits, [System.MidpointRounding]::AwayFromZero)
    $formatted = ('{0:N' + $Digits + '}') -f $number
    if ($number -gt 0) {
        return "+$formatted"
    }
    return $formatted
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
    $lines += "Workbook status: $($Result.WorkbookStatus)"
    $lines += "Analysis status: $($Result.AnalysisStatus)"
    $lines += "Email readiness: $($Result.EmailReadiness)"
    $lines += "Run mode: $($Result.Run.Mode)"
    $lines += "13-week rolling through: $($Result.Run.CurrentWeekEnding)"
    $lines += "Prior-year rolling through: $($Result.Run.PriorYearWeekEnding)"
    $lines += "Rows appended: $($Result.Run.RowsAppended)"
    $lines += "Workbook QA: $($Result.WorkbookQaStatus)"
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
            $lines += ('- {0}: {1} 13-week rolling value {2}; change versus previous rolling window {3}; change versus prior-year rolling window {4}; versus all franchisees {5}.' -f `
                $item.Entity, $item.Metric, (Format-GssNumber $item.Current), (Format-GssMovementNumber $item.ChangeVsPreviousRollingWindow), (Format-GssMovementNumber $item.YoYImprovement), (Format-GssMovementNumber $item.VsAllFranchisees))
        }
    }
    $lines += ''
    $lines += '## Strengths'
    if ($StrengthItems.Count -eq 0) {
        $lines += '- No positive movement items ranked above neutral.'
    }
    else {
        foreach ($item in $StrengthItems) {
            $lines += ('- {0}: {1} 13-week rolling value {2}; change versus previous rolling window {3}; change versus prior-year rolling window {4}; versus Sorensen total {5}.' -f `
                $item.Entity, $item.Metric, (Format-GssNumber $item.Current), (Format-GssMovementNumber $item.ChangeVsPreviousRollingWindow), (Format-GssMovementNumber $item.YoYImprovement), (Format-GssMovementNumber $item.VsSorensenTotal))
        }
    }
    $lines += ''
    $lines += '## Files'
    $lines += "- Detail CSV: $($Result.DetailCsvPath)"
    $lines += "- JSON: $($Result.JsonPath)"
    $lines += "- Review folder: $($Result.ReviewFolder)"
    $lines += ''
    $lines += 'These comparisons are directional. They do not establish statistical significance or causation.'
    $lines += ''

    return ($lines -join [Environment]::NewLine)
}

function Write-GssAnalysisSummary {
    param([object]$Result)

    Write-Host ''
    Write-Host 'GSS analytics review summary'
    Write-Host ('  Status: {0}' -f $Result.OverallStatus)
    Write-Host ('  13-week rolling through: {0}' -f $Result.Run.CurrentWeekEnding)
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
        [switch]$ReturnObject,
        [switch]$PublishPackage
    )

    $FolderPath = Resolve-GssAnalysisFolder $FolderPath
    if ([string]::IsNullOrWhiteSpace($RunLogPath)) {
        $RunLogPath = Get-LatestGssRunLog $FolderPath
    }
    else {
        $RunLogPath = (Resolve-Path -LiteralPath $RunLogPath.Trim('"')).Path
    }

    $runLog = Read-GssRunLog $RunLogPath
    $workbookPath = Resolve-GssAnalysisWorkbookPath $runLog $FolderPath $WorkbookName
    $workbookData = Read-GssWorkbookData $workbookPath
    $currentWeek = [datetime]::Parse($runLog.CurrentWeekEnding).Date
    $priorYearWeek = [datetime]::Parse($runLog.PriorYearWeekEnding).Date
    $metricDetail = @(New-GssMetricDetail $workbookData.RawRows $workbookData.Metrics $currentWeek $priorYearWeek $Weeks)
    $qa = Get-GssRunQa $runLog $FolderPath $workbookPath $workbookData.RawRows $workbookData.Metrics $workbookData.Headers $metricDetail $currentWeek $priorYearWeek $workbookData.QaSheetPresent $workbookData.WorkbookQaStatus
    $restaurantFindings = @(Select-GssRestaurantFindings -MetricDetail $metricDetail)
    $attentionItems = @($restaurantFindings | ForEach-Object { $_.Opportunities })
    $strengthItems = @($restaurantFindings | ForEach-Object { $_.Strengths })

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reviewFolder = Join-Path (Join-Path (Join-Path $FolderPath '_automation_runs') 'qa') "run_review_$timestamp"
    New-Item -ItemType Directory -Path $reviewFolder -Force | Out-Null
    $markdownPath = Join-Path $reviewFolder 'review.md'
    $jsonPath = Join-Path $reviewFolder 'review.json'
    $detailCsvPath = Join-Path $reviewFolder 'metric_detail.csv'

    $result = [pscustomobject]@{
        OverallStatus = $qa.Status
        WorkbookStatus = $qa.WorkbookStatus
        AnalysisStatus = $qa.AnalysisStatus
        EmailReadiness = if ($PublishPackage) { $qa.EmailReadiness } else { 'NotEvaluated' }
        WorkbookQaStatus = $workbookData.WorkbookQaStatus
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
        RestaurantFindings = $restaurantFindings
        MetricDetail = $metricDetail
        EmailPackage = $null
        RawDataSummary = [pscustomobject]@{
            RowCount = $workbookData.RawRows.Count
            WeekCount = @($workbookData.RawRows | Select-Object -ExpandProperty Week -Unique).Count
            MetricCount = $workbookData.Metrics.Count
        }
    }

    if ($PublishPackage) {
        try {
            $package = New-GssEmailPackage -FolderPath $FolderPath -RunLog $runLog -AnalysisResult $result
            $result.EmailPackage = $package
            $result.EmailReadiness = $package.EmailReadiness
            $result.Qa.EmailReadiness = $package.EmailReadiness
        }
        catch {
            $packageBlocker = $_.Exception.Message
            $result.EmailReadiness = 'Blocked'
            $result.Qa.EmailReadiness = 'Blocked'
            $result.Qa.EmailBlockers = @($result.Qa.EmailBlockers) + $packageBlocker
            $result.Qa.Blockers = @($result.Qa.Blockers) + $packageBlocker
            if ($result.WorkbookStatus -ne 'Blocked') { $result.OverallStatus = 'Review' }
        }
        $attentionItems = @($result.RestaurantFindings | ForEach-Object { $_.Opportunities })
        $strengthItems = @($result.RestaurantFindings | ForEach-Object { $_.Strengths })
        $result.TopAttention = $attentionItems
        $result.TopStrengths = $strengthItems
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
    Invoke-GssRunAnalysis -FolderPath $Folder -RunLogPath $LogPath -Weeks $LookbackWeeks -WorkbookName $MainWorkbookName -ReturnObject:$OutputObject -PublishPackage:$PublishEmailPackage
}
