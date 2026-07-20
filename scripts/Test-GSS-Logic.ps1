[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1')

function Assert-Equal {
    param(
        [object]$Actual,
        [object]$Expected,
        [string]$Name
    )

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

function Assert-ThrowsLike {
    param(
        [scriptblock]$Action,
        [string]$ExpectedMessage,
        [string]$Name
    )

    try {
        & $Action
    }
    catch {
        if ($_.Exception.Message -like $ExpectedMessage) {
            return
        }
        throw "Assertion failed for $Name. Expected error like '$ExpectedMessage'; actual '$($_.Exception.Message)'."
    }
    throw "Assertion failed for $Name. Expected an error, but no error was thrown."
}

function New-TestSource {
    param([string]$Path, [datetime]$WeekEnding, [datetime]$LastWriteTime)

    [pscustomobject]@{
        WeekEnding = $WeekEnding
        File = [pscustomobject]@{
            FullName = $Path
            LastWriteTime = $LastWriteTime
        }
    }
}

Assert-Equal (Normalize-Header 'Pace Of Meal Dissat (lower is better)') 'paceofmealdissatlowerisbetter' 'Normalize-Header punctuation removal'
Assert-Equal (Normalize-Header ' Restaurant ') 'restaurant' 'Normalize-Header whitespace'

$folder = 'C:\GSS Surveys'
Assert-Equal (Test-ExcludedGssPath 'C:\GSS Surveys\GSS Survey Workbook Automation\scripts\Update.ps1' $folder) $true 'Repo folder excluded'
Assert-Equal (Test-ExcludedGssPath 'C:\GSS Surveys\_automation_runs\logs\run.json' $folder) $true 'Automation output excluded'
Assert-Equal (Test-ExcludedGssPath 'C:\GSS Surveys\02 Weekly Rolling Source Workbooks\Sorensen FW01.xlsx' $folder) $false 'Source folder included'

$sources = @(
    (New-TestSource 'older-same-week.xlsx' ([datetime]'2026-06-28') ([datetime]'2026-07-01T08:00:00')),
    (New-TestSource 'newer-same-week.xlsx' ([datetime]'2026-06-28') ([datetime]'2026-07-01T09:00:00')),
    (New-TestSource 'prior-week.xlsx' ([datetime]'2026-06-21') ([datetime]'2026-07-02T09:00:00'))
)

$latest = Select-LatestRollingSource $sources
Assert-Equal $latest.File.FullName 'newer-same-week.xlsx' 'Latest source uses week then last write time'

$priorYear = Select-RollingSourceByWeek @(
    (New-TestSource 'wrong.xlsx' ([datetime]'2025-06-22') ([datetime]'2025-06-22')),
    (New-TestSource 'right.xlsx' ([datetime]'2025-06-29') ([datetime]'2025-06-29'))
) ([datetime]'2025-06-29')
Assert-Equal $priorYear.File.FullName 'right.xlsx' 'Prior-year source selection'

$layout = @(Get-GssWorkbookLayoutPlan 4 13)
Assert-Equal (($layout | Where-Object Sheet -eq 'README').PrintArea) '$A$1:$A$34' 'README print area'
Assert-Equal (($layout | Where-Object Sheet -eq 'README').FitTall) 2 'README two-page maximum'
Assert-Equal (($layout | Where-Object Sheet -eq 'Quick_Read_WoW').PrintArea) '$A$1:$Q$9' 'WoW dynamic print area'
Assert-Equal (($layout | Where-Object Sheet -eq 'Quick_Read_YoY').PrintArea) '$A$1:$Q$9' 'YoY dynamic print area'
Assert-Equal (($layout | Where-Object Sheet -eq 'Drivers_Detail').PrintArea) '$A$1:$AB$18' 'Drivers dynamic print area'
Assert-Equal (($layout | Where-Object Sheet -eq 'Exec_Dashboard').PrintArea) '$A$1:$AD$57' 'Dashboard print area includes chart'
Assert-Equal (($layout | Where-Object Sheet -eq 'Email Comparison').PrintArea) '$A$1:$AB$38' 'Email comparison print area'

$guardrails = Get-GssGuardrailPlan
Assert-Equal $guardrails.ProtectWorkbookStructure $true 'Workbook structure protection plan'
Assert-Equal $guardrails.ProtectAllWorksheets $true 'Worksheet protection plan'
Assert-Equal $guardrails.LockScope 'AllCells' 'Worksheet lock scope'
Assert-Equal ($guardrails.UnlockedCells -join ',') 'Exec_Dashboard!C3,Exec_Dashboard!F3' 'Only dashboard selectors unlocked'
Assert-Equal $guardrails.ValidationShowError $true 'Validation errors enabled'
Assert-Equal $guardrails.IntentionalGapIsBlocker $false 'Intentional 2025 gap is informational'

$qaPlan = @(Get-GssQaCheckPlan 5 14 5096)
Assert-Equal $qaPlan.Count 7 'QA check count'
Assert-True (($qaPlan | Where-Object Row -eq 7).Formula -like '*UPPER*Entities*TRUE*') 'Active entity text TRUE normalization'
Assert-True (($qaPlan | Where-Object Row -eq 7).Formula.Contains('="1"')) 'Active entity numeric 1 normalization'
Assert-True (($qaPlan | Where-Object Row -eq 11).Formula.Contains('="1"')) 'Dashboard selector numeric 1 normalization'
Assert-True (($qaPlan | Where-Object Row -eq 7).Formula -like '*MMULT*') 'Current-period blank metric detection'
Assert-True (($qaPlan | Where-Object Row -eq 8).Formula -like '*$B$3-7>DATE(2025,7,6)*$B$3-7<DATE(2025,10,5)*') 'Prior-week intentional gap exemption'
Assert-True (($qaPlan | Where-Object Row -eq 8).DetailFormula -like '*Intentional 2025 history gap*') 'Prior-week intentional gap detail'
Assert-True (-not (($qaPlan | Where-Object Row -eq 7).Formula -like '*DATE(2025,7,6)*')) 'Current week still requires complete data'
Assert-True (($qaPlan | Where-Object Row -eq 12).Formula -like '*>=8*') 'Capacity reserves two weekly loads'

$capacityAtLimit = Get-GssRawDataCapacityPlan 5088 8 5096
Assert-Equal $capacityAtLimit.AvailableRows 8 'Capacity rows available'
Assert-Equal $capacityAtLimit.ProjectedLastRow 5096 'Capacity projected last row'
Assert-Equal $capacityAtLimit.Fits $true 'Capacity accepts exact formula bound'
$capacityOverLimit = Get-GssRawDataCapacityPlan 5089 8 5096
Assert-Equal $capacityOverLimit.Fits $false 'Capacity rejects rows beyond formula bound'
Assert-ThrowsLike { Assert-GssRawDataCapacity 5089 8 5096 } '*only 7 row(s) remain through row 5096*No rows were appended.*' 'Capacity failure occurs before append'

$updaterSource = Get-Content -LiteralPath (Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1') -Raw
$capacityCallIndex = $updaterSource.IndexOf('$capacityPlan = Assert-GssRawDataCapacity')
$guardrailRemovalIndex = $updaterSource.IndexOf('Remove-GssWorkbookGuardrails $targetWb')
$appendCallIndex = $updaterSource.IndexOf('$written = Add-RawDataRows')
Assert-True ($capacityCallIndex -ge 0 -and $capacityCallIndex -lt $guardrailRemovalIndex) 'Capacity preflight precedes workbook mutation'
Assert-True ($capacityCallIndex -lt $appendCallIndex) 'Capacity preflight precedes Raw_Data append'

$allCellsAcquireIndex = $updaterSource.IndexOf('$allCells = $worksheet.Cells')
$allCellsLockIndex = $updaterSource.IndexOf('$allCells.Locked = $true')
$dashboardUnlockIndex = $updaterSource.IndexOf('Set-GssDashboardValidation $Workbook $Configuration')
Assert-True ($allCellsAcquireIndex -ge 0 -and $allCellsAcquireIndex -lt $allCellsLockIndex) 'Guardrails acquire all worksheet cells before locking'
Assert-True ($allCellsLockIndex -lt $dashboardUnlockIndex) 'All worksheet cells are locked before dashboard selectors are unlocked'

Write-Host 'GSS non-Excel logic tests passed.'
