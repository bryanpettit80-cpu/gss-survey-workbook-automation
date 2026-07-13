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
Assert-Equal ($guardrails.UnlockedCells -join ',') 'Exec_Dashboard!C3,Exec_Dashboard!F3' 'Only dashboard selectors unlocked'
Assert-Equal $guardrails.ValidationShowError $true 'Validation errors enabled'
Assert-Equal $guardrails.IntentionalGapIsBlocker $false 'Intentional 2025 gap is informational'

$qaPlan = @(Get-GssQaCheckPlan 5 14 5096)
Assert-Equal $qaPlan.Count 7 'QA check count'
Assert-True (($qaPlan | Where-Object Row -eq 7).Formula -like '*UPPER*Entities*TRUE*') 'Active entity text TRUE normalization'
Assert-True (($qaPlan | Where-Object Row -eq 7).Formula -like '*MMULT*') 'Current-period blank metric detection'
Assert-True (($qaPlan | Where-Object Row -eq 12).Formula -like '*>=8*') 'Capacity reserves two weekly loads'

Write-Host 'GSS non-Excel logic tests passed.'
