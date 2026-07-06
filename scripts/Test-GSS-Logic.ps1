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

Write-Host 'GSS non-Excel logic tests passed.'
