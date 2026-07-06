[CmdletBinding()]
param(
    [string]$Folder
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $scriptRoot = Split-Path -Parent $PSCommandPath
    $projectRoot = Split-Path -Parent $scriptRoot
    $Folder = Split-Path -Parent $projectRoot
}

$Folder = $Folder.Trim('"')
$Folder = (Resolve-Path -LiteralPath $Folder).Path
$scriptRoot = Split-Path -Parent $PSCommandPath
$updater = Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1'
$analyzer = Join-Path $scriptRoot 'Analyze-GSS-Run.ps1'

Write-Host 'Running GSS copy-test first. This does not change the live workbook.'
& $updater -Folder $Folder

Write-Host 'Running GSS analytics review for the copy-test.'
$copyReview = & $analyzer -Folder $Folder -OutputObject
if ($copyReview.OverallStatus -eq 'Blocked') {
    Write-Host 'Analytics review found a blocker. Live workbook update skipped.'
    exit 1
}

Write-Host 'Copy-test finished. Review the workbook summary and analytics review above before applying to the live workbook.'
Write-Host 'Type APPLY and press Enter to update the live workbook. Press Enter to stop here.'
$confirmation = Read-Host 'Confirmation'

if ($confirmation -ne 'APPLY') {
    Write-Host 'Live workbook update skipped. No live changes were applied.'
    exit 0
}

Write-Host 'Applying update to the live workbook. A backup will be saved first.'
& $updater -Folder $Folder -Apply

Write-Host 'Running GSS analytics review for the live workbook update.'
$liveReview = & $analyzer -Folder $Folder -OutputObject
if ($liveReview.OverallStatus -eq 'Blocked') {
    Write-Host 'Live update finished, but analytics review found a blocker. Review the analytics output before using the report.'
    exit 1
}
