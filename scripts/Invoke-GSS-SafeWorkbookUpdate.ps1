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

$Folder = (Resolve-Path -LiteralPath $Folder).Path
$scriptRoot = Split-Path -Parent $PSCommandPath
$updater = Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1'

Write-Host 'Running GSS copy-test first. This does not change the live workbook.'
& $updater -Folder $Folder

Write-Host 'Copy-test finished. Review the summary above before applying to the live workbook.'
Write-Host 'Type APPLY and press Enter to update the live workbook. Press Enter to stop here.'
$confirmation = Read-Host 'Confirmation'

if ($confirmation -ne 'APPLY') {
    Write-Host 'Live workbook update skipped. No live changes were applied.'
    exit 0
}

Write-Host 'Applying update to the live workbook. A backup will be saved first.'
& $updater -Folder $Folder -Apply
