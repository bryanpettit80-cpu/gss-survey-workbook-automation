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

Write-Host ''
Write-Host 'STEP 1 OF 3 - Testing the update'
Write-Host 'The live workbook is not being changed.'
& $updater -Folder $Folder

Write-Host ''
Write-Host 'STEP 2 OF 3 - Checking the test results'
$copyReview = & $analyzer -Folder $Folder -OutputObject
if ($copyReview.OverallStatus -eq 'Blocked') {
    Write-Host ''
    Write-Host 'ATTENTION NEEDED - NO LIVE CHANGES MADE'
    Write-Host 'The test review found a problem. The live workbook was not updated.'
    Write-Host "Review details: $($copyReview.MarkdownPath)"
    exit 1
}

Write-Host ''
Write-Host 'STEP 3 OF 3 - Choose whether to update the live workbook'
Write-Host 'The copy-test passed. Type APPLY and press Enter to continue.'
Write-Host 'Press Enter without typing APPLY to stop safely.'
$confirmation = Read-Host 'Confirmation'

if ($confirmation -ne 'APPLY') {
    Write-Host ''
    Write-Host 'NO LIVE CHANGES MADE'
    Write-Host 'The copy-test finished, but the live workbook was not updated.'
    exit 0
}

Write-Host ''
Write-Host 'Updating the live workbook now. A backup will be saved first.'
& $updater -Folder $Folder -Apply

Write-Host 'Checking the completed live update.'
$liveReview = & $analyzer -Folder $Folder -OutputObject
if ($liveReview.OverallStatus -eq 'Blocked') {
    Write-Host ''
    Write-Host 'ATTENTION NEEDED'
    Write-Host 'The live update ran, but the final review found a problem.'
    Write-Host "Review details: $($liveReview.MarkdownPath)"
    exit 1
}

Write-Host ''
Write-Host 'UPDATE COMPLETE'
Write-Host 'The live workbook was updated and the final review passed.'
Write-Host "Review details: $($liveReview.MarkdownPath)"
