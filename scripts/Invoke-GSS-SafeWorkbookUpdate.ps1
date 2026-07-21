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
$copyRun = & $updater -Folder $Folder -OutputObject
if (-not $copyRun -or [string]::IsNullOrWhiteSpace([string]$copyRun.LogPath)) {
    throw 'The copy-test updater did not return its exact run log path.'
}

Write-Host ''
Write-Host 'STEP 2 OF 3 - Checking the test results'
$copyReview = & $analyzer -Folder $Folder -LogPath $copyRun.LogPath -OutputObject
if ($copyReview.WorkbookStatus -eq 'Blocked') {
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
$liveRun = & $updater -Folder $Folder -Apply -OutputObject
if (-not $liveRun -or [string]::IsNullOrWhiteSpace([string]$liveRun.LogPath)) {
    throw 'The live updater did not return its exact run log path.'
}

Write-Host 'Checking the completed live update.'
$liveReview = & $analyzer -Folder $Folder -LogPath $liveRun.LogPath -OutputObject -PublishEmailPackage
if ($liveReview.WorkbookStatus -eq 'Blocked') {
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
if ($liveReview.EmailReadiness -eq 'Ready' -and $liveReview.EmailPackage) {
    Write-Host "Email package ready: $($liveReview.EmailPackage.PackagePath)"
}
else {
    Write-Host 'EMAIL PACKAGE NOT READY'
    Write-Host 'The workbook update succeeded, but drafting is blocked until the email-readiness issue is resolved.'
    foreach ($blocker in @($liveReview.Qa.EmailBlockers)) { Write-Host "  - $blocker" }
}
