[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$RunId,
    [string]$SettingsPath
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
$releaseValidator = Join-Path $scriptRoot 'Test-GSS-ReleaseIntegrity.ps1'
$backupCommand = Join-Path $scriptRoot 'Invoke-GSS-DriveBackup.ps1'
$excelValidator = Join-Path $scriptRoot 'Test-GSS-WorkbookIntegration.ps1'
. (Join-Path $scriptRoot 'Gss-DriveBackup.ps1')

$release = & $releaseValidator -RepoRoot $repoRoot

$backupArguments = @{
    Operation = 'VerifyRestore'
    RunId = $RunId
    OutputObject = $true
}
if (-not [string]::IsNullOrWhiteSpace($SettingsPath)) {
    $backupArguments.SettingsPath = $SettingsPath
}

$restore = & $backupCommand @backupArguments
if (-not $restore -or [string]$restore.Status -ne 'Verified' -or [bool]$restore.LiveWorkbookOverwritten) {
    throw "Verify-only Drive restore did not complete safely for run '$RunId'."
}

$mainWorkbookPortablePath = 'gss/01 Main Workbook/GSS Score Trends - Main.xlsx'
$integrationReceipt = Join-Path ([string]$restore.Destination) 'local-excel-validation-receipt.json'
$drillReceipt = Join-Path ([string]$restore.Destination) 'quarterly-restore-drill.json'
$status = 'Failed'
$errorMessage = $null
$restoredWorkbookMapping = $null
$restoredWorkbook = $null

try {
    $restoredWorkbookMapping = Resolve-GssDriveBackupRestoredFile `
        -ReceiptPath ([string]$restore.ReceiptPath) `
        -PortablePath $mainWorkbookPortablePath `
        -ExpectedDestination ([string]$restore.Destination)
    $restoredWorkbook = [string]$restoredWorkbookMapping.Path
    $restoredGssRoot = if ([string]$restoredWorkbookMapping.RestoredPath -ilike 'gss/*') {
        Join-Path ([string]$restore.Destination) 'gss'
    }
    else {
        [string]$restore.Destination
    }

    if (-not (Test-Path -LiteralPath $restoredWorkbook -PathType Leaf)) {
        throw "Restored snapshot does not contain the main workbook at the expected portable path: $restoredWorkbook"
    }

    & $excelValidator `
        -WorkbookPath $restoredWorkbook `
        -ReceiptPath $integrationReceipt `
        -Folder $restoredGssRoot
    $integration = Read-GssDriveBackupJson -Path $integrationReceipt
    if ([string](Get-GssDriveBackupProperty $integration @('Status', 'status')) -ne 'Passed') {
        throw 'Restored workbook failed the desktop Excel integration validation.'
    }
    $status = 'Passed'
}
catch {
    $errorMessage = $_.Exception.Message
    throw
}
finally {
    $receipt = [ordered]@{
        schema_version = 1
        operation = 'quarterly_verify_only_restore_drill'
        status = $status
        run_id = $RunId
        completed_at_utc = [datetime]::UtcNow.ToString('o')
        release_tag = [string]$release.ReleaseTag
        release_commit = [string]$release.HeadCommit
        drive_restore_receipt = [string]$restore.ReceiptPath
        drive_verification_level = [string]$restore.VerificationLevel
        restored_destination = [string]$restore.Destination
        restored_workbook_portable_path = $mainWorkbookPortablePath
        restored_workbook_path = if ($null -ne $restoredWorkbookMapping) {
            [string]$restoredWorkbookMapping.RestoredPath
        }
        else { $null }
        restored_workbook_sha256 = if (
            -not [string]::IsNullOrWhiteSpace($restoredWorkbook) -and
            (Test-Path -LiteralPath $restoredWorkbook -PathType Leaf)
        ) {
            Get-GssDriveBackupSha256 -Path $restoredWorkbook
        }
        else { $null }
        excel_validation_receipt = if (Test-Path -LiteralPath $integrationReceipt -PathType Leaf) {
            $integrationReceipt
        }
        else { $null }
        live_workbook_overwritten = $false
        error = $errorMessage
    }
    Write-GssDriveBackupAtomicJson -Path $drillReceipt -Value $receipt
}

Write-Output ([pscustomobject]@{
    Status = $status
    RunId = $RunId
    RestoreDestination = [string]$restore.Destination
    RestoredWorkbook = $restoredWorkbook
    DriveRestoreReceipt = [string]$restore.ReceiptPath
    ExcelValidationReceipt = $integrationReceipt
    DrillReceipt = $drillReceipt
    LiveWorkbookOverwritten = $false
})
