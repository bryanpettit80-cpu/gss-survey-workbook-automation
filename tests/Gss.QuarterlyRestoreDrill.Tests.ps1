Describe 'GSS quarterly restore drill' {
    It 'writes a combined failed receipt when the snapshot has no main workbook mapping' {
        $repoRoot = Split-Path -Parent $PSScriptRoot
        $quarterlyDrillScript = Join-Path $repoRoot 'scripts\Invoke-GSS-QuarterlyRestoreDrill.ps1'
        $driveModule = Join-Path $repoRoot 'scripts\Gss-DriveBackup.ps1'
        $sandboxRoot = Join-Path $TestDrive 'quarterly-drill'
        $sandboxScripts = Join-Path $sandboxRoot 'scripts'
        $restoreDestination = Join-Path $sandboxRoot 'restore'
        New-Item -ItemType Directory -Path $sandboxScripts, $restoreDestination -Force | Out-Null
        Copy-Item -LiteralPath $quarterlyDrillScript -Destination (Join-Path $sandboxScripts 'Invoke-GSS-QuarterlyRestoreDrill.ps1')
        Copy-Item -LiteralPath $driveModule -Destination (Join-Path $sandboxScripts 'Gss-DriveBackup.ps1')

        @'
[pscustomobject]@{
    ReleaseTag = 'v-test'
    HeadCommit = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
}
'@ | Set-Content -LiteralPath (Join-Path $sandboxScripts 'Test-GSS-ReleaseIntegrity.ps1') -Encoding UTF8

        @'
[CmdletBinding()]
param(
    [string]$Operation,
    [string]$RunId,
    [switch]$OutputObject,
    [string]$SettingsPath
)
$destination = Join-Path (Split-Path -Parent $PSScriptRoot) 'restore'
$receiptPath = Join-Path $destination 'restore-verification.json'
[ordered]@{
    schema_version = 1
    operation = 'verify_only_restore'
    status = 'Verified'
    run_id = $RunId
    destination = $destination
    live_workbook_overwritten = $false
    verification_level = 'drivefs_hash_verified'
    file_count = 0
    files = @()
} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $receiptPath -Encoding UTF8
[pscustomobject]@{
    Status = 'Verified'
    LiveWorkbookOverwritten = $false
    ReceiptPath = $receiptPath
    Destination = $destination
    VerificationLevel = 'drivefs_hash_verified'
}
'@ | Set-Content -LiteralPath (Join-Path $sandboxScripts 'Invoke-GSS-DriveBackup.ps1') -Encoding UTF8

        @'
$sentinel = Join-Path (Split-Path -Parent $PSScriptRoot) 'excel-invoked.txt'
'invoked' | Set-Content -LiteralPath $sentinel -Encoding UTF8
throw 'Excel validation must not run when the workbook mapping is absent.'
'@ | Set-Content -LiteralPath (Join-Path $sandboxScripts 'Test-GSS-WorkbookIntegration.ps1') -Encoding UTF8

        $sandboxDrill = Join-Path $sandboxScripts 'Invoke-GSS-QuarterlyRestoreDrill.ps1'
        { & $sandboxDrill -RunId 'recovery-only-test' } |
            Should -Throw '*must contain exactly one mapping*found 0*'

        $drillReceiptPath = Join-Path $restoreDestination 'quarterly-restore-drill.json'
        $restoreReceiptPath = Join-Path $restoreDestination 'restore-verification.json'
        $drillReceiptPath | Should -Exist
        (Join-Path $sandboxRoot 'excel-invoked.txt') | Should -Not -Exist
        $receipt = Get-Content -LiteralPath $drillReceiptPath -Raw | ConvertFrom-Json
        $receipt.schema_version | Should -Be 1
        $receipt.operation | Should -Be 'quarterly_verify_only_restore_drill'
        $receipt.status | Should -Be 'Failed'
        $receipt.run_id | Should -Be 'recovery-only-test'
        $receipt.drive_restore_receipt | Should -Be $restoreReceiptPath
        $receipt.drive_verification_level | Should -Be 'drivefs_hash_verified'
        $receipt.restored_destination | Should -Be $restoreDestination
        $receipt.restored_workbook_portable_path | Should -Be 'gss/01 Main Workbook/GSS Score Trends - Main.xlsx'
        $receipt.restored_workbook_path | Should -BeNullOrEmpty
        $receipt.restored_workbook_sha256 | Should -BeNullOrEmpty
        $receipt.excel_validation_receipt | Should -BeNullOrEmpty
        $receipt.live_workbook_overwritten | Should -BeFalse
        $receipt.error | Should -Be "Restore receipt must contain exactly one mapping for portable path 'gss/01 Main Workbook/GSS Score Trends - Main.xlsx'; found 0."
    }
}
