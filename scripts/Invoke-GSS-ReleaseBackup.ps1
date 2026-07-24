[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Inventory', 'Prepare', 'Finalize', 'RetryFinalize', 'Abort', 'VerifyRestore')]
    [string]$Operation,
    [string]$Folder,
    [string]$SettingsPath,
    [switch]$OutputObject
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
$driveScript = Join-Path $scriptRoot 'Invoke-GSS-DriveBackup.ps1'
. (Join-Path $scriptRoot 'Gss-DriveBackup.ps1')

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $Folder = Split-Path -Parent $repoRoot
}
$gssRoot = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Folder).Path).TrimEnd('\', '/')
$expectedRepoRoot = [System.IO.Path]::GetFullPath((Join-Path $gssRoot 'GSS Survey Workbook Automation')).TrimEnd('\', '/')
$resolvedRepoRoot = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd('\', '/')
if (-not $resolvedRepoRoot.Equals($expectedRepoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "ReleaseOnly backup must run from the canonical program repository under the requested GSS root: $expectedRepoRoot"
}

$releaseIntegrity = & (Join-Path $scriptRoot 'Test-GSS-ReleaseIntegrity.ps1') -RepoRoot $resolvedRepoRoot
$manifestPath = Join-Path $resolvedRepoRoot 'release\release-manifest.json'
$manifest = Read-GssDriveBackupJson -Path $manifestPath
$version = [string](Get-GssDriveBackupProperty $manifest @('release_version'))
$tag = [string](Get-GssDriveBackupProperty $manifest @('release_tag'))
$archiveName = [string](Get-GssDriveBackupProperty $manifest @('archive_name'))
if ([string]$releaseIntegrity.ReleaseVersion -cne $version -or
    [string]$releaseIntegrity.ReleaseTag -cne $tag) {
    throw 'ReleaseOnly backup release-integrity evidence does not match the release manifest.'
}

$releaseDirectory = Join-Path $gssRoot '_automation_runs\state\release'
$archivePath = Join-Path $releaseDirectory $archiveName
$excelReceiptPath = Join-Path $releaseDirectory 'local-excel-validation-receipt.json'
$backupInventory = @(
    [pscustomobject][ordered]@{
        SourcePath = $archivePath
        PortablePath = "release/$archiveName"
        Role = 'release_archive'
        Classification = 'restricted_operational'
    }
    [pscustomobject][ordered]@{
        SourcePath = $manifestPath
        PortablePath = 'release/release-manifest.json'
        Role = 'release_manifest'
        Classification = 'restricted_operational'
    }
    [pscustomobject][ordered]@{
        SourcePath = $excelReceiptPath
        PortablePath = 'release/local-excel-validation-receipt.json'
        Role = 'release_excel_receipt'
        Classification = 'restricted_operational'
    }
)

$inventory = @(Get-GssDriveBackupInventory `
    -GssRoot $gssRoot `
    -AdditionalItems $backupInventory `
    -InventoryMode ReleaseOnly)
$fingerprintHash = Get-GssReleaseOnlyInventoryFingerprint -Inventory $inventory -Release $tag
$fingerprint = "sha256:$fingerprintHash"
$runId = "gss-release-$tag-$($fingerprintHash.Substring(0, 12))"

$generatedAt = [datetime]::MinValue
$dateStyles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
if (-not [datetime]::TryParse(
    [string](Get-GssDriveBackupProperty $manifest @('generated_at_utc')),
    [Globalization.CultureInfo]::InvariantCulture,
    $dateStyles,
    [ref]$generatedAt
)) {
    throw 'ReleaseOnly manifest generated_at_utc is invalid.'
}

$summaryPath = Join-Path $releaseDirectory ("drive-release-summary-$tag.json")
$summary = [ordered]@{
    schema_version = 'gss-release-drive-summary/v1'
    GeneratedAtUtc = $generatedAt.ToUniversalTime().ToString('o')
    RunId = $runId
    RunFingerprint = $fingerprint
    Folder = $gssRoot
    CurrentWeekEnding = $generatedAt.ToUniversalTime().Date.ToString('yyyy-MM-dd')
    ProgramRelease = $tag
    SnapshotPurpose = 'ReleaseOnly'
    BackupInventory = $backupInventory
}
Write-GssDriveBackupAtomicJson -Path $summaryPath -Value $summary

$arguments = @{
    Operation = $Operation
    RunSummaryPath = $summaryPath
    OutputObject = $true
}
if (-not [string]::IsNullOrWhiteSpace($SettingsPath)) {
    $arguments.SettingsPath = $SettingsPath
}
$result = & $driveScript @arguments
$result | Add-Member -NotePropertyName ReleaseSummaryPath -NotePropertyValue $summaryPath -Force
$result | Add-Member -NotePropertyName ReleaseTag -NotePropertyValue $tag -Force
$result | Add-Member -NotePropertyName ReleaseFingerprint -NotePropertyValue $fingerprint -Force

if ($OutputObject) {
    Write-Output $result
}
else {
    $result | Format-List
}
