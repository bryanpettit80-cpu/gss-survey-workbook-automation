[CmdletBinding()]
param(
    [string]$BackupFolder,
    [switch]$ClearBackup,
    [switch]$OutputObject
)
$ErrorActionPreference = 'Stop'
if ($ClearBackup -and $BackupFolder) { throw 'Choose BackupFolder or ClearBackup, not both.' }
if (-not $ClearBackup -and [string]::IsNullOrWhiteSpace($BackupFolder)) { throw 'Supply an existing private BackupFolder, or ClearBackup to keep local recovery only.' }
if ($BackupFolder) {
    if (-not (Test-Path -LiteralPath $BackupFolder -PathType Container)) { throw 'The backup folder must already exist and be available on this computer.' }
    $BackupFolder = (Resolve-Path -LiteralPath $BackupFolder).Path
}
. (Join-Path $PSScriptRoot 'Gss-Operations.ps1')
$path = Join-Path $env:LOCALAPPDATA 'GSSSurveyWorkbookAutomation\runtime-settings.json'
$settings = [ordered]@{ schema_version = 1; backup_folder = if ($ClearBackup) { $null } else { $BackupFolder }; updated_utc = [datetime]::UtcNow.ToString('o') }
Write-GssOperationJson -Path $path -Value $settings
$result = [pscustomobject]@{ Status = 'Configured'; SettingsPath = $path; BackupFolder = $settings.backup_folder; LocalRecoveryRequired = $true; CloudRequiredBeforeReport = $false; SchedulesChanged = $false }
if ($OutputObject) { $result } else { $result | Format-List }
