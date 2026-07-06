[CmdletBinding()]
param(
    [string]$Folder
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$projectRoot = Split-Path -Parent $scriptRoot

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $Folder = Split-Path -Parent $projectRoot
}

$Folder = (Resolve-Path -LiteralPath $Folder).Path
$template = Join-Path $projectRoot 'templates\Run GSS Update After Upload.cmd'
$operatorLauncher = Join-Path $Folder 'Run GSS Update After Upload.cmd'
$task = Get-ScheduledTask -TaskName 'GSS Survey Main Workbook Weekly Update' -ErrorAction SilentlyContinue
$templateHash = if (Test-Path -LiteralPath $template) { (Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash } else { $null }
$operatorHash = if (Test-Path -LiteralPath $operatorLauncher) { (Get-FileHash -LiteralPath $operatorLauncher -Algorithm SHA256).Hash } else { $null }

$status = [ordered]@{
    RepoRoot = $projectRoot
    DropboxFolder = $Folder
    ManualWorkflow = 'Enabled; use Run GSS Update After Upload.cmd after uploading the newest source workbook.'
    OperatorLauncher = $operatorLauncher
    OperatorLauncherExists = [bool](Test-Path -LiteralPath $operatorLauncher)
    OperatorLauncherMatchesTemplate = ($templateHash -and $operatorHash -and $templateHash -eq $operatorHash)
    ScheduledTaskName = 'GSS Survey Main Workbook Weekly Update'
    ScheduledTaskExists = [bool]$task
    ScheduledTaskState = if ($task) { [string]$task.State } else { $null }
    ScheduledTaskEnabled = if ($task) { [bool]($task.Settings.Enabled) } else { $false }
    ScheduledTaskAction = if ($task) { ($task.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join '; ' } else { $null }
}

[pscustomobject]$status | Format-List
