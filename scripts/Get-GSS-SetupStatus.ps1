[CmdletBinding()]
param([string]$Folder, [switch]$OutputObject)
$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
. (Join-Path $scriptRoot 'Gss-Operations.ps1')
. (Join-Path $scriptRoot 'Gss-Common.ps1')
if (-not $Folder) { $Folder = Split-Path -Parent $repoRoot }
$Folder = (Resolve-Path -LiteralPath $Folder).Path
$runtime = Get-GssRuntimeInfo -RepoRoot $repoRoot
$backupError = $null
$backupRoot = $null
try { $backupRoot = Get-GssBackupFolder } catch { $backupError = $_.Exception.Message }
$excelAvailable = $false
try { $excelAvailable = $null -ne [type]::GetTypeFromProgID('Excel.Application') } catch { $excelAvailable = $false }
$task = $null
if (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue) {
    $task = Get-ScheduledTask -TaskName 'GSS Survey Main Workbook Weekly Update' -ErrorAction SilentlyContinue
}
$operator = Join-Path $Folder 'Run GSS Update After Upload.cmd'
$template = Join-Path $repoRoot 'templates\Run GSS Update After Upload.cmd'
$launcherMatches = (Test-Path -LiteralPath $operator) -and (Test-Path -LiteralPath $template) -and ((Get-GssOperationHash $operator) -ceq (Get-GssOperationHash $template))
$mainWorkbook = $null
$mainError = $null
try { $mainWorkbook = Resolve-MainWorkbookPath -FolderPath $Folder -WorkbookName 'GSS Score Trends - Main.xlsx' } catch { $mainError = $_.Exception.Message }
$pending = @()
$operations = Join-Path $Folder '_automation_runs\state\operations'
if (Test-Path -LiteralPath $operations -PathType Container) {
    foreach ($file in Get-ChildItem -LiteralPath $operations -File -Filter 'transaction-*.json') {
        try {
            $receipt = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
            if ($receipt.ReportStatus -eq 'Ready' -and $receipt.BackupStatus -ne 'Committed') { $pending += [string]$receipt.RunId }
        }
        catch { $pending += "Unreadable receipt: $($file.Name)" }
    }
}
$activePath = Join-Path $Folder '_automation_runs\state\active-transaction.json'
$active = if (Test-Path -LiteralPath $activePath -PathType Leaf) { Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json } else { $null }
$status = [pscustomobject]@{
    ProgramVersion = $runtime.ReleaseTag
    ProgramFolder = $repoRoot
    DataFolder = $Folder
    ExcelAvailable = $excelAvailable
    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    MainWorkbook = $mainWorkbook
    MainWorkbookError = $mainError
    OperatorLauncher = $operator
    OperatorLauncherExists = [bool](Test-Path -LiteralPath $operator)
    OperatorLauncherMatchesTemplate = [bool]$launcherMatches
    OrdinaryConfirmationRequired = $false
    GitRequiredForReporting = $false
    HistoricalReleaseReceiptRequired = $false
    LocalRecoveryCopyRequired = $true
    BackupFolder = $backupRoot
    BackupConfigured = -not [string]::IsNullOrWhiteSpace($backupRoot)
    BackupAvailable = [bool]($backupRoot -and (Test-Path -LiteralPath $backupRoot -PathType Container))
    BackupError = $backupError
    CloudRequiredBeforeReport = $false
    CloudUploadConfirmed = $false
    BackupVerification = 'Hashes verified through the local filesystem; cloud synchronization is not asserted.'
    PendingBackupRunIds = $pending
    ActiveTransaction = if ($active) { [string]$active.RunId } else { $null }
    ActiveTransactionStatus = if ($active) { [string]$active.TransactionStatus } else { $null }
    ScheduledTaskExists = [bool]$task
    ScheduledTaskState = if ($task) { [string]$task.State } else { $null }
    ScheduledTaskEnabled = [bool]($task -and $task.Settings.Enabled)
    SchedulesChanged = $false
    AutomaticSending = 'Disabled'
}
if ($OutputObject) { $status } else { $status | Format-List }
