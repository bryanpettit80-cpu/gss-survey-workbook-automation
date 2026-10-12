[CmdletBinding()]
param(
    [string]$Folder,
    [string]$MainWorkbookName = 'GSS Score Trends - Main.xlsx',
    [string]$ReportingWeekEnding,
    [string[]]$CorrectionSourcePath = @(),
    [string]$BackupFolder,
    [switch]$NonInteractive,
    [switch]$OutputObject
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
. (Join-Path $scriptRoot 'Gss-Operations.ps1')
$runtime = Get-GssRuntimeInfo -RepoRoot $repoRoot
if ([string]::IsNullOrWhiteSpace($Folder)) { $Folder = Split-Path -Parent $repoRoot }
$Folder = (Resolve-Path -LiteralPath $Folder.Trim('"')).Path
$updater = Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1'
$analyzer = Join-Path $scriptRoot 'Analyze-GSS-Run.ps1'
$stateDirectory = Join-Path $Folder '_automation_runs\state'
$activeRunPath = Join-Path $stateDirectory 'active-transaction.json'
$runId = [guid]::NewGuid().ToString('D')
$receiptPath = Join-Path $stateDirectory "operations\transaction-$runId.json"
$receipt = [ordered]@{
    ReceiptSchemaVersion = 2
    RunId = $runId
    HostName = [Environment]::MachineName
    ProgramRelease = $runtime.ReleaseTag
    StartedUtc = [datetime]::UtcNow.ToString('o')
    UpdatedUtc = [datetime]::UtcNow.ToString('o')
    TransactionStatus = 'Preparing'
    ReportStatus = 'Preparing'
    BackupStatus = 'NotStarted'
    PackageStatus = 'NotStarted'
    DraftEvidenceBackupStatus = 'NotCreated'
    PackagePublished = $false
    AutomaticSending = 'Disabled'
    ReceiptPath = $receiptPath
}
$mutex = New-Object Threading.Mutex($false, 'Global\GSSSurveyWorkbookAutomationTransaction')
$ownsMutex = $false
$ownsMarker = $false
$liveRun = $null
$reportValidated = $false

function Save-GssLauncherReceipt {
    $receipt.ReportBackupStatus = [string]$receipt.BackupStatus
    $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
    Write-GssOperationJson -Path $receiptPath -Value $receipt
    if ($ownsMarker) { Write-GssOperationJson -Path $activeRunPath -Value $receipt }
}
function Invoke-GssLauncherAnalysis {
    param([object]$Run, [switch]$Publish)
    $arguments = @{ Folder = $Folder; LogPath = [string]$Run.LogPath; OutputObject = $true }
    if ($Publish) { $arguments.PublishEmailPackage = $true }
    & $analyzer @arguments
}
function Select-GssCorrectionSource {
    param([object]$Failure)
    if ($NonInteractive -or [string]$Failure.Exception.Data['GssChoice'] -ne 'CorrectionSource') { throw $Failure }
    $candidates = @($Failure.Exception.Data['Candidates'])
    if ($candidates.Count -eq 0) { throw $Failure }
    Write-Information ([string]$Failure.Exception.Data['Message']) -InformationAction Continue
    for ($index = 0; $index -lt $candidates.Count; $index++) {
        Write-Information ("{0}. {1} - week ending {2}" -f ($index + 1), $candidates[$index].Path, $candidates[$index].WeekEnding) -InformationAction Continue
    }
    $answer = Read-Host 'Choose the source revision to use (number), or press Enter to leave the report unchanged'
    $choice = 0
    if (-not [int]::TryParse($answer, [ref]$choice) -or $choice -lt 1 -or $choice -gt $candidates.Count) { throw 'Source correction was not selected. No report changes were made.' }
    return [string]$candidates[$choice - 1].Path
}
function Update-GssDetailAssignment {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'The explicit file/week and supersession choices below are the operator decision; a second generic confirmation adds no protection.')]
    [CmdletBinding()]
    param([object]$Review)
    if ($NonInteractive -or -not $Review.DetailReadiness -or [string]$Review.DetailReadiness.Status -notin @('RequiresAssignment', 'Conflict', 'Missing')) { return $false }
    $candidates = @($Review.DetailReadiness.Candidates)
    if ($candidates.Count -eq 0) { return $false }
    Write-Information "The report is ready. Select the guest-detail export for report week $($liveRun.CurrentWeekEnding) to prepare its draft package." -InformationAction Continue
    for ($index = 0; $index -lt $candidates.Count; $index++) {
        Write-Information ("{0}. {1} ({2} responses; visits {3} through {4})" -f ($index + 1), $candidates[$index].RelativePath, $candidates[$index].ResponseCount, $candidates[$index].VisitDateStart, $candidates[$index].VisitDateEnd) -InformationAction Continue
    }
    $answer = Read-Host "Enter the file number to assign to $($liveRun.CurrentWeekEnding), or press Enter to prepare the package later"
    $choice = 0
    if (-not [int]::TryParse($answer, [ref]$choice) -or $choice -lt 1 -or $choice -gt $candidates.Count) { return $false }
    . (Join-Path $scriptRoot 'Gss-EmailPackage.ps1')
    $selected = $candidates[$choice - 1]
    $supersedes = @()
    if ([string]$Review.DetailReadiness.Status -in @('Conflict', 'Missing')) {
        $week = [datetime]::ParseExact([string]$liveRun.CurrentWeekEnding, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $assignments = @(Read-GssDetailReportingAssignment -FolderPath $Folder) + @(Get-GssPackagedDetailAssignment -FolderPath $Folder -ReportingDate $week)
        $previousChoices = @($candidates | Where-Object ContentSha256 -ne $selected.ContentSha256 | ForEach-Object {
            [pscustomobject]@{ ContentSha256 = $_.ContentSha256; Label = $_.RelativePath }
        })
        foreach ($assignment in @($assignments | Where-Object { $_.reporting_date -eq $liveRun.CurrentWeekEnding -and $_.content_sha256 -ne $selected.ContentSha256 })) {
            if ([string]$assignment.content_sha256 -notin @($previousChoices.ContentSha256)) {
                $previousChoices += [pscustomobject]@{ ContentSha256 = $assignment.content_sha256; Label = 'Previously assigned export (file unavailable)' }
            }
        }
        if ($previousChoices.Count -eq 0) { return $false }
        Write-Information 'Choose the old export version(s) this correction replaces. Files remain archived.' -InformationAction Continue
        for ($index = 0; $index -lt $previousChoices.Count; $index++) {
            Write-Information ("{0}. {1} [{2}]" -f ($index + 1), $previousChoices[$index].Label, ([string]$previousChoices[$index].ContentSha256).Substring(0, 8)) -InformationAction Continue
        }
        $answer = Read-Host 'Enter predecessor number(s), separated by commas, or press Enter to leave this conflict unresolved'
        if ([string]::IsNullOrWhiteSpace($answer)) { return $false }
        foreach ($token in $answer.Split(',')) {
            $previous = 0
            if (-not [int]::TryParse($token.Trim(), [ref]$previous) -or $previous -lt 1 -or $previous -gt $previousChoices.Count) { throw 'The correction predecessor selection was invalid; no assignment was changed.' }
            $supersedes += [string]$previousChoices[$previous - 1].ContentSha256
        }
    }
    Set-GssDetailReportingAssignment -FolderPath $Folder -DetailPath ([string]$selected.Path) -ReportingDate ([string]$liveRun.CurrentWeekEnding) -SupersedesContentSha256 @($supersedes | Sort-Object -Unique) | Out-Null
    return $true
}

try {
    try { $ownsMutex = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $ownsMutex = $true }
    if (-not $ownsMutex) { throw 'Another GSS update is running on this computer.' }
    if (Test-Path -LiteralPath $activeRunPath -PathType Leaf) {
        $active = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
        $allowed = @('Committed', 'AlreadyCurrent', 'PackageBlocked', 'PendingFinalize', 'BackupCommitted', 'BackupBlocked', 'Aborted', 'Blocked')
        $legacyRecovery = [int]$active.ReceiptSchemaVersion -lt 2 -and [string]$active.TransactionStatus -notin @('Committed', 'PackageBlocked', 'Aborted', 'Blocked')
        if ($legacyRecovery -or [string]$active.TransactionStatus -notin $allowed) {
            $resumeScript = Join-Path $scriptRoot 'Resume-GSS-PendingFinalize.ps1'
            $backupOnly = [string]$active.TransactionStatus -in @('ReportReady', 'LocalSnapshotPending', 'PendingFinalize', 'BackupCommitted', 'BackupBlocked', 'PackageBlocked')
            $resumeCommand = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -Folder "{1}" -RunId "{2}"{3}' -f $resumeScript, $Folder, $active.RunId, $(if ($backupOnly) { '' } else { ' -RestorePreviousReport' })
            $recoveryReason = "Run $($active.RunId) is $($active.TransactionStatus). Recovery command: $resumeCommand"
            if ($NonInteractive) { throw $recoveryReason }
            Write-Information $recoveryReason -InformationAction Continue
            $word = if ($backupOnly) { 'RETRY' } else { 'RESTORE' }
            $action = if ($backupOnly) { 'retry the pending backup and package' } else { 'restore the previous good report without overwriting unrelated changes' }
            $answer = Read-Host "Type $word to $action, or press Enter to stop"
            if ($answer -cne $word) { throw "Recovery was not selected. $recoveryReason" }
            $resumed = & $resumeScript -Folder $Folder -RunId ([string]$active.RunId) -RestorePreviousReport:(-not $backupOnly) -BackupFolder $BackupFolder -OutputObject
            if ($resumed.ReportStatus -notin @('Ready', 'PreviousReportRestored', 'PreviousReportUnchanged')) { throw "Recovery did not establish a safe report state. $recoveryReason" }
            if ([string]$resumed.TransactionStatus -eq 'LocalSnapshotPending') { throw "The report still needs a verified local recovery snapshot. $recoveryReason" }
        }
        # Backup-only pending work is independently retryable and never a write lock.
        if (Test-Path -LiteralPath $activeRunPath) { Remove-Item -LiteralPath $activeRunPath -Force }
    }
    $ownsMarker = $true
    Save-GssLauncherReceipt
    Write-Information 'Preparing and checking the report.' -InformationAction Continue
    while ($true) {
        try {
            $arguments = @{ Folder = $Folder; MainWorkbookName = $MainWorkbookName; RunId = $runId; ProgramRelease = $runtime.ReleaseTag; MutexAlreadyHeld = $true; OutputObject = $true }
            if ($ReportingWeekEnding) { $arguments.ReportingWeekEnding = $ReportingWeekEnding }
            if ($CorrectionSourcePath.Count -gt 0) { $arguments.CorrectionSourcePath = $CorrectionSourcePath }
            $copyRun = & $updater @arguments
            break
        }
        catch { $CorrectionSourcePath += Select-GssCorrectionSource -Failure $_ }
    }
    if (-not $copyRun -or -not $copyRun.LogPath) { throw 'The updater did not return its prepared report.' }
    $receipt.CopyRunLogPath = Get-GssOperationRelativePath -Root $Folder -Path $copyRun.LogPath
    $receipt.RunFingerprint = $copyRun.RunFingerprint
    $receipt.ReportingWeekEnding = $copyRun.CurrentWeekEnding
    $copyReview = Invoke-GssLauncherAnalysis -Run $copyRun
    if ($copyReview.WorkbookStatus -eq 'Blocked') { throw "Report validation failed; the previous report is unchanged. $($copyReview.MarkdownPath)" }
    $existingReport = Find-GssExistingReport -PreparedRun $copyRun -Folder $Folder
    if ($existingReport) {
        $receipt.TransactionStatus = 'AlreadyCurrent'
        $receipt.ReportStatus = 'AlreadyCurrent'
        $receipt.ReportRunId = [string]$existingReport.RunId
        $receipt.LiveRunLogPath = Get-GssOperationRelativePath -Root $Folder -Path $existingReport.LogPath
        $receipt.WorkbookPath = $existingReport.TargetWorkbookRelativePath
        $receipt.PdfPath = $existingReport.EmailComparisonPdfRelativePath
        $reportValidated = $true
        $liveRun = $existingReport
        if ($copyRun.AcceptedSourceChoices) {
            $choiceReceiptPath = & {
                param($UpdaterPath, $Prepared, $Existing, $DataFolder)
                . $UpdaterPath
                Write-GssAcceptedSourceChoiceReceipt -PreparedRun $Prepared -ExistingRun $Existing -FolderPath $DataFolder -LogDirectory (Join-Path $DataFolder '_automation_runs\logs')
            } $updater $copyRun $existingReport $Folder
            if ($choiceReceiptPath) { $receipt.AcceptedSourceChoiceLogPath = Get-GssOperationRelativePath -Root $Folder -Path $choiceReceiptPath }
        }
        try { [void](Update-GssDetailAssignment -Review $copyReview) } catch { $receipt.PackageError = $_.Exception.Message }
        # The new check has made no live write, so it must not hold a marker
        # against retrying the existing report's independent backup/package.
        if (Test-Path -LiteralPath $activeRunPath) { Remove-Item -LiteralPath $activeRunPath -Force }
        $ownsMarker = $false
        try {
            $resumed = & (Join-Path $scriptRoot 'Resume-GSS-PendingFinalize.ps1') -Folder $Folder -LiveRunLogPath $existingReport.LogPath -FreshPreparedRunLogPath $copyRun.LogPath -BackupFolder $BackupFolder -OutputObject
            $receipt.BackupStatus = $resumed.BackupStatus
            $receipt.PackageStatus = $resumed.PackageStatus
            $receipt.PackagePublished = [bool]$resumed.PackagePublished
            $receipt.DraftEvidenceBackupStatus = [string]$resumed.DraftEvidenceBackupStatus
            if ($resumed.PackagePath) { $receipt.PackagePath = $resumed.PackagePath }
        }
        catch { $receipt.BackupStatus = 'Pending'; $receipt.PackageStatus = 'WaitingForRecovery'; $receipt.Error = $_.Exception.Message }
        $receipt.CompletedUtc = [datetime]::UtcNow.ToString('o')
        Save-GssLauncherReceipt
        Write-Information "REPORT ALREADY CURRENT - week ending $($existingReport.CurrentWeekEnding). No report files were replaced. Backup: $($receipt.BackupStatus). Draft package: $($receipt.PackageStatus)." -InformationAction Continue
        if ($OutputObject) { [pscustomobject]$receipt }
        return
    }
    $receipt.StartingWorkbookRelativePath = if ($copyRun.StartingWorkbookRelativePath) { [string]$copyRun.StartingWorkbookRelativePath } else { Get-GssOperationRelativePath -Root $Folder -Path (Join-Path $Folder "01 Main Workbook\$MainWorkbookName") }
    $receipt.StartingWorkbookSha256 = [string]$copyRun.StartingWorkbookSha256
    $startingPdfPath = Join-Path $Folder ('04 Email Comparison PDFs\' + [IO.Path]::GetFileName([string]$copyRun.EmailComparisonPdf))
    $receipt.StartingPdfRelativePath = Get-GssOperationRelativePath -Root $Folder -Path $startingPdfPath
    $receipt.StartingPdfExisted = Test-Path -LiteralPath $startingPdfPath -PathType Leaf
    $receipt.StartingPdfSha256 = if ($receipt.StartingPdfExisted) { Get-GssOperationHash $startingPdfPath } else { $null }
    $receipt.TransactionStatus = 'Promoting'
    Save-GssLauncherReceipt
    $liveRun = & $updater -Folder $Folder -MainWorkbookName $MainWorkbookName -Apply -PreparedRunLogPath $copyRun.LogPath -ExpectedFingerprint $copyRun.RunFingerprint -ProgramRelease $runtime.ReleaseTag -MutexAlreadyHeld -OutputObject
    if (-not $liveRun -or -not $liveRun.LogPath) { throw 'The updater did not return its committed transaction log.' }
    $receipt.LiveRunLogPath = Get-GssOperationRelativePath -Root $Folder -Path $liveRun.LogPath
    $receipt.TransactionStatus = 'Validating'
    Save-GssLauncherReceipt
    $liveReview = Invoke-GssLauncherAnalysis -Run $liveRun
    if ($liveReview.WorkbookStatus -eq 'Blocked') { throw "The completed report failed validation. $($liveReview.MarkdownPath)" }
    $reportValidated = $true
    $receipt.ReportStatus = 'Ready'
    $receipt.TransactionStatus = 'ReportReady'
    $receipt.ReviewPath = Get-GssOperationRelativePath -Root $Folder -Path $liveReview.MarkdownPath
    $receipt.WorkbookPath = $liveRun.TargetWorkbookRelativePath
    $receipt.PdfPath = $liveRun.EmailComparisonPdfRelativePath
    Save-GssLauncherReceipt

    try {
        if (Update-GssDetailAssignment -Review $liveReview) { $liveReview = Invoke-GssLauncherAnalysis -Run $liveRun }
    }
    catch { $receipt.PackageError = $_.Exception.Message }

    try {
        $snapshotPath = New-GssRunSnapshot -Run $liveRun -Folder $Folder
        $receipt.LocalSnapshotPath = Get-GssOperationRelativePath -Root $Folder -Path $snapshotPath
        $receipt.LocalSnapshotStatus = 'Verified'
    }
    catch {
        $receipt.TransactionStatus = 'LocalSnapshotPending'
        $receipt.LocalSnapshotStatus = 'Failed'
        $receipt.BackupStatus = 'Pending'
        $receipt.BackupError = $_.Exception.Message
    }
    if ($receipt.LocalSnapshotStatus -eq 'Verified') {
        try {
            $backup = Sync-GssRunSnapshot -ManifestPath $snapshotPath -BackupFolder $BackupFolder
            $receipt.BackupStatus = $backup.BackupStatus
            $receipt.Backup = $backup
        }
        catch { $receipt.BackupStatus = 'Pending'; $receipt.BackupError = $_.Exception.Message }
    }
    Save-GssLauncherReceipt

    try {
        if ($receipt.BackupStatus -eq 'Committed') {
            $published = Invoke-GssLauncherAnalysis -Run $liveRun -Publish
            $receipt.PackagePublished = [bool]($published.EmailReadiness -eq 'Ready' -and $published.EmailPackage)
            $receipt.PackageStatus = if ($receipt.PackagePublished) { 'Ready' } elseif ($published.EmailReadiness -eq 'WaitingForInput') { 'WaitingForInput' } else { 'Blocked' }
            $receipt.PackagePath = if ($receipt.PackagePublished) { Get-GssOperationRelativePath -Root $Folder -Path $published.EmailPackage.PackagePath } else { $null }
            if ($receipt.PackagePublished) { $receipt.DraftEvidenceBackupStatus = 'LocalOnly' }
            if (-not $receipt.PackagePublished) { $receipt.PackageBlockers = @($published.Qa.EmailBlockers) }
        }
        else { $receipt.PackageStatus = 'WaitingForBackup' }
    }
    catch { $receipt.PackageStatus = 'Blocked'; $receipt.PackageError = $_.Exception.Message }
    if ($receipt.LocalSnapshotStatus -eq 'Verified') { $receipt.TransactionStatus = 'Committed' }
    $receipt.CompletedUtc = [datetime]::UtcNow.ToString('o')
    Save-GssLauncherReceipt
    Write-Information "REPORT READY - week ending $($liveRun.CurrentWeekEnding). Backup: $($receipt.BackupStatus). Draft package: $($receipt.PackageStatus)." -InformationAction Continue
    Write-Information "Review: $($liveReview.MarkdownPath)" -InformationAction Continue
}
catch {
    $failure = $_
    $receipt.Error = $failure.Exception.Message
    if ($liveRun -and -not $reportValidated) {
        try {
            $receipt.TransactionStatus = 'RollbackAttempting'
            Save-GssLauncherReceipt
            $rollback = & $updater -Folder $Folder -MainWorkbookName $MainWorkbookName -RollbackRunLogPath $liveRun.LogPath -ProgramRelease $runtime.ReleaseTag -MutexAlreadyHeld -OutputObject
            if ([string]$rollback.Status -ne 'RolledBack') { throw 'Rollback did not verify the previous report.' }
            $receipt.TransactionStatus = 'Aborted'
            $receipt.ReportStatus = 'PreviousReportRestored'
            $receipt.Rollback = $rollback
        }
        catch { $receipt.TransactionStatus = 'RollbackBlocked'; $receipt.ReportStatus = 'RecoveryRequired'; $receipt.RollbackError = $_.Exception.Message }
    }
    elseif ($receipt.TransactionStatus -eq 'Promoting') {
        $receipt.TransactionStatus = 'RecoveryRequired'
        $receipt.ReportStatus = 'RecoveryRequired'
        try {
            $journals = @(Get-ChildItem -LiteralPath (Join-Path $Folder '_automation_runs\logs') -File -Filter "*_$runId`_apply.json")
            if ($journals.Count -eq 0) {
                $closed = Complete-GssUnstartedOperation -Folder $Folder -RunId $runId -ResumeScript (Join-Path $scriptRoot 'Resume-GSS-PendingFinalize.ps1')
                $receipt.TransactionStatus = $closed.TransactionStatus
                $receipt.ReportStatus = $closed.ReportStatus
                $receipt.Recovery = $closed.Recovery
            }
            elseif ($journals.Count -eq 1) {
                $failedRun = Get-Content -LiteralPath $journals[0].FullName -Raw | ConvertFrom-Json
                if ([string]$failedRun.TransactionStatus -in @('Aborted', 'Blocked')) {
                    $closed = & (Join-Path $scriptRoot 'Resume-GSS-PendingFinalize.ps1') -Folder $Folder -RunId $runId -RestorePreviousReport -OutputObject
                    $receipt.TransactionStatus = $closed.TransactionStatus
                    $receipt.ReportStatus = $closed.ReportStatus
                    $receipt.Recovery = $closed.Recovery
                }
            }
        }
        catch { $receipt.RecoveryError = $_.Exception.Message }
    }
    elseif (-not $reportValidated) { $receipt.TransactionStatus = 'Blocked'; $receipt.ReportStatus = 'Blocked' }
    if ($ownsMutex -and $ownsMarker) { Save-GssLauncherReceipt }
    throw $failure
}
finally {
    if ($ownsMarker -or $reportValidated) {
        try {
            Write-GssOperationSummary -Folder $Folder -Receipt ([pscustomobject]$receipt) -ResumeScript (Join-Path $scriptRoot 'Resume-GSS-PendingFinalize.ps1')
            if ($ownsMarker -and $receipt.TransactionStatus -notin @('RecoveryRequired', 'RollbackAttempting', 'RollbackBlocked', 'Promoting', 'Validating', 'ReportReady', 'LocalSnapshotPending')) {
                if (Test-Path -LiteralPath $activeRunPath) {
                    $active = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
                    if ([string]$active.RunId -eq $runId) { Remove-Item -LiteralPath $activeRunPath -Force }
                }
            }
        }
        catch { Write-Warning "Could not publish the final status: $($_.Exception.Message)" }
    }
    if ($ownsMutex) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
if ($OutputObject) { [pscustomobject]$receipt }
