[CmdletBinding()]
param(
    [string]$LiveRunLogPath,
    [string]$FreshPreparedRunLogPath,
    [string]$RunId,
    [string]$Folder,
    [string]$BackupFolder,
    [switch]$RestorePreviousReport,
    [switch]$OutputObject
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
. (Join-Path $scriptRoot 'Gss-Operations.ps1')
if ([string]::IsNullOrWhiteSpace($Folder)) { $Folder = Split-Path -Parent $repoRoot }
$Folder = (Resolve-Path -LiteralPath $Folder).Path
$stateRoot = Join-Path $Folder '_automation_runs\state'
$logsRoot = Join-Path $Folder '_automation_runs\logs'
if ($RunId -and $RunId -notmatch '^[a-fA-F0-9-]{36}$') { throw 'RunId must be a GUID.' }
if (-not $LiveRunLogPath) {
    if (-not $RunId) { throw 'Supply RunId or LiveRunLogPath and the GSS data Folder.' }
    $runLogMatches = @(if (Test-Path -LiteralPath $logsRoot -PathType Container) { Get-ChildItem -LiteralPath $logsRoot -File -Filter "*_$RunId`_apply.json" })
    if (@($runLogMatches).Count -eq 0) {
        $unstarted = Complete-GssUnstartedOperation -Folder $Folder -RunId $RunId -ResumeScript $PSCommandPath
        Write-Information 'The interrupted check stopped before report replacement. The existing report is unchanged; launch a fresh update.' -InformationAction Continue
        if ($OutputObject) { $unstarted }
        return
    }
    if ($runLogMatches.Count -ne 1) { throw "Expected one apply log for run $RunId; found $($runLogMatches.Count)." }
    $LiveRunLogPath = $runLogMatches[0].FullName
}
elseif (-not [IO.Path]::IsPathRooted($LiveRunLogPath)) {
    $LiveRunLogPath = Resolve-GssOperationPath -Root $Folder -RelativePath $LiveRunLogPath -RequireFile
}
$LiveRunLogPath = (Resolve-Path -LiteralPath $LiveRunLogPath).Path
$relativeLog = Get-GssOperationRelativePath -Root $Folder -Path $LiveRunLogPath
if (-not $relativeLog.StartsWith('_automation_runs/logs/', [StringComparison]::OrdinalIgnoreCase)) { throw 'Resume requires an apply log under the GSS log directory.' }
$originalRun = Get-Content -LiteralPath $LiveRunLogPath -Raw | ConvertFrom-Json
if ($RunId -and $RunId -cne [string]$originalRun.RunId) { throw 'RunId does not match the selected apply log.' }
$RunId = [string]$originalRun.RunId
if ($RunId -notmatch '^[a-fA-F0-9-]{36}$' -or [string]$originalRun.Mode -ne 'ApplyToMainWorkbook') { throw 'Resume requires an identified workbook apply transaction.' }
$run = ConvertTo-GssPortableRun -Run $originalRun -Folder $Folder
$run | Add-Member -NotePropertyName LogPath -NotePropertyValue $LiveRunLogPath -Force
$run | Add-Member -NotePropertyName LogRelativePath -NotePropertyValue $relativeLog -Force
$activePath = Join-Path $stateRoot 'active-transaction.json'
$receiptPath = Join-Path $stateRoot "operations\transaction-$RunId.json"
$receipt = if (Test-Path -LiteralPath $receiptPath) {
    Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
}
else {
    [pscustomobject]@{ ReceiptSchemaVersion = 2; RunId = $RunId; ProgramRelease = [string]$run.ProgramRelease; AutomaticSending = 'Disabled'; PackagePublished = $false }
}
function Set-GssResumeValue {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Updates an in-memory result object; persistence is handled by the enclosing recovery operation.')]
    [CmdletBinding()]
    param([string]$Name, [object]$Value)
    $receipt | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
}
function Save-GssResumeResult {
    Set-GssResumeValue 'ReportBackupStatus' ([string]$receipt.BackupStatus)
    Set-GssResumeValue 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
    Set-GssResumeValue 'ResumeHostName' ([Environment]::MachineName)
    Set-GssResumeValue 'LiveRunLogPath' $relativeLog
    Set-GssResumeValue 'ReportingWeekEnding' ([string]$run.CurrentWeekEnding)
    Write-GssOperationJson -Path $receiptPath -Value $receipt
    $latestPath = Join-Path $stateRoot 'latest-result.json'
    $latest = if (Test-Path -LiteralPath $latestPath -PathType Leaf) { Get-Content -LiteralPath $latestPath -Raw | ConvertFrom-Json } else { $null }
    if ($null -eq $latest -or [string]$latest.RunId -eq $RunId -or [string]$latest.ReportRunId -eq $RunId) {
        Write-GssOperationSummary -Folder $Folder -Receipt $receipt -ResumeScript $PSCommandPath
    }
}
$mutex = New-Object Threading.Mutex($false, 'Global\GSSSurveyWorkbookAutomationTransaction')
$ownsMutex = $false
try {
    try { $ownsMutex = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $ownsMutex = $true }
    if (-not $ownsMutex) { throw 'Another GSS update is running on this computer.' }
    foreach ($field in @('Error', 'BackupError', 'PackageError', 'PackagePath', 'PackageBlockers')) { Set-GssResumeValue $field $null }
    Set-GssResumeValue 'PackagePublished' $false
    Set-GssResumeValue 'DraftEvidenceBackupStatus' 'NotCreated'
    $active = $null
    if (Test-Path -LiteralPath $activePath) {
        $active = Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json
        if ([string]$active.RunId -ne $RunId -and [string]$active.TransactionStatus -notin @('Committed', 'PendingFinalize', 'PackageBlocked', 'BackupBlocked', 'Aborted', 'Blocked')) {
            throw "A different local update needs recovery first: $($active.RunId)"
        }
    }
    $activeRequiresRestore = $active -and [string]$active.RunId -eq $RunId -and [string]$active.TransactionStatus -in @('RecoveryRequired', 'RollbackAttempting', 'RollbackBlocked', 'Validating', 'Promoting')
    if ($RestorePreviousReport -or $activeRequiresRestore -or [string]$run.TransactionStatus -ne 'Committed') {
        if ([string]$run.TransactionStatus -notin @('Committed', 'Promoting', 'WorkbookPromoted', 'PdfPromoted', 'Aborted', 'Blocked', 'Conflict')) { throw "This transaction cannot be recovered automatically: $($run.TransactionStatus)" }
        if ([string]$run.TransactionStatus -eq 'Blocked' -and $run.PSObject.Properties['WorkbookWasPromoted'] -and $run.PSObject.Properties['PdfWasPromoted'] -and -not $run.WorkbookWasPromoted -and -not $run.PdfWasPromoted) {
            $recovery = [pscustomobject]@{ Status = 'NoLiveWriteObserved'; Reason = 'The durable failure receipt confirms that neither report file was replaced.' }
        }
        else {
            $recovery = & {
                param($UpdaterPath, $RunObject, $DataRoot)
                . $UpdaterPath
                Repair-GssInterruptedTransaction -Run $RunObject -FolderPath $DataRoot
            } (Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1') $run $Folder
        }
        if ([string]$recovery.Status -notin @('RolledBack', 'Unchanged', 'NoLiveWriteObserved')) { throw 'Recovery did not verify the previous report.' }
        Set-GssResumeValue 'TransactionStatus' 'Aborted'
        Set-GssResumeValue 'ReportStatus' $(if ($recovery.Status -eq 'NoLiveWriteObserved') { 'PreviousReportUnchanged' } else { 'PreviousReportRestored' })
        Set-GssResumeValue 'Recovery' $recovery
        Set-GssResumeValue 'PackageStatus' 'NotPublished'
        Save-GssResumeResult
        if (Test-Path -LiteralPath $activePath) {
            $active = Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json
            if ([string]$active.RunId -eq $RunId) { Invoke-GssFileOperationRetry { Remove-Item -LiteralPath $activePath -Force -ErrorAction Stop } }
        }
        $message = if ($recovery.Status -eq 'NoLiveWriteObserved') { 'The failed update did not replace the report. Launch the normal update to prepare a fresh report.' } else { 'The previous good report has been verified or restored. Launch the normal update to prepare a fresh report.' }
        Write-Information $message -InformationAction Continue
    }
    else {
        Set-GssResumeValue 'ReportStatus' 'Ready'
        $currentReportMatches = $true
        foreach ($check in @(@($run.TargetWorkbook, $run.PromotedWorkbookSha256), @($run.EmailComparisonPdf, $run.PromotedPdfSha256))) {
            if (-not (Test-Path -LiteralPath $check[0] -PathType Leaf) -or (Get-GssOperationHash $check[0]) -cne [string]$check[1]) { $currentReportMatches = $false; break }
        }
        Set-GssResumeValue 'CurrentReportMatches' $currentReportMatches
        Set-GssResumeValue 'TransactionStatus' $(if ($currentReportMatches) { 'LocalSnapshotPending' } else { 'Committed' })
        if ($currentReportMatches) { Write-GssOperationJson -Path $activePath -Value $receipt }
        try {
            $snapshot = New-GssRunSnapshot -Run $run -Folder $Folder
            Set-GssResumeValue 'LocalSnapshotPath' (Get-GssOperationRelativePath -Root $Folder -Path $snapshot)
            Set-GssResumeValue 'LocalSnapshotStatus' 'Verified'
            Set-GssResumeValue 'TransactionStatus' 'Committed'
            if ($currentReportMatches) { Write-GssOperationJson -Path $activePath -Value $receipt }
        }
        catch {
            Set-GssResumeValue 'LocalSnapshotStatus' 'Failed'
            Set-GssResumeValue 'BackupStatus' 'Pending'
            Set-GssResumeValue 'BackupError' $_.Exception.Message
        }
        if ([string]$receipt.LocalSnapshotStatus -eq 'Verified') {
            try {
                $backup = Sync-GssRunSnapshot -ManifestPath $snapshot -BackupFolder $BackupFolder
                Set-GssResumeValue 'BackupStatus' $backup.BackupStatus
                Set-GssResumeValue 'Backup' $backup
            }
            catch { Set-GssResumeValue 'BackupStatus' 'Pending'; Set-GssResumeValue 'BackupError' $_.Exception.Message }
        }
        if ([string]$receipt.BackupStatus -eq 'Committed') {
            $analysisRun = $run
            if ($FreshPreparedRunLogPath) {
                $freshRelative = Get-GssOperationRelativePath -Root $Folder -Path $FreshPreparedRunLogPath
                if (-not $freshRelative.StartsWith('_automation_runs/logs/', [StringComparison]::OrdinalIgnoreCase)) { throw 'Fresh report evidence must be a prepared log under the GSS log directory.' }
                $fresh = Get-Content -LiteralPath $FreshPreparedRunLogPath -Raw | ConvertFrom-Json
                $fresh | Add-Member -NotePropertyName LogPath -NotePropertyValue $FreshPreparedRunLogPath -Force
                $analysisRun = ConvertTo-GssRefreshedAnalysisRun -ReportRun $run -PreparedRun $fresh -Folder $Folder
            }
            # Package v2 still verifies its current external sources. An older
            # report's backup can finish after another week, but its package must
            # not silently use that newer week's live workbook or changed inputs.
            $packageInputsMatch = $true
            foreach ($check in @(
                @($analysisRun.TargetWorkbook, $analysisRun.PromotedWorkbookSha256),
                @($analysisRun.EmailComparisonPdf, $analysisRun.PromotedPdfSha256),
                @($analysisRun.CurrentSourceWorkbook, $analysisRun.CurrentSourceSha256),
                @($analysisRun.PriorYearSourceWorkbook, $analysisRun.PriorYearSourceSha256)
            )) {
                if (-not (Test-Path -LiteralPath $check[0] -PathType Leaf) -or (Get-GssOperationHash $check[0]) -cne [string]$check[1]) { $packageInputsMatch = $false; break }
            }
            if (-not $packageInputsMatch) {
                Set-GssResumeValue 'PackageStatus' 'SupersededOrChanged'
                Set-GssResumeValue 'PackageError' 'Backup completed. Current report inputs differ; prepare a fresh report/package instead of publishing changed evidence.'
            }
            else {
                try {
                    # Preserve the original log. Only a derived portable view is
                    # written when its root changed; the original stays auditable.
                    $analysisLog = $LiveRunLogPath
                    if ([string]$originalRun.Folder -cne $Folder -or $FreshPreparedRunLogPath) {
                        $analysisLog = Join-Path $logsRoot "gss_resume_$RunId`_$([guid]::NewGuid().ToString('N')).json"
                        $analysisRun | Add-Member -NotePropertyName OriginalRunLogSha256 -NotePropertyValue (Get-GssOperationHash $LiveRunLogPath) -Force
                        $analysisRun.LogPath = $analysisLog
                        Write-GssOperationJson -Path $analysisLog -Value $analysisRun
                    }
                    $analyzer = Join-Path $scriptRoot 'Analyze-GSS-Run.ps1'
                    $review = & $analyzer -Folder $Folder -LogPath $analysisLog -OutputObject
                    if ($review.WorkbookStatus -eq 'Blocked') { throw 'The current workbook no longer passes report validation.' }
                    $published = & $analyzer -Folder $Folder -LogPath $analysisLog -OutputObject -PublishEmailPackage
                    $ready = [bool]($published.EmailReadiness -eq 'Ready' -and $published.EmailPackage)
                    Set-GssResumeValue 'PackagePublished' $ready
                    Set-GssResumeValue 'PackageStatus' $(if ($ready) { 'Ready' } elseif ($published.EmailReadiness -eq 'WaitingForInput') { 'WaitingForInput' } else { 'Blocked' })
                    if ($ready) { Set-GssResumeValue 'PackagePath' (Get-GssOperationRelativePath -Root $Folder -Path $published.EmailPackage.PackagePath) }
                    if ($ready) { Set-GssResumeValue 'DraftEvidenceBackupStatus' 'LocalOnly' }
                }
                catch { Set-GssResumeValue 'PackageStatus' 'Blocked'; Set-GssResumeValue 'PackageError' $_.Exception.Message }
            }
        }
        else { Set-GssResumeValue 'PackageStatus' 'WaitingForBackup' }
        Save-GssResumeResult
        if (Test-Path -LiteralPath $activePath) {
            $active = Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json
            if ([string]$active.RunId -eq $RunId) {
                if ([string]$receipt.LocalSnapshotStatus -eq 'Verified' -or -not $currentReportMatches) { Invoke-GssFileOperationRetry { Remove-Item -LiteralPath $activePath -Force -ErrorAction Stop } }
                else { Write-GssOperationJson -Path $activePath -Value $receipt }
            }
        }
        Write-Information "Report remains ready. Backup: $($receipt.BackupStatus). Draft package: $($receipt.PackageStatus)." -InformationAction Continue
    }
}
catch {
    Set-GssResumeValue 'Error' $_.Exception.Message
    if ($ownsMutex) { Save-GssResumeResult }
    throw
}
finally {
    if ($ownsMutex) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
if ($OutputObject) { $receipt }
