[CmdletBinding()]
param(
    [string]$Folder,
    [string]$PopulationExportPath
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
$releaseValidator = Join-Path $scriptRoot 'Test-GSS-ReleaseIntegrity.ps1'
# Production runs are permitted only from the exact clean, tagged release
# described by release/release-manifest.json. There is intentionally no runtime
# bypass; development copy-tests invoke Update-GSS-MainWorkbook.ps1 directly.
$releaseIntegrity = & $releaseValidator -RepoRoot $repoRoot

function Write-GssSafeAtomicJson {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$InputObject
    )

    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $temporaryPath = Join-Path $directory ('.t-{0}' -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
    try {
        $json = $InputObject | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($temporaryPath, $json, (New-Object System.Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            [System.IO.File]::Replace($temporaryPath, $Path, $null, $true)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $Path)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}

function Get-GssSafeHash {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required transaction artifact is missing: $Path"
    }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Assert-GssPreparedEvidenceUnchanged {
    param([Parameter(Mandatory)][object]$Run)

    if ([string]$Run.TransactionStatus -ne 'Prepared' -or [string]::IsNullOrWhiteSpace([string]$Run.RunFingerprint)) {
        throw 'The copy-test did not produce a promotable Prepared fingerprint.'
    }
    if ([string]$Run.HostName -cne [Environment]::MachineName) {
        throw "The copy-test was prepared on '$($Run.HostName)', not this workstation."
    }
    foreach ($check in @(
        [pscustomobject]@{ Name = 'starting live workbook'; Path = [string]$Run.StartingWorkbook; Expected = [string]$Run.StartingWorkbookSha256 },
        [pscustomobject]@{ Name = 'current source workbook'; Path = [string]$Run.CurrentSourceWorkbook; Expected = [string]$Run.CurrentSourceSha256 },
        [pscustomobject]@{ Name = 'prior-year source workbook'; Path = [string]$Run.PriorYearSourceWorkbook; Expected = [string]$Run.PriorYearSourceSha256 },
        [pscustomobject]@{ Name = 'staged workbook'; Path = [string]$Run.StagedWorkbook; Expected = [string]$Run.StagedWorkbookSha256 },
        [pscustomobject]@{ Name = 'staged PDF'; Path = [string]$Run.StagedPdf; Expected = [string]$Run.StagedPdfSha256 }
    )) {
        $actual = Get-GssSafeHash $check.Path
        if ($actual -ne $check.Expected.ToLowerInvariant()) {
            throw "The $($check.Name) changed after review. Run a fresh copy-test."
        }
    }
}

function Get-GssBackupStatus {
    param([object]$Result)

    if ($null -eq $Result) { return '' }
    if ($Result.PSObject.Properties.Name -contains 'BackupStatus') { return [string]$Result.BackupStatus }
    if ($Result.PSObject.Properties.Name -contains 'Status') { return [string]$Result.Status }
    return ''
}

function Invoke-GssPreparedBackupAbort {
    param(
        [Parameter(Mandatory)][string]$CoordinatorPath,
        [Parameter(Mandatory)][string]$RunSummaryPath
    )

    $abort = & $CoordinatorPath -Operation Abort -RunSummaryPath $RunSummaryPath -OutputObject
    if ((Get-GssBackupStatus $abort) -ne 'Aborted') {
        throw "Drive prepared snapshot did not reach Aborted status (actual: '$(Get-GssBackupStatus $abort)')."
    }
    return $abort
}

function Invoke-GssSafeAnalysis {
    param(
        [Parameter(Mandatory)][string]$AnalyzerPath,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$RunLogPath,
        [string]$PopulationPath,
        [switch]$PublishPackage
    )

    $arguments = @{
        Folder = $FolderPath
        LogPath = $RunLogPath
        OutputObject = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($PopulationPath)) {
        $arguments.PopulationExportPath = $PopulationPath
    }
    if ($PublishPackage) {
        $arguments.PublishEmailPackage = $true
    }
    return & $AnalyzerPath @arguments
}

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $Folder = Split-Path -Parent $repoRoot
}
$Folder = $Folder.Trim('"')
$Folder = (Resolve-Path -LiteralPath $Folder).Path

$DriveBackupScript = Join-Path $scriptRoot 'Invoke-GSS-DriveBackup.ps1'
$updater = Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1'
$analyzer = Join-Path $scriptRoot 'Analyze-GSS-Run.ps1'
$resumeFinalizer = Join-Path $scriptRoot 'Resume-GSS-PendingFinalize.ps1'
$runId = [guid]::NewGuid().ToString('D')
$stateDirectory = Join-Path $Folder '_automation_runs\state'
$activeRunPath = Join-Path $stateDirectory 'active-transaction.json'
$receiptPath = Join-Path $stateDirectory "transaction-$runId.json"
$mutex = New-Object System.Threading.Mutex($false, 'Global\GSSSurveyWorkbookAutomationTransaction')
$ownsMutex = $false
$ownsActiveMarker = $false
$receipt = [ordered]@{
    ReceiptSchemaVersion = 1
    RunId = $runId
    HostName = [Environment]::MachineName
    StartedUtc = [datetime]::UtcNow.ToString('o')
    UpdatedUtc = [datetime]::UtcNow.ToString('o')
    ProgramRelease = [string]$releaseIntegrity.ReleaseTag
    ReleaseCommit = [string]$releaseIntegrity.HeadCommit
    ReleaseIntegrityStatus = [string]$releaseIntegrity.Status
    TransactionStatus = 'Starting'
    BackupStatus = 'NotStarted'
    AutomaticSending = 'Disabled'
    ReceiptPath = $receiptPath
}

try {
    try {
        $ownsMutex = $mutex.WaitOne(0)
    }
    catch [System.Threading.AbandonedMutexException] {
        $ownsMutex = $true
    }
    if (-not $ownsMutex) {
        throw 'Another GSS workbook transaction is already active on this workstation.'
    }

    New-Item -ItemType Directory -Path $stateDirectory -Force | Out-Null
    if (Test-Path -LiteralPath $activeRunPath -PathType Leaf) {
        $existingActive = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
        $existingStatus = [string]$existingActive.TransactionStatus
        $existingReceiptPath = [string]$existingActive.ReceiptPath
        if (-not [string]::IsNullOrWhiteSpace($existingReceiptPath) -and
            (Test-Path -LiteralPath $existingReceiptPath -PathType Leaf)) {
            $existingReceipt = Get-Content -LiteralPath $existingReceiptPath -Raw | ConvertFrom-Json
            if ([string]$existingReceipt.RunId -eq [string]$existingActive.RunId) {
                $existingStatus = [string]$existingReceipt.TransactionStatus
            }
        }
        if ($existingStatus -notin @('Committed', 'PackageBlocked', 'Aborted', 'Blocked')) {
            throw "Unresolved GSS transaction '$($existingActive.RunId)' is $existingStatus. Resume or review its exact receipt before starting another run: $existingReceiptPath"
        }
        Remove-Item -LiteralPath $activeRunPath -Force
    }
    Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
    $ownsActiveMarker = $true
    Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt

    Write-Information '' -InformationAction Continue
    Write-Information 'STEP 1 OF 4 - Testing the update' -InformationAction Continue
    Write-Information 'The live workbook is not being changed.' -InformationAction Continue
    $copyRun = & $updater `
        -Folder $Folder `
        -RunId $runId `
        -ProgramRelease ([string]$releaseIntegrity.ReleaseTag) `
        -MutexAlreadyHeld `
        -OutputObject
    if (-not $copyRun -or [string]::IsNullOrWhiteSpace([string]$copyRun.LogPath)) {
        throw 'The copy-test updater did not return its exact run log path.'
    }
    $receipt.CopyRunLogPath = [string]$copyRun.LogPath
    $receipt.RunFingerprint = [string]$copyRun.RunFingerprint
    $receipt.TransactionStatus = 'CopyPrepared'
    $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
    Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
    Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt

    Write-Information '' -InformationAction Continue
    Write-Information 'STEP 2 OF 4 - Checking the test results' -InformationAction Continue
    $copyReview = Invoke-GssSafeAnalysis `
        -AnalyzerPath $analyzer `
        -FolderPath $Folder `
        -RunLogPath $copyRun.LogPath `
        -PopulationPath $PopulationExportPath
    if ($copyReview.WorkbookStatus -eq 'Blocked') {
        $receipt.TransactionStatus = 'Blocked'
        $receipt.Error = 'Copy-test review blocked live promotion.'
        $receipt.ReviewPath = [string]$copyReview.MarkdownPath
        $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
        Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt
        Write-Information '' -InformationAction Continue
        Write-Information 'ATTENTION NEEDED - NO LIVE CHANGES MADE' -InformationAction Continue
        Write-Information 'The test review found a problem. The live workbook was not updated.' -InformationAction Continue
        Write-Information "Review details: $($copyReview.MarkdownPath)" -InformationAction Continue
        exit 1
    }

    Write-Information '' -InformationAction Continue
    Write-Information 'STEP 3 OF 4 - Choose whether to update the live workbook' -InformationAction Continue
    Write-Information "Reviewed transaction fingerprint: $($copyRun.RunFingerprint)" -InformationAction Continue
    Write-Information 'The copy-test passed. Type APPLY and press Enter to continue.' -InformationAction Continue
    Write-Information 'Press Enter without typing APPLY to stop safely.' -InformationAction Continue
    $confirmation = Read-Host 'Confirmation'

    if ($confirmation -cne 'APPLY') {
        $receipt.TransactionStatus = 'Aborted'
        $receipt.BackupStatus = 'NotStarted'
        $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
        Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt
        Write-Information '' -InformationAction Continue
        Write-Information 'NO LIVE CHANGES MADE' -InformationAction Continue
        Write-Information 'The copy-test finished, but the live workbook was not updated.' -InformationAction Continue
        exit 0
    }

    # Recheck every fingerprint input after the literal APPLY confirmation and
    # before asking Drive to prepare the required recovery snapshot.
    Assert-GssPreparedEvidenceUnchanged $copyRun
    if (-not (Test-Path -LiteralPath $DriveBackupScript -PathType Leaf)) {
        throw "Required Google Drive backup coordinator is unavailable: $DriveBackupScript"
    }

    Write-Information '' -InformationAction Continue
    Write-Information 'Preparing the required private Google Drive recovery snapshot.' -InformationAction Continue
    $backupPrepare = & $DriveBackupScript `
        -Operation Prepare `
        -RunSummaryPath $copyRun.LogPath `
        -OutputObject
    $prepareStatus = Get-GssBackupStatus $backupPrepare
    if ($prepareStatus -ne 'Prepared' -or
        [string]$backupPrepare.RunId -ne [string]$copyRun.RunId -or
        [string]$backupPrepare.Fingerprint -ne [string]$copyRun.RunFingerprint) {
        throw "Google Drive backup preparation did not reach Prepared status (actual: '$prepareStatus')."
    }
    if ([string]::IsNullOrWhiteSpace([string]$backupPrepare.PreparedManifestPath) -or
        [string]::IsNullOrWhiteSpace([string]$backupPrepare.PreparedManifestSha256)) {
        throw 'Google Drive preparation did not return its exact manifest path and hash.'
    }
    $receipt.BackupStatus = 'Prepared'
    $receipt.BackupPrepare = $backupPrepare
    $receipt.TransactionStatus = 'BackupPrepared'
    $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
    Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
    Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt

    Write-Information '' -InformationAction Continue
    Write-Information 'STEP 4 OF 4 - Promoting the exact reviewed workbook' -InformationAction Continue
    $liveRun = & $updater `
        -Folder $Folder `
        -Apply `
        -PreparedRunLogPath $copyRun.LogPath `
        -ExpectedFingerprint $copyRun.RunFingerprint `
        -DrivePreparedManifestPath $backupPrepare.PreparedManifestPath `
        -ExpectedDrivePreparedManifestSha256 $backupPrepare.PreparedManifestSha256 `
        -ProgramRelease ([string]$releaseIntegrity.ReleaseTag) `
        -MutexAlreadyHeld `
        -OutputObject
    if (-not $liveRun -or [string]::IsNullOrWhiteSpace([string]$liveRun.LogPath)) {
        throw 'The live updater did not return its exact transaction log path.'
    }
    $receipt.LiveRunLogPath = [string]$liveRun.LogPath
    $receipt.TransactionStatus = 'WorkbookCommitted'
    $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
    Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
    Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt

    Write-Information 'Checking the completed live update before finalizing Drive.' -InformationAction Continue
    $liveReview = Invoke-GssSafeAnalysis `
        -AnalyzerPath $analyzer `
        -FolderPath $Folder `
        -RunLogPath $liveRun.LogPath `
        -PopulationPath $PopulationExportPath
    if ($liveReview.WorkbookStatus -eq 'Blocked') {
        $receipt.TransactionStatus = 'RollbackAttempting'
        $receipt.BackupStatus = 'Prepared'
        $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
        Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
        Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt
        try {
            $rollback = & $updater `
                -Folder $Folder `
                -RollbackRunLogPath $liveRun.LogPath `
                -ProgramRelease ([string]$releaseIntegrity.ReleaseTag) `
                -MutexAlreadyHeld `
                -OutputObject
        }
        catch {
            $receipt.TransactionStatus = 'RollbackBlocked'
            $receipt.BackupStatus = 'Prepared'
            $receipt.RollbackError = $_.Exception.Message
            $receipt.Error = 'The post-promotion review blocked the live workbook and conflict-aware rollback did not complete.'
            $receipt.ReviewPath = [string]$liveReview.MarkdownPath
            $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
            Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
            Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt
            throw
        }
        if ([string]$rollback.Status -eq 'RolledBack') {
            $backupAbort = Invoke-GssPreparedBackupAbort -CoordinatorPath $DriveBackupScript -RunSummaryPath $copyRun.LogPath
            $receipt.TransactionStatus = 'Aborted'
            $receipt.BackupStatus = 'Aborted'
            $receipt.BackupAbort = $backupAbort
        }
        else {
            $receipt.TransactionStatus = 'RollbackBlocked'
            $receipt.BackupStatus = 'Prepared'
        }
        $receipt.Rollback = $rollback
        $receipt.Error = 'The post-promotion review blocked the live workbook.'
        $receipt.ReviewPath = [string]$liveReview.MarkdownPath
        $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
        Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
        Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt
        Write-Information '' -InformationAction Continue
        Write-Information 'ATTENTION NEEDED' -InformationAction Continue
        Write-Information 'The final review failed and the conflict-aware rollback was attempted.' -InformationAction Continue
        Write-Information "Review details: $($liveReview.MarkdownPath)" -InformationAction Continue
        exit 1
    }

    # A valid workbook is never rolled back solely because cloud finalization
    # is unavailable. It remains applied, while package publication stays
    # blocked and RetryFinalize can resume idempotently from this log.
    try {
        $backupFinalize = & $DriveBackupScript `
            -Operation Finalize `
            -RunSummaryPath $liveRun.LogPath `
            -OutputObject
        $finalizeStatus = Get-GssBackupStatus $backupFinalize
        if ($finalizeStatus -eq 'Committed' -and
            ([string]$backupFinalize.RunId -ne [string]$liveRun.RunId -or
                [string]$backupFinalize.Fingerprint -ne [string]$liveRun.RunFingerprint)) {
            throw 'Drive finalization returned Committed for different transaction evidence.'
        }
    }
    catch {
        $backupFinalize = [pscustomobject]@{
            BackupStatus = 'PendingFinalize'
            Error = $_.Exception.Message
        }
        $finalizeStatus = 'PendingFinalize'
    }
    $receipt.BackupStatus = $finalizeStatus
    $receipt.BackupFinalize = $backupFinalize
    $receipt.ReviewPath = [string]$liveReview.MarkdownPath
    $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')

    if ($finalizeStatus -eq 'Blocked') {
        $receipt.TransactionStatus = 'BackupBlocked'
        $receipt.BackupStatus = 'Blocked'
        Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
        Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt
        Write-Information '' -InformationAction Continue
        Write-Information 'WORKBOOK READY - DRIVE BACKUP BLOCKED' -InformationAction Continue
        Write-Information 'The validated workbook remains applied, but Drive reported a non-retryable conflict.' -InformationAction Continue
        Write-Information 'Package publication is blocked. Review the Drive backup status and chain evidence manually.' -InformationAction Continue
        exit 1
    }
    if ($finalizeStatus -ne 'Committed') {
        $receipt.TransactionStatus = 'PendingFinalize'
        $receipt.BackupStatus = 'PendingFinalize'
        Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
        Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt
        Write-Information '' -InformationAction Continue
        Write-Information 'WORKBOOK READY - DRIVE FINALIZATION PENDING' -InformationAction Continue
        Write-Information 'The validated workbook remains applied. Email package publication is blocked.' -InformationAction Continue
        Write-Information "Resume with: powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$resumeFinalizer`" -LiveRunLogPath `"$($liveRun.LogPath)`"" -InformationAction Continue
        exit 1
    }

    # Package publication occurs only after Drive returned Committed.
    $receipt.TransactionStatus = 'BackupCommitted'
    $receipt.BackupStatus = 'Committed'
    $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
    Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
    Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt
    $publishedReview = Invoke-GssSafeAnalysis `
        -AnalyzerPath $analyzer `
        -FolderPath $Folder `
        -RunLogPath $liveRun.LogPath `
        -PopulationPath $PopulationExportPath `
        -PublishPackage
    $receipt.PackagePublished = [bool]($publishedReview.EmailReadiness -eq 'Ready' -and $publishedReview.EmailPackage)
    $receipt.PackagePath = if ($publishedReview.EmailPackage) { [string]$publishedReview.EmailPackage.PackagePath } else { $null }
    if ($receipt.PackagePublished) {
        $receipt.TransactionStatus = 'Committed'
        $receipt.CompletedUtc = [datetime]::UtcNow.ToString('o')
    }
    else {
        $receipt.TransactionStatus = 'PackageBlocked'
        $receipt.Error = @($publishedReview.Qa.EmailBlockers) -join '; '
    }
    $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
    Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt

    Write-Information '' -InformationAction Continue
    Write-Information 'UPDATE COMPLETE' -InformationAction Continue
    Write-Information 'The exact reviewed workbook was applied, validated, and committed to the private Drive backup.' -InformationAction Continue
    Write-Information "Review details: $($publishedReview.MarkdownPath)" -InformationAction Continue
    if ($publishedReview.EmailReadiness -eq 'Ready' -and $publishedReview.EmailPackage) {
        Write-Information "Email package ready: $($publishedReview.EmailPackage.PackagePath)" -InformationAction Continue
    }
    else {
        Write-Information 'EMAIL PACKAGE NOT READY' -InformationAction Continue
        Write-Information 'The workbook and Drive backup succeeded, but package publication remains blocked.' -InformationAction Continue
        foreach ($blocker in @($publishedReview.Qa.EmailBlockers)) {
            Write-Information "  - $blocker" -InformationAction Continue
        }
        exit 1
    }
}
catch {
    $outerFailure = $_
    if ($receipt.BackupStatus -eq 'Prepared' -and
        $receipt.TransactionStatus -notin @('WorkbookCommitted', 'PendingFinalize', 'BackupCommitted', 'Committed', 'PackageBlocked', 'RollbackAttempting', 'RollbackBlocked', 'BackupBlocked')) {
        try {
            $canAbort = $false
            if ($copyRun -and $copyRun.StartingWorkbook -and (Test-Path -LiteralPath $copyRun.StartingWorkbook -PathType Leaf)) {
                $canAbort = (Get-GssSafeHash ([string]$copyRun.StartingWorkbook)) -eq ([string]$copyRun.StartingWorkbookSha256).ToLowerInvariant()
            }
            if ($canAbort) {
                $backupAbort = Invoke-GssPreparedBackupAbort -CoordinatorPath $DriveBackupScript -RunSummaryPath $copyRun.LogPath
                $receipt.BackupStatus = 'Aborted'
                $receipt.BackupAbort = $backupAbort
                $receipt.TransactionStatus = 'Aborted'
            }
            else {
                $receipt.TransactionStatus = 'RollbackBlocked'
            }
        }
        catch {
            $receipt.TransactionStatus = 'Blocked'
            $receipt.BackupAbortError = $_.Exception.Message
        }
    }
    $receipt.TransactionStatus = if ($receipt.BackupStatus -eq 'Committed' -or $receipt.TransactionStatus -eq 'BackupCommitted') {
        'PackageBlocked'
    }
    elseif ($receipt.TransactionStatus -eq 'WorkbookCommitted') {
        'PendingFinalize'
    }
    elseif ($receipt.TransactionStatus -eq 'RollbackAttempting') {
        'RollbackBlocked'
    }
    elseif ($receipt.TransactionStatus -in @('Aborted', 'RollbackBlocked', 'BackupBlocked')) {
        $receipt.TransactionStatus
    }
    else {
        'Blocked'
    }
    $receipt.Error = $outerFailure.Exception.Message
    $receipt.UpdatedUtc = [datetime]::UtcNow.ToString('o')
    try {
        if ($ownsActiveMarker -and (Test-Path -LiteralPath $activeRunPath -PathType Leaf)) {
            Write-GssSafeAtomicJson -Path $activeRunPath -InputObject $receipt
        }
        Write-GssSafeAtomicJson -Path $receiptPath -InputObject $receipt
    }
    catch { Write-Verbose "The final transaction receipt could not be updated: $($_.Exception.Message)" }
    throw
}
finally {
    if ($ownsActiveMarker -and
        $receipt.TransactionStatus -in @('Committed', 'PackageBlocked', 'Aborted', 'Blocked') -and
        (Test-Path -LiteralPath $activeRunPath -PathType Leaf)) {
        try {
            $active = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
            if ([string]$active.RunId -eq $runId) {
                Remove-Item -LiteralPath $activeRunPath -Force
            }
        }
        catch { Write-Verbose "The active transaction marker could not be inspected or removed: $($_.Exception.Message)" }
    }
    if ($ownsMutex) {
        try { $mutex.ReleaseMutex() }
        catch { Write-Verbose "The GSS transaction mutex could not be released cleanly: $($_.Exception.Message)" }
    }
    $mutex.Dispose()
}
