[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Commission', 'RecordMetadataReadback', 'Prepare', 'Finalize', 'RetryFinalize', 'Abort', 'Inventory', 'RetentionReport', 'VerifyRestore', 'RestoreDrillStatus', 'CapacityProjection')]
    [string]$Operation,
    [string]$RunSummaryPath,
    [string]$RunId,
    [string]$SettingsPath,
    [string]$DriveRootPath,
    [string]$DriveFolderId,
    [string]$ExpectedOwner,
    [string]$MarkerId,
    [Nullable[bool]]$Shared,
    [Nullable[int]]$PermissionCount,
    [datetime]$VerifiedAtUtc = [datetime]::UtcNow,
    [ValidateSet('Final', 'Prepared')]
    [string]$RestorePhase = 'Final',
    [Nullable[long]]$FreeBytes,
    [switch]$OutputObject
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Gss-DriveBackup.ps1')

if ([string]::IsNullOrWhiteSpace($SettingsPath)) {
    $SettingsPath = Get-GssDriveBackupDefaultSettingsPath
}

function Get-RunSummaryContext {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $summary = Read-GssDriveBackupJson -Path $resolved
    $transaction = Get-GssDriveBackupProperty $summary @('Transaction', 'transaction')
    $backup = Get-GssDriveBackupProperty $summary @('Backup', 'backup')

    $resolvedRunId = [string](Get-GssDriveBackupProperty $summary @('RunId', 'run_id'))
    if ([string]::IsNullOrWhiteSpace($resolvedRunId)) {
        $resolvedRunId = [string](Get-GssDriveBackupProperty $transaction @('RunId', 'run_id'))
    }
    $fingerprint = [string](Get-GssDriveBackupProperty $summary @('RunFingerprint', 'Fingerprint', 'run_fingerprint', 'fingerprint'))
    if ([string]::IsNullOrWhiteSpace($fingerprint)) {
        $fingerprint = [string](Get-GssDriveBackupProperty $transaction @('RunFingerprint', 'Fingerprint', 'run_fingerprint', 'fingerprint'))
    }
    $gssRoot = [string](Get-GssDriveBackupProperty $summary @('Folder', 'GssRoot', 'gss_root'))
    if ([string]::IsNullOrWhiteSpace($gssRoot)) {
        $gssRoot = [string](Get-GssDriveBackupProperty $transaction @('Folder', 'GssRoot', 'gss_root'))
    }
    $reportWeekText = [string](Get-GssDriveBackupProperty $summary @('CurrentWeekEnding', 'ReportWeek', 'report_week'))
    if ([string]::IsNullOrWhiteSpace($reportWeekText)) {
        $reportWeekText = [string](Get-GssDriveBackupProperty $transaction @('CurrentWeekEnding', 'ReportWeek', 'report_week'))
    }
    $release = [string](Get-GssDriveBackupProperty $summary @('ProgramRelease', 'Release', 'program_release', 'release') 'unversioned')
    if ($release -eq 'unversioned') {
        $release = [string](Get-GssDriveBackupProperty $transaction @('ProgramRelease', 'Release', 'program_release', 'release') 'unversioned')
    }
    $snapshotPurpose = [string](Get-GssDriveBackupProperty $summary @('SnapshotPurpose', 'snapshot_purpose') 'WorkbookTransaction')
    if ($snapshotPurpose -eq 'WorkbookTransaction') {
        $snapshotPurpose = [string](Get-GssDriveBackupProperty $transaction @('SnapshotPurpose', 'snapshot_purpose') 'WorkbookTransaction')
    }
    if ($snapshotPurpose -notin @('WorkbookTransaction', 'RecoveryOnly')) {
        throw "Run summary SnapshotPurpose must be WorkbookTransaction or RecoveryOnly: $snapshotPurpose"
    }
    $inventoryMode = if ($snapshotPurpose -eq 'RecoveryOnly') { 'RecoveryOnly' } else { 'Full' }

    if ([string]::IsNullOrWhiteSpace($resolvedRunId)) {
        throw "Run summary must contain RunId so Drive preparation is bound to the reviewed transaction: $resolved"
    }
    if ([string]::IsNullOrWhiteSpace($fingerprint)) {
        throw "Run summary must contain RunFingerprint or Fingerprint so Drive preparation is bound to the reviewed transaction: $resolved"
    }
    if ([string]::IsNullOrWhiteSpace($gssRoot)) {
        throw "Run summary must contain Folder or GssRoot: $resolved"
    }
    if ([string]::IsNullOrWhiteSpace($reportWeekText)) {
        throw "Run summary must contain CurrentWeekEnding or ReportWeek: $resolved"
    }
    try {
        $reportWeek = [datetime]::Parse($reportWeekText, [Globalization.CultureInfo]::InvariantCulture).Date
    }
    catch {
        throw "Run summary report week is invalid: $reportWeekText"
    }

    $additionalItems = @()
    foreach ($container in @($summary, $transaction, $backup)) {
        if ($null -eq $container) { continue }
        $items = Get-GssDriveBackupProperty $container @('BackupInventory', 'backup_inventory') @()
        if ($null -ne $items) { $additionalItems += @($items) }
    }

    $transactionArtifactPaths = @()
    foreach ($container in @($summary, $transaction, $backup)) {
        if ($null -eq $container) { continue }
        $items = Get-GssDriveBackupProperty $container @('TransactionArtifacts', 'transaction_artifacts') @()
        foreach ($item in @($items)) {
            if ($item -is [string]) {
                $transactionArtifactPaths += [string]$item
            }
            else {
                $pathValue = [string](Get-GssDriveBackupProperty $item @('SourcePath', 'source_path', 'Path', 'path'))
                if (-not [string]::IsNullOrWhiteSpace($pathValue)) { $transactionArtifactPaths += $pathValue }
            }
        }
    }

    $releaseArchivePaths = @()
    foreach ($container in @($summary, $transaction, $backup)) {
        if ($null -eq $container) { continue }
        $items = Get-GssDriveBackupProperty $container @('ReleaseArchives', 'release_archives') @()
        foreach ($item in @($items)) {
            if ($item -is [string]) {
                $releaseArchivePaths += [string]$item
            }
            else {
                $pathValue = [string](Get-GssDriveBackupProperty $item @('SourcePath', 'source_path', 'Path', 'path'))
                if (-not [string]::IsNullOrWhiteSpace($pathValue)) { $releaseArchivePaths += $pathValue }
            }
        }
    }

    return [pscustomobject]@{
        SummaryPath = $resolved
        Summary = $summary
        RunId = $resolvedRunId
        Fingerprint = $fingerprint
        GssRoot = [System.IO.Path]::GetFullPath($gssRoot)
        ReportWeek = $reportWeek
        Release = $release
        SnapshotPurpose = $snapshotPurpose
        InventoryMode = $inventoryMode
        AdditionalItems = @($additionalItems)
        TransactionArtifactPaths = @($transactionArtifactPaths | Select-Object -Unique)
        ReleaseArchivePaths = @($releaseArchivePaths | Select-Object -Unique)
    }
}

function Get-ContextInventory {
    param(
        [Parameter(Mandatory)]
        [object]$Context
    )

    return @(Get-GssDriveBackupInventory `
        -GssRoot $Context.GssRoot `
        -AdditionalItems $Context.AdditionalItems `
        -TransactionArtifactPaths $Context.TransactionArtifactPaths `
        -ReleaseArchivePaths $Context.ReleaseArchivePaths `
        -InventoryMode $Context.InventoryMode)
}

$result = switch ($Operation) {
    'Commission' {
        if ([string]::IsNullOrWhiteSpace($DriveRootPath) -or
            [string]::IsNullOrWhiteSpace($DriveFolderId) -or
            [string]::IsNullOrWhiteSpace($ExpectedOwner)) {
            throw 'Commission requires DriveRootPath, DriveFolderId, and ExpectedOwner. MarkerId is generated if omitted.'
        }
        Initialize-GssDriveBackupConfiguration `
            -DriveRootPath $DriveRootPath `
            -DriveFolderId $DriveFolderId `
            -ExpectedOwner $ExpectedOwner `
            -MarkerId $MarkerId `
            -SettingsPath $SettingsPath
        break
    }
    'RecordMetadataReadback' {
        if ([string]::IsNullOrWhiteSpace($DriveFolderId) -or
            [string]::IsNullOrWhiteSpace($ExpectedOwner) -or
            $null -eq $Shared -or
            $null -eq $PermissionCount) {
            throw 'RecordMetadataReadback requires DriveFolderId, ExpectedOwner, Shared, and PermissionCount from the Google Drive connector.'
        }
        Write-GssDriveBackupCommissioningReadback `
            -DriveFolderId $DriveFolderId `
            -Owner $ExpectedOwner `
            -Shared ([bool]$Shared) `
            -PermissionCount ([int]$PermissionCount) `
            -VerifiedAtUtc $VerifiedAtUtc `
            -SettingsPath $SettingsPath
        break
    }
    'Prepare' {
        if ([string]::IsNullOrWhiteSpace($RunSummaryPath)) { throw 'Prepare requires RunSummaryPath.' }
        $context = Get-RunSummaryContext -Path $RunSummaryPath
        $inventory = @(Get-ContextInventory -Context $context)
        New-GssDriveBackupPreparedSnapshot `
            -RunId $context.RunId `
            -Fingerprint $context.Fingerprint `
            -ReportWeek $context.ReportWeek `
            -Inventory $inventory `
            -Release $context.Release `
            -SnapshotPurpose $context.SnapshotPurpose `
            -SettingsPath $SettingsPath
        break
    }
    { $_ -in @('Finalize', 'RetryFinalize') } {
        if ([string]::IsNullOrWhiteSpace($RunSummaryPath)) { throw "$Operation requires RunSummaryPath." }
        $context = Get-RunSummaryContext -Path $RunSummaryPath
        try {
            $inventory = @(Get-ContextInventory -Context $context)
            Complete-GssDriveBackupSnapshot `
                -RunId $context.RunId `
                -Fingerprint $context.Fingerprint `
                -FinalInventory $inventory `
                -SnapshotPurpose $context.SnapshotPurpose `
                -SettingsPath $SettingsPath
        }
        catch {
            $backupStatus = 'PendingFinalize'
            try {
                $recordedStatus = Get-GssDriveBackupRunStatus -RunId $context.RunId -SettingsPath $SettingsPath
                if ($recordedStatus -eq 'Blocked') {
                    $backupStatus = 'Blocked'
                }
            }
            catch { Write-Verbose "Could not read recorded Drive backup status after finalization failure: $($_.Exception.Message)" }
            [pscustomobject]@{
                Status = $backupStatus
                BackupStatus = $backupStatus
                RunId = $context.RunId
                Fingerprint = $context.Fingerprint
                Error = $_.Exception.Message
                RetryOperation = 'RetryFinalize'
                VerificationLevel = 'drivefs_hash_verified'
            }
        }
        break
    }
    'Abort' {
        if ([string]::IsNullOrWhiteSpace($RunSummaryPath)) { throw 'Abort requires RunSummaryPath.' }
        $context = Get-RunSummaryContext -Path $RunSummaryPath
        Stop-GssDriveBackupPreparedSnapshot `
            -RunId $context.RunId `
            -Fingerprint $context.Fingerprint `
            -SettingsPath $SettingsPath
        break
    }
    'Inventory' {
        if ([string]::IsNullOrWhiteSpace($RunSummaryPath)) { throw 'Inventory requires RunSummaryPath.' }
        $context = Get-RunSummaryContext -Path $RunSummaryPath
        $inventory = @(Get-ContextInventory -Context $context)
        [pscustomobject]@{
            Status = 'Inventoried'
            RunId = $context.RunId
            FileCount = $inventory.Count
            TotalBytes = [long](($inventory | ForEach-Object { (Get-Item -LiteralPath $_.SourcePath).Length } | Measure-Object -Sum).Sum)
            ContainsPersonalData = [bool](@($inventory | Where-Object Classification -eq 'restricted_personal_data').Count -gt 0)
            SnapshotPurpose = $context.SnapshotPurpose
            InventoryMode = $context.InventoryMode
            Excluded = @('test-output', '_automation_runs/backups except explicit transaction artifacts', 'quarantine', 'staging', 'temp', 'repository working tree', '.git')
            Inventory = $inventory
        }
        break
    }
    'RetentionReport' {
        Write-GssDriveBackupRetentionReport -SettingsPath $SettingsPath
        break
    }
    'VerifyRestore' {
        $restoreRunId = $RunId
        if ([string]::IsNullOrWhiteSpace($restoreRunId) -and -not [string]::IsNullOrWhiteSpace($RunSummaryPath)) {
            $restoreRunId = (Get-RunSummaryContext -Path $RunSummaryPath).RunId
        }
        if ([string]::IsNullOrWhiteSpace($restoreRunId)) { throw 'VerifyRestore requires RunId or RunSummaryPath.' }
        Restore-GssDriveBackupForVerification -RunId $restoreRunId -SettingsPath $SettingsPath -Phase $RestorePhase
        break
    }
    'RestoreDrillStatus' {
        Get-GssDriveBackupRestoreDrillStatus
        break
    }
    'CapacityProjection' {
        if ([string]::IsNullOrWhiteSpace($RunSummaryPath)) { throw 'CapacityProjection requires RunSummaryPath.' }
        $context = Get-RunSummaryContext -Path $RunSummaryPath
        $inventory = @(Get-ContextInventory -Context $context)
        $projectedBytes = [long](($inventory | ForEach-Object { (Get-Item -LiteralPath $_.SourcePath).Length } | Measure-Object -Sum).Sum)
        Get-GssDriveBackupCapacityProjection -ProjectedWeeklyBytes $projectedBytes -FreeBytes $FreeBytes -SettingsPath $SettingsPath
        break
    }
}

if ($OutputObject) {
    Write-Output $result
}
else {
    $result | Format-List
}
