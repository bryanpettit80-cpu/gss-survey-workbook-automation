[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$invokeScript = Join-Path $scriptRoot 'Invoke-GSS-DriveBackup.ps1'
. (Join-Path $scriptRoot 'Gss-DriveBackup.ps1')

function Assert-GssDriveBackupTest {
    param(
        [Parameter(Mandatory)]
        [bool]$Condition,
        [Parameter(Mandatory)]
        [string]$Message
    )
    if (-not $Condition) {
        throw "Drive backup test failed: $Message"
    }
}

$systemTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\', '/')
$testRoot = Join-Path $systemTemp ("gss-drive-backup-test-$([guid]::NewGuid().ToString('N'))")
$resolvedTestRoot = [System.IO.Path]::GetFullPath($testRoot)
if (-not $resolvedTestRoot.StartsWith("$systemTemp\", [System.StringComparison]::OrdinalIgnoreCase) -or
    -not ([System.IO.Path]::GetFileName($resolvedTestRoot)).StartsWith('gss-drive-backup-test-', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Unsafe test directory resolution: $resolvedTestRoot"
}

try {
    $gssRoot = Join-Path $resolvedTestRoot 'GSS Surveys'
    $driveRoot = Join-Path $resolvedTestRoot 'Drive\GSS Survey Backups'
    $localAppData = Join-Path $resolvedTestRoot 'LocalAppData'
    $settingsPath = Join-Path $localAppData 'GSSSurveyWorkbookAutomation\settings.json'
    $releasePath = Join-Path $resolvedTestRoot 'gss-release-v1.0.0.zip'
    foreach ($folder in @(
        $gssRoot,
        $driveRoot,
        (Join-Path $gssRoot '01 Main Workbook'),
        (Join-Path $gssRoot '02 Weekly Rolling Source Workbooks'),
        (Join-Path $gssRoot '03 Uploaded Survey Workbooks'),
        (Join-Path $gssRoot '04 Email Comparison PDFs'),
        (Join-Path $gssRoot '05 Reference Materials'),
        (Join-Path $gssRoot '06 Exports and Images'),
        (Join-Path $gssRoot '_automation_runs\qa'),
        (Join-Path $gssRoot '_automation_runs\logs'),
        (Join-Path $gssRoot '_automation_runs\state'),
        (Join-Path $gssRoot '_automation_runs\backups'),
        (Join-Path $gssRoot '_automation_runs\quarantine'),
        (Join-Path $gssRoot '_automation_runs\test-output'),
        (Join-Path $gssRoot '_automation_runs\email_outbox\ready-run'),
        (Join-Path $gssRoot '_automation_runs\email_outbox\.staging-run'),
        (Join-Path $gssRoot 'GSS Survey Workbook Automation\.git')
    )) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    Set-Content -LiteralPath (Join-Path $gssRoot 'Run GSS Update After Upload.cmd') -Value '@echo test' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '00 START HERE - GSS Survey Updates.txt') -Value 'operator guide' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '01 Main Workbook\GSS Score Trends - Main.xlsx') -Value 'pre-apply-workbook' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '02 Weekly Rolling Source Workbooks\rolling.xlsx') -Value 'rolling' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '03 Uploaded Survey Workbooks\GSS Guest Detail.xlsx') -Value 'guest personal data' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '04 Email Comparison PDFs\comparison.pdf') -Value 'pdf' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '05 Reference Materials\reference.txt') -Value 'reference' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '06 Exports and Images\export.csv') -Value 'export' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\qa\qa.json') -Value '{}' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\logs\run.json') -Value '{}' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\state\state.json') -Value '{}' -Encoding UTF8
    $transactionArtifact = Join-Path $gssRoot '_automation_runs\backups\pre-apply.xlsx'
    Set-Content -LiteralPath $transactionArtifact -Value 'exact pre-apply artifact' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\quarantine\excluded.xlsx') -Value 'excluded' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\test-output\excluded.xlsx') -Value 'excluded' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\email_outbox\ready-run\READY') -Value 'ready' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\email_outbox\ready-run\GSS Guest Detail.xlsx') -Value 'ready personal data' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\email_outbox\.staging-run\READY') -Value 'not actually ready' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\email_outbox\.staging-run\excluded.xlsx') -Value 'excluded' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot 'GSS Survey Workbook Automation\.git\config') -Value 'excluded' -Encoding UTF8
    Set-Content -LiteralPath $releasePath -Value 'release archive' -Encoding UTF8

    $commission = & $invokeScript `
        -Operation Commission `
        -DriveRootPath $driveRoot `
        -DriveFolderId 'test-drive-folder-id' `
        -ExpectedOwner 'owner@example.com' `
        -SettingsPath $settingsPath `
        -OutputObject
    Assert-GssDriveBackupTest ($commission.Status -eq 'Configured') 'Commission did not succeed.'
    Assert-GssDriveBackupTest $commission.ReadbackVerified 'Commission did not report marker/settings readback verification.'
    Assert-GssDriveBackupTest (Test-Path -LiteralPath $commission.ClassificationNoticePath -PathType Leaf) 'Classification notice was not created.'
    Assert-GssDriveBackupTest (-not [string]::IsNullOrWhiteSpace($commission.MarkerId)) 'Commission did not generate a marker ID.'
    $commissionMarker = Read-GssDriveBackupJson -Path $commission.MarkerPath
    Assert-GssDriveBackupTest ($commissionMarker.classification -eq $script:GssDriveBackupClassificationLabel) 'Commission marker classification label is incorrect.'

    $missingMetadataRefused = $false
    try {
        [void](Get-GssDriveBackupRootContext -SettingsPath $settingsPath)
    }
    catch {
        $missingMetadataRefused = $_.Exception.Message -like '*commissioning-readback.json*'
    }
    Assert-GssDriveBackupTest $missingMetadataRefused 'Operational Drive access did not require owner-only connector metadata evidence.'

    $longAtomicParent = Join-Path $resolvedTestRoot ('long-atomic-' + ('x' * 70))
    New-Item -ItemType Directory -Path $longAtomicParent -Force | Out-Null
    $longAtomicPath = Join-Path $longAtomicParent ('restore-validation-' + ('y' * 50) + '.json')
    Write-GssDriveBackupAtomicJson -Path $longAtomicPath -Value ([ordered]@{ status = 'Verified' })
    Assert-GssDriveBackupTest ((Read-GssDriveBackupJson -Path $longAtomicPath).status -eq 'Verified') 'Atomic JSON failed near the legacy Windows path-length boundary.'

    $metadataReadback = & $invokeScript `
        -Operation RecordMetadataReadback `
        -DriveFolderId 'test-drive-folder-id' `
        -ExpectedOwner 'owner@example.com' `
        -Shared:$false `
        -PermissionCount 1 `
        -SettingsPath $settingsPath `
        -OutputObject
    Assert-GssDriveBackupTest ($metadataReadback.Status -eq 'Recorded' -and $metadataReadback.Verification -eq 'cloud_metadata_readback_verified') 'Connector metadata readback was not recorded.'
    $sharedReadbackRefused = $false
    try {
        [void](& $invokeScript -Operation RecordMetadataReadback -DriveFolderId 'test-drive-folder-id' -ExpectedOwner 'owner@example.com' -Shared:$true -PermissionCount 2 -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $sharedReadbackRefused = $_.Exception.Message -match 'shared'
    }
    Assert-GssDriveBackupTest $sharedReadbackRefused 'Shared connector metadata was not refused.'

    $runId = 'gss-test-20260719-abcdef123456'
    $fingerprint = 'sha256:0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef'
    $runSummaryPath = Join-Path $resolvedTestRoot 'run-summary.json'
    $runSummary = [ordered]@{
        RunId = $runId
        RunFingerprint = $fingerprint
        Folder = $gssRoot
        CurrentWeekEnding = '2026-07-19'
        ProgramRelease = 'v1.0.0-test'
        TransactionArtifacts = @($transactionArtifact)
        ReleaseArchives = @($releasePath)
    }
    Write-GssDriveBackupAtomicJson -Path $runSummaryPath -Value $runSummary

    $inventoryResult = & $invokeScript -Operation Inventory -RunSummaryPath $runSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($inventoryResult.Status -eq 'Inventoried') 'Inventory operation did not succeed.'
    Assert-GssDriveBackupTest $inventoryResult.ContainsPersonalData 'PII classification was not propagated.'
    Assert-GssDriveBackupTest (@($inventoryResult.Inventory | Where-Object { $_.PortablePath -like '*quarantine*' }).Count -eq 0) 'Quarantine was included.'
    Assert-GssDriveBackupTest (@($inventoryResult.Inventory | Where-Object { $_.PortablePath -like '*test-output*' }).Count -eq 0) 'Test output was included.'
    Assert-GssDriveBackupTest (@($inventoryResult.Inventory | Where-Object { $_.PortablePath -like '*GSS Survey Workbook Automation*' }).Count -eq 0) 'Repository content was included.'
    Assert-GssDriveBackupTest (@($inventoryResult.Inventory | Where-Object { $_.Role -eq 'transaction_artifact' }).Count -eq 1) 'Explicit current transaction artifact was not included.'
    Assert-GssDriveBackupTest (@($inventoryResult.Inventory | Where-Object { $_.Role -eq 'release_archive' }).Count -eq 1) 'Release archive was not included.'
    Assert-GssDriveBackupTest (@($inventoryResult.Inventory | Where-Object { $_.PortablePath -like 'gss/05 Reference Materials/*' }).Count -eq 1) 'Production reference-material folder was omitted.'
    Assert-GssDriveBackupTest (@($inventoryResult.Inventory | Where-Object { $_.PortablePath -like 'gss/06 Exports and Images/*' }).Count -eq 1) 'Production exports-and-images folder was omitted.'
    Assert-GssDriveBackupTest (@($inventoryResult.Inventory | Where-Object { $_.PortablePath -like '*staging*' }).Count -eq 0) 'Staging package was included.'

    $prepared = & $invokeScript -Operation Prepare -RunSummaryPath $runSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($prepared.Status -eq 'Prepared') 'Prepare did not succeed.'
    Assert-GssDriveBackupTest ($prepared.VerificationLevel -eq 'drivefs_hash_verified') 'Prepare verification level is incorrect.'
    Assert-GssDriveBackupTest (Test-Path -LiteralPath $prepared.PreparedManifestPath -PathType Leaf) 'Prepared manifest is missing.'
    $preparedManifest = Read-GssDriveBackupJson -Path $prepared.PreparedManifestPath
    Assert-GssDriveBackupTest ([bool]$preparedManifest.data_classification.contains_personal_data) 'Prepared manifest omitted PII classification.'
    Assert-GssDriveBackupTest ($preparedManifest.data_classification.label -eq $script:GssDriveBackupClassificationLabel) 'Prepared manifest classification label is incorrect.'
    Assert-GssDriveBackupTest ($preparedManifest.scope.excluded -contains 'quarantine') 'Prepared manifest did not record exclusions.'

    Set-Content -LiteralPath (Join-Path $gssRoot '01 Main Workbook\GSS Score Trends - Main.xlsx') -Value 'post-apply-workbook' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $gssRoot '_automation_runs\qa\post.json') -Value '{"status":"READY"}' -Encoding UTF8

    $committed = & $invokeScript -Operation Finalize -RunSummaryPath $runSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($committed.Status -eq 'Committed') "Finalize did not commit: $($committed.Error)"
    Assert-GssDriveBackupTest (Test-Path -LiteralPath $committed.CommitReceiptPath -PathType Leaf) 'Commit receipt is missing.'
    $committedValidation = Test-GssCommittedBackupSnapshot -SnapshotPath $committed.SnapshotPath
    $liveEntry = @($committedValidation.Manifest.files | Where-Object portable_path -eq 'gss/01 Main Workbook/GSS Score Trends - Main.xlsx')
    Assert-GssDriveBackupTest ($liveEntry.Count -eq 1) 'Final manifest did not contain the live workbook.'
    $committedLivePath = Join-Path $committed.SnapshotPath $liveEntry[0].snapshot_path.Replace('/', '\')
    Assert-GssDriveBackupTest ((Get-Content -LiteralPath $committedLivePath -Raw).Trim() -eq 'post-apply-workbook') 'Finalize did not refresh the post-apply workbook.'
    $preparedLiveEntry = @($committedValidation.PreparedManifest.files | Where-Object portable_path -eq 'gss/01 Main Workbook/GSS Score Trends - Main.xlsx')
    Assert-GssDriveBackupTest ($preparedLiveEntry.Count -eq 1) 'Prepared manifest did not retain the live workbook.'
    $preparedLivePath = Join-Path $committed.SnapshotPath $preparedLiveEntry[0].snapshot_path.Replace('/', '\')
    Assert-GssDriveBackupTest ((Get-Content -LiteralPath $preparedLivePath -Raw).Trim() -eq 'pre-apply-workbook') 'Finalize overwrote the prepared pre-apply workbook bytes.'
    Assert-GssDriveBackupTest ($preparedLivePath -ne $committedLivePath) 'Prepared and final workbook payload paths are not distinct.'

    $chainHeadPath = Join-Path $driveRoot 'chain-head.json'
    Remove-Item -LiteralPath $chainHeadPath -Force
    $retry = & $invokeScript -Operation RetryFinalize -RunSummaryPath $runSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($retry.Status -eq 'Committed' -and $retry.Idempotent) 'RetryFinalize was not idempotent.'
    Assert-GssDriveBackupTest (Test-Path -LiteralPath $chainHeadPath -PathType Leaf) 'RetryFinalize did not repair the missing chain head.'

    $staleSummaryPath = Join-Path $resolvedTestRoot 'stale-run-summary.json'
    $staleSummary = [ordered]@{
        RunId = 'gss-test-stale-20260719-abcdef'
        RunFingerprint = 'sha256:stale-0123456789abcdef0123456789abcdef0123456789abcdef'
        Folder = $gssRoot
        CurrentWeekEnding = '2026-07-19'
        ProgramRelease = 'v1.0.0-test'
    }
    Write-GssDriveBackupAtomicJson -Path $staleSummaryPath -Value $staleSummary
    $stalePrepared = & $invokeScript -Operation Prepare -RunSummaryPath $staleSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($stalePrepared.Status -eq 'Prepared') 'Stale-chain test preparation failed.'

    $newerSummaryPath = Join-Path $resolvedTestRoot 'newer-run-summary.json'
    $newerSummary = [ordered]@{
        RunId = 'gss-test-newer-20260719-abcdef'
        RunFingerprint = 'sha256:newer-0123456789abcdef0123456789abcdef0123456789abcdef'
        Folder = $gssRoot
        CurrentWeekEnding = '2026-07-19'
        ProgramRelease = 'v1.0.0-test'
    }
    Write-GssDriveBackupAtomicJson -Path $newerSummaryPath -Value $newerSummary
    [void](& $invokeScript -Operation Prepare -RunSummaryPath $newerSummaryPath -SettingsPath $settingsPath -OutputObject)
    $newerCommitted = & $invokeScript -Operation Finalize -RunSummaryPath $newerSummaryPath -SettingsPath $settingsPath -OutputObject
    if ($newerCommitted.Status -ne 'Committed') {
        $newerPreparedEvidence = Read-GssDriveBackupJson -Path (Join-Path $driveRoot '.partial-gss-test-newer-20260719-abcdef\prepared-manifest.json')
        $currentHeadEvidence = Get-GssDriveBackupChainHead -RootPath $driveRoot
        throw "Newer-chain test snapshot did not commit: $($newerCommitted.Error) Expected='$($newerPreparedEvidence.prior_manifest_sha256)' Actual='$($currentHeadEvidence.ManifestSha256)'"
    }

    $staleFinalize = & $invokeScript -Operation Finalize -RunSummaryPath $staleSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($staleFinalize.Status -eq 'Blocked' -and $staleFinalize.BackupStatus -eq 'Blocked') 'Stale prepared chain was not returned as Blocked.'
    Assert-GssDriveBackupTest (Test-Path -LiteralPath (Join-Path $stalePrepared.PreparedPath 'backup-manifest.json') -PathType Leaf) 'Stale-chain test did not reach the existing-manifest retry case.'
    Assert-GssDriveBackupTest (-not (Test-Path -LiteralPath (Join-Path $stalePrepared.PreparedPath 'commit-receipt.json') -PathType Leaf)) 'Blocked stale chain wrote a commit receipt.'
    $staleRetry = & $invokeScript -Operation RetryFinalize -RunSummaryPath $staleSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($staleRetry.Status -eq 'Blocked') 'RetryFinalize skipped the chain check when backup-manifest already existed.'

    $aborted = & $invokeScript -Operation Abort -RunSummaryPath $staleSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($aborted.Status -eq 'Aborted' -and $aborted.EvidenceRetained) 'Abort did not retain the prepared evidence.'
    Assert-GssDriveBackupTest (Test-Path -LiteralPath $stalePrepared.PreparedManifestPath -PathType Leaf) 'Abort deleted prepared evidence.'
    $abortedRetry = & $invokeScript -Operation Abort -RunSummaryPath $staleSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($abortedRetry.Status -eq 'Aborted' -and $abortedRetry.Idempotent) 'Abort was not idempotent.'
    $abortCommittedRefused = $false
    try {
        [void](& $invokeScript -Operation Abort -RunSummaryPath $runSummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $abortCommittedRefused = $_.Exception.Message -match 'committed snapshot'
    }
    Assert-GssDriveBackupTest $abortCommittedRefused 'Abort did not refuse a committed snapshot.'

    $restored = Restore-GssDriveBackupForVerification -RunId $runId -SettingsPath $settingsPath -LocalAppDataPath $localAppData -Phase Final
    Assert-GssDriveBackupTest ($restored.Status -eq 'Verified') 'Final verify-only restore did not succeed.'
    Assert-GssDriveBackupTest (-not $restored.LiveWorkbookOverwritten) 'Verify-only restore claimed a live overwrite.'
    Assert-GssDriveBackupTest ($restored.Destination.StartsWith([System.IO.Path]::GetFullPath($localAppData), [System.StringComparison]::OrdinalIgnoreCase)) 'Verify-only restore escaped test LOCALAPPDATA.'
    $restoredFinalWorkbook = Join-Path $restored.Destination 'gss\01 Main Workbook\GSS Score Trends - Main.xlsx'
    Assert-GssDriveBackupTest ((Get-Content -LiteralPath $restoredFinalWorkbook -Raw).Trim() -eq 'post-apply-workbook') 'Final verify-only restore did not recover the post-apply workbook.'
    $restoredPrepared = Restore-GssDriveBackupForVerification -RunId $runId -SettingsPath $settingsPath -LocalAppDataPath $localAppData -Phase Prepared
    $restoredPreparedWorkbook = Join-Path $restoredPrepared.Destination 'gss\01 Main Workbook\GSS Score Trends - Main.xlsx'
    Assert-GssDriveBackupTest ((Get-Content -LiteralPath $restoredPreparedWorkbook -Raw).Trim() -eq 'pre-apply-workbook') 'Prepared verify-only restore did not recover the pre-apply workbook.'
    $hashOnlyDrillStatus = Get-GssDriveBackupRestoreDrillStatus -LocalAppDataPath $localAppData -AsOfDate ([datetime]'2026-07-23')
    Assert-GssDriveBackupTest ($hashOnlyDrillStatus.Status -eq 'Due') 'A hash-only VerifyRestore incorrectly satisfied the quarterly desktop Excel drill.'
    Assert-GssDriveBackupTest ($hashOnlyDrillStatus.HashOnlyVerificationCount -eq 2 -and -not $hashOnlyDrillStatus.HashOnlyVerificationSatisfiesQuarterlyDrill) 'Hash-only restore evidence was not exposed separately.'

    $excelReceiptPath = Join-Path $restored.Destination 'local-excel-validation-receipt.json'
    Write-GssDriveBackupAtomicJson -Path $excelReceiptPath -Value ([ordered]@{
        schema_version = 1
        Status = 'Passed'
        validated_at_utc = '2026-07-23T12:00:00Z'
    })
    $quarterlyReceiptPath = Join-Path $restored.Destination 'quarterly-restore-drill.json'
    Write-GssDriveBackupAtomicJson -Path $quarterlyReceiptPath -Value ([ordered]@{
        schema_version = 1
        operation = 'quarterly_verify_only_restore_drill'
        status = 'Passed'
        run_id = $runId
        completed_at_utc = '2026-07-23T12:00:00Z'
        drive_restore_receipt = $restored.ReceiptPath
        excel_validation_receipt = $excelReceiptPath
        live_workbook_overwritten = $false
    })
    $fullDrillStatus = Get-GssDriveBackupRestoreDrillStatus -LocalAppDataPath $localAppData -AsOfDate ([datetime]'2026-07-24')
    Assert-GssDriveBackupTest ($fullDrillStatus.Status -eq 'Current' -and $fullDrillStatus.LastRunId -eq $runId) 'Passed desktop Excel quarterly drill evidence was not recognized.'

    $retentionPath = Join-Path $resolvedTestRoot 'retention-report.json'
    $retention = Write-GssDriveBackupRetentionReport -AsOfDate ([datetime]'2026-07-23') -SettingsPath $settingsPath -ReportPath $retentionPath
    Assert-GssDriveBackupTest ($retention.Status -eq 'Reported' -and -not $retention.AutomaticDeletion) 'Retention report was not report-only.'
    Assert-GssDriveBackupTest (Test-Path -LiteralPath $committed.SnapshotPath -PathType Container) 'Retention reporting deleted a snapshot.'

    $capacity = Get-GssDriveBackupCapacityProjection -ProjectedWeeklyBytes 1024 -FreeBytes ([long](208 * 1024)) -SettingsPath $settingsPath
    Assert-GssDriveBackupTest ($capacity.WorstCaseWeeklyLoadsAvailable -eq 104) 'Capacity projection load calculation is incorrect.'
    Assert-GssDriveBackupTest ($capacity.StructuralRedesignStatus -eq 'Deferred') 'Capacity threshold should be deferred at 104 loads.'

    $missingSettingsPath = Join-Path $resolvedTestRoot 'missing-drive-settings.json'
    $missingSettings = [ordered]@{
        schema_version = 1
        drive_root_path = (Join-Path $resolvedTestRoot 'missing-drive')
        drive_folder_id = 'test-drive-folder-id'
        expected_owner = 'owner@example.com'
        marker_id = $commission.MarkerId
        verification_level = 'drivefs_hash_verified'
        retention = [ordered]@{ weekly = 13; monthly = 12 }
        require_before_apply = $true
    }
    Write-GssDriveBackupAtomicJson -Path $missingSettingsPath -Value $missingSettings
    $blocked = $false
    try {
        [void](Get-GssDriveBackupRootContext -SettingsPath $missingSettingsPath)
    }
    catch {
        $blocked = $_.Exception.Message -match 'No fallback destination'
    }
    Assert-GssDriveBackupTest $blocked 'Unavailable Drive root did not fail closed with no fallback.'

    Write-Output "GSS Drive backup tests passed: commissioned, curated, prepared, finalized, retried, retention-reported, capacity-projected, and verify-restored $($committed.FileCount) file(s)."
}
finally {
    if (Test-Path -LiteralPath $resolvedTestRoot -PathType Container) {
        $verifiedCleanupPath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $resolvedTestRoot).Path)
        if (-not $verifiedCleanupPath.StartsWith("$systemTemp\", [System.StringComparison]::OrdinalIgnoreCase) -or
            -not ([System.IO.Path]::GetFileName($verifiedCleanupPath)).StartsWith('gss-drive-backup-test-', [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing unsafe recursive test cleanup: $verifiedCleanupPath"
        }
        Remove-Item -LiteralPath $verifiedCleanupPath -Recurse -Force
    }
}
