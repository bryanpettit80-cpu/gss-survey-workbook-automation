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
$testRoot = Join-Path $systemTemp ("gdbt-$([guid]::NewGuid().ToString('N').Substring(0, 12))")
$resolvedTestRoot = [System.IO.Path]::GetFullPath($testRoot)
if (-not $resolvedTestRoot.StartsWith("$systemTemp\", [System.StringComparison]::OrdinalIgnoreCase) -or
    -not ([System.IO.Path]::GetFileName($resolvedTestRoot)).StartsWith('gdbt-', [System.StringComparison]::OrdinalIgnoreCase)) {
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
    (Get-Item -LiteralPath $longAtomicPath).LastWriteTimeUtc = [datetime]'2020-01-02T03:04:05Z'
    $unchangedTimestamp = (Get-Item -LiteralPath $longAtomicPath).LastWriteTimeUtc
    Write-GssDriveBackupAtomicJson -Path $longAtomicPath -Value ([ordered]@{ status = 'Verified' })
    Assert-GssDriveBackupTest `
        ((Get-Item -LiteralPath $longAtomicPath).LastWriteTimeUtc -eq $unchangedTimestamp) `
        'Atomic JSON rewrote an already-identical synced-file payload.'

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

    $recoveredArchiveRoot = Join-Path $gssRoot '03 Uploaded Survey Workbooks\Archive - Previous Uploads\Recovered Historical Detail\FY26'
    New-Item -ItemType Directory -Path $recoveredArchiveRoot -Force | Out-Null
    $recoveredTemporaryOne = Join-Path $recoveredArchiveRoot '.fixture-fw1.tmp'
    $recoveredTemporaryTwo = Join-Path $recoveredArchiveRoot '.fixture-fw2.tmp'
    Set-Content -LiteralPath $recoveredTemporaryOne -Value 'recovered historical workbook one' -Encoding UTF8
    Set-Content -LiteralPath $recoveredTemporaryTwo -Value 'recovered historical workbook two' -Encoding UTF8
    $recoveredHashOne = Get-GssDriveBackupSha256 -Path $recoveredTemporaryOne
    $recoveredHashTwo = Get-GssDriveBackupSha256 -Path $recoveredTemporaryTwo
    $recoveredPathOne = Join-Path $recoveredArchiveRoot "FY26-FW1-$recoveredHashOne.xlsx"
    $recoveredPathTwo = Join-Path $recoveredArchiveRoot "FY26-FW2-$recoveredHashTwo.xlsx"
    Move-Item -LiteralPath $recoveredTemporaryOne -Destination $recoveredPathOne
    Move-Item -LiteralPath $recoveredTemporaryTwo -Destination $recoveredPathTwo
    $recoveredRelativeOne = (Get-GssDriveBackupRelativePath -Path $recoveredPathOne -Root $gssRoot).Replace('\', '/')
    $recoveredRelativeTwo = (Get-GssDriveBackupRelativePath -Path $recoveredPathTwo -Root $gssRoot).Replace('\', '/')
    $recoveredBytesOne = [long](Get-Item -LiteralPath $recoveredPathOne).Length
    $recoveredBytesTwo = [long](Get-Item -LiteralPath $recoveredPathTwo).Length

    $ledgerArtifact = Join-Path $gssRoot '_automation_runs\state\gss_feedback_first_seen.json'
    $responseHashOne = '1111111111111111111111111111111111111111111111111111111111111111'
    $responseHashTwo = '2222222222222222222222222222222222222222222222222222222222222222'
    $responseHashThree = '3333333333333333333333333333333333333333333333333333333333333333'
    Write-GssDriveBackupAtomicJson -Path $ledgerArtifact -Value ([ordered]@{
        schema_version = 'gss-feedback-first-seen/v1'
        entries = @(
            [ordered]@{ response_hash = $responseHashOne; first_seen_reporting_date = '2025-05-31'; first_seen_package_id = 'historical-recovery:test' }
            [ordered]@{ response_hash = $responseHashTwo; first_seen_reporting_date = '2025-05-31'; first_seen_package_id = 'historical-recovery:test' }
            [ordered]@{ response_hash = $responseHashThree; first_seen_reporting_date = '2025-06-07'; first_seen_package_id = 'historical-recovery:test' }
        )
    })
    $ledgerHash = Get-GssDriveBackupSha256 -Path $ledgerArtifact

    $manifestStagingPath = Join-Path $resolvedTestRoot 'recovery-manifest-staging.json'
    $recoveryManifest = [ordered]@{
        schema_version = 'gss-historical-recovery/v1'
        fiscal_year = 'FY26'
        created_at_utc = '2026-07-23T12:00:00Z'
        sources = @(
            [ordered]@{
                source_kind = 'gmail_attachment'
                drive_file_id = ''
                gmail_message_ids = @('message-one')
                subject_week = 'FY26 FW1'
                source_report_week = 'FY26 FW1'
                assignment_basis = 'paired_rolling_terminal_week'
                visit_date_start = '2025-05-25'
                visit_date_end = '2025-05-31'
                row_count = 2
                byte_size = $recoveredBytesOne
                sha256 = $recoveredHashOne
                response_set_sha256 = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
                destination_path = $recoveredRelativeOne
                validation = [ordered]@{
                    detail_schema_valid = $true
                    response_identity_version = 'gss-feedback-response-identity/v1'
                    duplicate_response_count = 0
                }
            }
            [ordered]@{
                source_kind = 'gmail_attachment'
                drive_file_id = ''
                gmail_message_ids = @('message-two')
                subject_week = 'FY26 FW2'
                source_report_week = 'FY26 FW2'
                assignment_basis = 'paired_rolling_terminal_week'
                visit_date_start = '2025-06-01'
                visit_date_end = '2025-06-07'
                row_count = 1
                byte_size = $recoveredBytesTwo
                sha256 = $recoveredHashTwo
                response_set_sha256 = 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb'
                destination_path = $recoveredRelativeTwo
                validation = [ordered]@{
                    detail_schema_valid = $true
                    response_identity_version = 'gss-feedback-response-identity/v1'
                    duplicate_response_count = 0
                }
            }
        )
    }
    Write-GssDriveBackupAtomicJson -Path $manifestStagingPath -Value $recoveryManifest
    $recoveryTransactionHash = Get-GssDriveBackupSha256 -Path $manifestStagingPath
    $recoveryTransactionRoot = Join-Path $gssRoot "_automation_runs\historical-recovery\$recoveryTransactionHash"
    New-Item -ItemType Directory -Path $recoveryTransactionRoot -Force | Out-Null
    $recoveryManifestArtifact = Join-Path $recoveryTransactionRoot 'recovery-manifest.json'
    Move-Item -LiteralPath $manifestStagingPath -Destination $recoveryManifestArtifact

    $recoveryArtifact = Join-Path $recoveryTransactionRoot 'transaction-receipt.json'
    $receiptFiles = @(
        [ordered]@{
            source_index = 1
            sha256 = $recoveredHashOne
            byte_size = $recoveredBytesOne
            row_count = 2
            response_set_sha256 = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
            destination_path = $recoveredRelativeOne
            destination_full_path = $recoveredPathOne
            destination_preexisting = $false
            published_by_transaction = $true
        }
        [ordered]@{
            source_index = 2
            sha256 = $recoveredHashTwo
            byte_size = $recoveredBytesTwo
            row_count = 1
            response_set_sha256 = 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb'
            destination_path = $recoveredRelativeTwo
            destination_full_path = $recoveredPathTwo
            destination_preexisting = $false
            published_by_transaction = $true
        }
    )
    Write-GssDriveBackupAtomicJson -Path $recoveryArtifact -Value ([ordered]@{
        schema_version = 'gss-historical-recovery-receipt/v1'
        classification = 'CONTAINS PERSONAL DATA - RESTRICTED'
        contains_personal_data = $true
        manifest_sha256 = $recoveryTransactionHash
        manifest_snapshot_path = $recoveryManifestArtifact
        transaction_id = "historical-recovery:$recoveryTransactionHash"
        state = 'Committed'
        ledger_path = $ledgerArtifact
        ledger_sha256_after = $ledgerHash
        planned_entries = @(
            [ordered]@{ response_hash = $responseHashOne; visit_date = '2025-05-30' }
            [ordered]@{ response_hash = $responseHashTwo; visit_date = '2025-05-31' }
            [ordered]@{ response_hash = $responseHashThree; visit_date = '2025-06-07' }
        )
        inserted_response_hashes = @($responseHashOne, $responseHashTwo, $responseHashThree)
        files = $receiptFiles
        published_file_count = 2
        error = ''
    })

    $recoveryQaArtifact = Join-Path $recoveryTransactionRoot 'recovery-qa.json'
    Write-GssDriveBackupAtomicJson -Path $recoveryQaArtifact -Value ([ordered]@{
        schema_version = 'gss-historical-recovery-qa/v1'
        status = 'Passed'
        manifest_sha256 = $recoveryTransactionHash
        transaction_id = "historical-recovery:$recoveryTransactionHash"
        source_count = 2
        recovered_file_count = 2
        row_count = 3
        unique_response_count = 3
        inserted_response_count = 3
        published_file_count = 2
        ledger_entry_count_after = 3
        controls = [ordered]@{
            manifest_hash_verified = $true
            receipt_committed = $true
            destinations_verified = $true
            ledger_hash_verified = $true
            no_unrelated_recovered_xlsx = $true
            live_workbook_unchanged = $true
            email_package_unchanged = $true
            automatic_sending_disabled = $true
            scheduled_task_disabled = $true
            contains_row_level_data = $false
        }
        generated_at_utc = '2026-07-23T12:00:00Z'
    })

    $recoverySummaryPath = Join-Path $recoveryTransactionRoot 'drive-recovery-summary.json'
    $recoveryBackupInventory = @(
        [ordered]@{
            SourcePath = $recoveredPathOne
            PortablePath = "recovery/recovered-detail/$([System.IO.Path]::GetFileName($recoveredPathOne))"
            Role = 'recovered_historical_detail'
            Classification = 'restricted_personal_data'
        }
        [ordered]@{
            SourcePath = $recoveredPathTwo
            PortablePath = "recovery/recovered-detail/$([System.IO.Path]::GetFileName($recoveredPathTwo))"
            Role = 'recovered_historical_detail'
            Classification = 'restricted_personal_data'
        }
        [ordered]@{
            SourcePath = $ledgerArtifact
            PortablePath = 'recovery/state/gss_feedback_first_seen.json'
            Role = 'recovery_ledger'
            Classification = 'restricted_operational'
        }
        [ordered]@{
            SourcePath = $recoveryManifestArtifact
            PortablePath = 'recovery/evidence/recovery-manifest.json'
            Role = 'recovery_manifest'
            Classification = 'restricted_operational'
        }
        [ordered]@{
            SourcePath = $recoveryArtifact
            PortablePath = 'recovery/evidence/transaction-receipt.json'
            Role = 'recovery_receipt'
            Classification = 'restricted_operational'
        }
        [ordered]@{
            SourcePath = $recoveryQaArtifact
            PortablePath = 'recovery/evidence/recovery-qa.json'
            Role = 'recovery_qa'
            Classification = 'restricted_operational'
        }
        [ordered]@{
            SourcePath = $recoverySummaryPath
            PortablePath = 'recovery/evidence/drive-recovery-summary.json'
            Role = 'recovery_run_summary'
            Classification = 'restricted_operational'
        }
    )
    $recoverySummary = [ordered]@{
        schema_version = 'gss-recovery-drive-summary/v1'
        RunId = 'gss-recovery-20260719-abcdef'
        RunFingerprint = "sha256:$recoveryTransactionHash"
        Folder = $gssRoot
        CurrentWeekEnding = '2026-07-19'
        ProgramRelease = 'v1.1.0-test'
        SnapshotPurpose = 'RecoveryOnly'
        BackupInventory = $recoveryBackupInventory
    }
    Write-GssDriveBackupAtomicJson -Path $recoverySummaryPath -Value $recoverySummary
    $recoveryInventory = & $invokeScript -Operation Inventory -RunSummaryPath $recoverySummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($recoveryInventory.SnapshotPurpose -eq 'RecoveryOnly') 'Recovery-only purpose was not propagated.'
    Assert-GssDriveBackupTest ($recoveryInventory.InventoryMode -eq 'RecoveryOnly') 'Recovery-only inventory mode was not propagated.'
    Assert-GssDriveBackupTest ($recoveryInventory.FileCount -eq 7) 'Recovery-only inventory was not the exact two-XLSX plus five-evidence transaction.'
    Assert-GssDriveBackupTest (@($recoveryInventory.Inventory | Where-Object Role -eq 'recovered_historical_detail').Count -eq 2) 'Recovery-only inventory omitted a manifest destination.'

    $malformedQa = Read-GssDriveBackupJson -Path $recoveryQaArtifact
    $malformedQa.controls.automatic_sending_disabled = 'false'
    Write-GssDriveBackupAtomicJson -Path $recoveryQaArtifact -Value $malformedQa
    $stringQaBooleanRefused = $false
    try {
        [void](& $invokeScript -Operation Inventory -RunSummaryPath $recoverySummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $stringQaBooleanRefused = $_.Exception.Message -match 'must be a JSON boolean'
    }
    Assert-GssDriveBackupTest $stringQaBooleanRefused 'Recovery QA accepted a truthy string in place of a JSON boolean.'
    $malformedQa.controls.automatic_sending_disabled = $true
    Write-GssDriveBackupAtomicJson -Path $recoveryQaArtifact -Value $malformedQa

    $malformedReceipt = Read-GssDriveBackupJson -Path $recoveryArtifact
    $malformedReceipt.files[0].published_by_transaction = 'false'
    Write-GssDriveBackupAtomicJson -Path $recoveryArtifact -Value $malformedReceipt
    $stringReceiptBooleanRefused = $false
    try {
        [void](& $invokeScript -Operation Inventory -RunSummaryPath $recoverySummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $stringReceiptBooleanRefused = $_.Exception.Message -match 'must be JSON booleans'
    }
    Assert-GssDriveBackupTest $stringReceiptBooleanRefused 'Recovery receipt accepted a truthy string in place of a JSON boolean.'
    $malformedReceipt.files[0].published_by_transaction = $true
    Write-GssDriveBackupAtomicJson -Path $recoveryArtifact -Value $malformedReceipt

    $outsideArtifact = Join-Path $resolvedTestRoot 'outside-recovery-secret.txt'
    Set-Content -LiteralPath $outsideArtifact -Value 'must not enter RecoveryOnly snapshot' -Encoding UTF8
    $outsideSummaryPath = Join-Path $resolvedTestRoot 'outside-recovery-summary.json'
    $outsideSummary = [ordered]@{
        RunId = 'gss-recovery-outside-20260719'
        RunFingerprint = 'sha256:outside-0123456789abcdef0123456789abcdef0123456789abcdef'
        Folder = $gssRoot
        CurrentWeekEnding = '2026-07-19'
        ProgramRelease = 'v1.1.0-test'
        SnapshotPurpose = 'RecoveryOnly'
        BackupInventory = @(
            [ordered]@{
                SourcePath = $outsideArtifact
                PortablePath = 'recovery/outside-recovery-secret.txt'
                Role = 'recovery_receipt'
                Classification = 'restricted_operational'
            }
        )
    }
    Write-GssDriveBackupAtomicJson -Path $outsideSummaryPath -Value $outsideSummary
    $outsideRefused = $false
    try {
        [void](& $invokeScript -Operation Inventory -RunSummaryPath $outsideSummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $outsideRefused = $_.Exception.Message -match 'outside the expected root'
    }
    Assert-GssDriveBackupTest $outsideRefused 'RecoveryOnly inventory accepted an arbitrary file outside the GSS root.'

    $misrepresentedEvidence = Join-Path $recoveryTransactionRoot 'guest-raw-detail.xlsx'
    Set-Content -LiteralPath $misrepresentedEvidence -Value 'raw guest detail' -Encoding UTF8
    $misrepresentedSummaryPath = Join-Path $resolvedTestRoot 'misrepresented-recovery-summary.json'
    $misrepresentedSummary = [ordered]@{
        RunId = 'gss-recovery-misrepresented-20260719'
        RunFingerprint = 'sha256:misrepresented-0123456789abcdef0123456789abcdef0123456789abcdef'
        Folder = $gssRoot
        CurrentWeekEnding = '2026-07-19'
        ProgramRelease = 'v1.1.0-test'
        SnapshotPurpose = 'RecoveryOnly'
        BackupInventory = @(
            [ordered]@{
                SourcePath = $misrepresentedEvidence
                PortablePath = 'recovery/evidence/recovery-manifest.json'
                Role = 'recovery_manifest'
                Classification = 'restricted_operational'
            }
        )
    }
    Write-GssDriveBackupAtomicJson -Path $misrepresentedSummaryPath -Value $misrepresentedSummary
    $misrepresentedRefused = $false
    try {
        [void](& $invokeScript -Operation Inventory -RunSummaryPath $misrepresentedSummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $misrepresentedRefused = $_.Exception.Message -match 'exact filename'
    }
    Assert-GssDriveBackupTest $misrepresentedRefused 'A raw-detail XLSX was accepted while masquerading as recovery-manifest evidence.'

    $junctionTarget = Join-Path $resolvedTestRoot 'junction-target'
    New-Item -ItemType Directory -Path $junctionTarget -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $junctionTarget 'junction-secret.xlsx') -Value 'outside GSS through junction' -Encoding UTF8
    $recoveredArchiveRoot = Join-Path $gssRoot '03 Uploaded Survey Workbooks\Archive - Previous Uploads\Recovered Historical Detail'
    New-Item -ItemType Directory -Path $recoveredArchiveRoot -Force | Out-Null
    $junctionPath = Join-Path $recoveredArchiveRoot 'linked'
    New-Item -ItemType Junction -Path $junctionPath -Target $junctionTarget | Out-Null
    $junctionSummaryPath = Join-Path $resolvedTestRoot 'junction-recovery-summary.json'
    $junctionSummary = [ordered]@{
        RunId = 'gss-recovery-junction-20260719'
        RunFingerprint = 'sha256:junction-0123456789abcdef0123456789abcdef0123456789abcdef'
        Folder = $gssRoot
        CurrentWeekEnding = '2026-07-19'
        ProgramRelease = 'v1.1.0-test'
        SnapshotPurpose = 'RecoveryOnly'
        BackupInventory = @(
            [ordered]@{
                SourcePath = (Join-Path $junctionPath 'junction-secret.xlsx')
                PortablePath = 'recovery/recovered-detail/junction-secret.xlsx'
                Role = 'recovered_historical_detail'
                Classification = 'restricted_personal_data'
            }
        )
    }
    Write-GssDriveBackupAtomicJson -Path $junctionSummaryPath -Value $junctionSummary
    $junctionRefused = $false
    try {
        [void](& $invokeScript -Operation Inventory -RunSummaryPath $junctionSummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $junctionRefused = $_.Exception.Message -match 'symbolic link or junction'
    }
    Assert-GssDriveBackupTest $junctionRefused 'RecoveryOnly inventory traversed a junction to a file outside the GSS root.'

    $longPathSnapshotRoot = Join-Path $driveRoot ('.partial-' + [guid]::NewGuid().ToString())
    $longPathPortable = 'gss/03 Uploaded Survey Workbooks/Archive - Previous Uploads/Recovered Historical Detail/FY26/' + [System.IO.Path]::GetFileName($recoveredPathOne)
    $uncompactedDestination = Join-Path $longPathSnapshotRoot ('prepared-payload/' + $longPathPortable).Replace('/', '\')
    Assert-GssDriveBackupTest ($uncompactedDestination.Length -ge 260) 'Long-path regression fixture did not cross the Windows MAX_PATH boundary.'
    $longPathCopy = @(Copy-GssDriveBackupInventory -Inventory @(
        [pscustomobject]@{
            SourcePath = $recoveredPathOne
            PortablePath = $longPathPortable
            Role = 'recovered_historical_detail'
            Classification = 'restricted_personal_data'
        }
    ) -SnapshotDirectory $longPathSnapshotRoot -PayloadPrefix 'prepared-payload')
    Assert-GssDriveBackupTest ($longPathCopy.Count -eq 1) 'Long-path backup did not return exactly one manifest entry.'
    Assert-GssDriveBackupTest ($longPathCopy[0].snapshot_path -match '^prepared-payload/long-path/[0-9a-f]{64}\.xlsx$') 'Long-path backup did not use a deterministic compact snapshot path.'
    $longPathCopiedFile = Join-Path $longPathSnapshotRoot $longPathCopy[0].snapshot_path.Replace('/', '\')
    Assert-GssDriveBackupTest ($longPathCopiedFile.Length -lt 248) 'Compacted snapshot path still approaches the Windows MAX_PATH boundary.'
    Assert-GssDriveBackupTest (Test-Path -LiteralPath $longPathCopiedFile -PathType Leaf) 'Long-path backup did not promote the copied file.'
    Assert-GssDriveBackupTest ((Get-GssDriveBackupSha256 -Path $longPathCopiedFile) -eq (Get-GssDriveBackupSha256 -Path $recoveredPathOne)) 'Long-path backup changed the copied file bytes.'

    # A path can fit in .partial-<run-id> yet fail after the directory is moved
    # beneath snapshots\YYYY\YYYY-MM. Budgeting must include that promotion.
    $promotionStageRoot = Join-Path $driveRoot '.partial-edge'
    New-Item -ItemType Directory -Path $promotionStageRoot -Force | Out-Null
    $promotionPrefix = 'prepared-payload/'
    $promotionLeafLength = 246 - (Join-Path $promotionStageRoot $promotionPrefix.Replace('/', '\')).Length
    $promotionPortable = ('p' * ($promotionLeafLength - 4)) + '.txt'
    $promotionStagedPath = Join-Path $promotionStageRoot ($promotionPrefix + $promotionPortable).Replace('/', '\')
    $promotionFinalRoot = Join-Path $driveRoot 'snapshots\2026\2026-07\edge'
    $promotionFinalPath = Join-Path $promotionFinalRoot ($promotionPrefix + $promotionPortable).Replace('/', '\')
    Assert-GssDriveBackupTest ($promotionStagedPath.Length -in @(246, 247)) 'Promotion regression fixture did not reach the 246-247 character staging edge.'
    Assert-GssDriveBackupTest ($promotionFinalPath.Length -ge 248) 'Promotion regression fixture did not exceed the safe path budget after promotion.'
    $promotionCopy = @(Copy-GssDriveBackupInventory -Inventory @(
        [pscustomobject]@{ SourcePath = $recoveredPathOne; PortablePath = $promotionPortable; Role = 'test'; Classification = 'restricted_operational' }
    ) -SnapshotDirectory $promotionStageRoot -PayloadPrefix 'prepared-payload' -PathBudgetDirectories @($promotionFinalRoot))
    Assert-GssDriveBackupTest ($promotionCopy[0].snapshot_path -match '^prepared-payload/long-path/[0-9a-f]{64}\.txt$') 'A staging-safe path was not compacted for its longer promoted destination.'
    $promotionParent = Split-Path -Parent $promotionFinalRoot
    New-Item -ItemType Directory -Path $promotionParent -Force | Out-Null
    Move-Item -LiteralPath $promotionStageRoot -Destination $promotionFinalRoot
    $promotedEdgeFile = Join-Path $promotionFinalRoot $promotionCopy[0].snapshot_path.Replace('/', '\')
    Assert-GssDriveBackupTest (Test-Path -LiteralPath $promotedEdgeFile -PathType Leaf) 'The compacted 246-247 character staging fixture did not survive promotion.'
    Assert-GssDriveBackupTest ((Get-GssDriveBackupSha256 -Path $promotedEdgeFile) -eq (Get-GssDriveBackupSha256 -Path $recoveredPathOne)) 'Promotion changed the compacted edge fixture bytes.'

    $longExtension = '.' + ('extension' * 24)
    $longExtensionPortable = 'deep/' + ('q' * 20) + $longExtension
    $longExtensionSource = Join-Path $resolvedTestRoot 'source-long-extension.bin'
    Set-Content -LiteralPath $longExtensionSource -Value 'long extension bytes' -Encoding UTF8
    $boundedExtensionCandidate = Get-GssDriveBackupCompactRelativePath -PortablePath $longExtensionPortable -Prefix 'extension-payload'
    $extensionlessCandidate = Get-GssDriveBackupCompactRelativePath -PortablePath $longExtensionPortable -Prefix 'extension-payload' -OmitExtension
    $boundedRelativeLength = $boundedExtensionCandidate.Replace('/', '\').Length
    $extensionBudgetRootLength = 248 - 1 - $boundedRelativeLength
    $extensionBudgetRootPrefix = Join-Path $resolvedTestRoot 'extension-budget-'
    $extensionBudgetPadding = $extensionBudgetRootLength - $extensionBudgetRootPrefix.Length
    Assert-GssDriveBackupTest ($extensionBudgetPadding -gt 0) 'Long-extension budget fixture cannot be calibrated beneath the test root.'
    $extensionBudgetRoot = $extensionBudgetRootPrefix + ('e' * $extensionBudgetPadding)
    $boundedExtensionDestination = Join-Path $extensionBudgetRoot $boundedExtensionCandidate.Replace('/', '\')
    $extensionlessDestination = Join-Path $extensionBudgetRoot $extensionlessCandidate.Replace('/', '\')
    Assert-GssDriveBackupTest ($boundedExtensionDestination.Length -eq 248 -and $extensionlessDestination.Length -lt 248) 'Long-extension budget fixture did not require whole-extension omission.'
    $longExtensionCopy = @(Copy-GssDriveBackupInventory -Inventory @(
        [pscustomobject]@{ SourcePath = $longExtensionSource; PortablePath = $longExtensionPortable; Role = 'test'; Classification = 'restricted_operational' }
    ) -SnapshotDirectory $extensionBudgetRoot -PayloadPrefix 'extension-payload')
    $longExtensionName = [System.IO.Path]::GetFileName([string]$longExtensionCopy[0].snapshot_path)
    $longExtensionDestination = Join-Path $extensionBudgetRoot ([string]$longExtensionCopy[0].snapshot_path).Replace('/', '\')
    Assert-GssDriveBackupTest ($longExtensionCopy[0].snapshot_path -eq $extensionlessCandidate -and $longExtensionCopy[0].snapshot_path -match '^extension-payload/long-path/[0-9a-f]{64}$') 'Long-extension backup did not omit the whole bounded extension when the full path required it.'
    Assert-GssDriveBackupTest ($longExtensionName.Length -le (64 + $script:GssDriveBackupMaxRetainedExtensionLength)) 'Compacted filename retained an unbounded extension.'
    Assert-GssDriveBackupTest ($longExtensionName.Length -le 255) 'Compacted filename component exceeded the Windows limit.'
    Assert-GssDriveBackupTest ($longExtensionDestination.Length -lt 248 -and (Test-Path -LiteralPath $longExtensionDestination -PathType Leaf)) 'Long-extension backup did not omit extension retention to fit the full Windows path budget.'
    $longExtensionVariant = Get-GssDriveBackupCompactRelativePath -PortablePath ($longExtensionPortable + 'x') -Prefix 'extension-payload' -OmitExtension
    Assert-GssDriveBackupTest ($longExtensionCopy[0].snapshot_path -ne $longExtensionVariant) 'Extension truncation discarded the full portable-path digest collision protection.'

    $maxRunId = 'r' + ('x' * 127)
    $overBudgetRoot = Join-Path $driveRoot (('deep-' + ('d' * 90)) + "\snapshots\2026\2026-07\$maxRunId")
    $compactBudgetRefused = $false
    try {
        [void](Copy-GssDriveBackupInventory -Inventory @(
            [pscustomobject]@{ SourcePath = $recoveredPathOne; PortablePath = $longPathPortable; Role = 'test'; Classification = 'restricted_operational' }
        ) -SnapshotDirectory $longPathSnapshotRoot -PayloadPrefix 'prepared-payload' -PathBudgetDirectories @($overBudgetRoot))
    }
    catch {
        $compactBudgetRefused = $_.Exception.Message -match 'Compacted snapshot destination still exceeds the safe Windows path budget before copy'
    }
    Assert-GssDriveBackupTest $compactBudgetRefused 'An over-budget compact destination with a valid 128-character RunId was not rejected before copy.'

    # Keep the unrelated full-snapshot/restore fixture below the legacy
    # Windows MAX_PATH boundary. RecoveryOnly itself still exercises the
    # production collision-proof archive names.
    $recoveredHoldingRoot = Join-Path $resolvedTestRoot 'hold'
    New-Item -ItemType Directory -Path $recoveredHoldingRoot -Force | Out-Null
    Move-Item -LiteralPath $recoveredPathOne -Destination $recoveredHoldingRoot
    Move-Item -LiteralPath $recoveredPathTwo -Destination $recoveredHoldingRoot

    try {
        $prepared = & $invokeScript -Operation Prepare -RunSummaryPath $runSummaryPath -SettingsPath $settingsPath -OutputObject
    }
    catch {
        throw "Full snapshot preparation fixture failed: $($_.Exception.Message)"
    }
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

    Move-Item -LiteralPath (Join-Path $recoveredHoldingRoot ([System.IO.Path]::GetFileName($recoveredPathOne))) -Destination $recoveredPathOne
    Move-Item -LiteralPath (Join-Path $recoveredHoldingRoot ([System.IO.Path]::GetFileName($recoveredPathTwo))) -Destination $recoveredPathTwo
    try {
        $recoveryPrepared = & $invokeScript -Operation Prepare -RunSummaryPath $recoverySummaryPath -SettingsPath $settingsPath -OutputObject
    }
    catch {
        throw "RecoveryOnly snapshot preparation fixture failed: $($_.Exception.Message)"
    }
    $recoveryPreparedManifest = Read-GssDriveBackupJson -Path $recoveryPrepared.PreparedManifestPath
    Assert-GssDriveBackupTest ($recoveryPreparedManifest.snapshot_purpose -eq 'RecoveryOnly') 'Prepared recovery snapshot omitted its purpose.'
    Assert-GssDriveBackupTest ($recoveryPreparedManifest.scope.inventory_mode -eq 'RecoveryOnly') 'Prepared recovery snapshot omitted its narrow inventory mode.'
    $originalQaCopy = Join-Path $resolvedTestRoot 'recovery-qa-original.json'
    Copy-Item -LiteralPath $recoveryQaArtifact -Destination $originalQaCopy
    $originalQaHash = Get-GssDriveBackupSha256 -Path $recoveryQaArtifact
    $mutatedQa = Read-GssDriveBackupJson -Path $recoveryQaArtifact
    $mutatedQa.generated_at_utc = '2026-07-23T12:01:00Z'
    Write-GssDriveBackupAtomicJson -Path $recoveryQaArtifact -Value $mutatedQa
    $changedPrepareRefused = $false
    $changedPrepareError = ''
    try {
        [void](& $invokeScript -Operation Prepare -RunSummaryPath $recoverySummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $changedPrepareError = $_.Exception.Message
        $changedPrepareRefused = $changedPrepareError -match 'changed (sha256|byte_size)|does not match the prepared inventory'
    }
    finally {
        Copy-Item -LiteralPath $originalQaCopy -Destination $recoveryQaArtifact -Force
    }
    Assert-GssDriveBackupTest $changedPrepareRefused "Idempotent RecoveryOnly preparation accepted a changed inventory. Error='$changedPrepareError'"
    Assert-GssDriveBackupTest ((Get-GssDriveBackupSha256 -Path $recoveryQaArtifact) -eq $originalQaHash) 'Recovery QA fixture was not restored after the mutation regression.'
    $widenedRecoverySummaryPath = Join-Path $resolvedTestRoot 'widened-recovery-summary.json'
    $widenedRecoverySummary = [ordered]@{
        RunId = $recoverySummary.RunId
        RunFingerprint = $recoverySummary.RunFingerprint
        Folder = $gssRoot
        CurrentWeekEnding = '2026-07-19'
        ProgramRelease = 'v1.1.0-test'
        SnapshotPurpose = 'WorkbookTransaction'
    }
    Write-GssDriveBackupAtomicJson -Path $widenedRecoverySummaryPath -Value $widenedRecoverySummary
    $widenedFinalize = & $invokeScript -Operation Finalize -RunSummaryPath $widenedRecoverySummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($widenedFinalize.Status -eq 'PendingFinalize' -and $widenedFinalize.Error -match 'identity') 'RecoveryOnly finalization accepted a widened Full summary.'
    $recoveryCommitted = & $invokeScript -Operation Finalize -RunSummaryPath $recoverySummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($recoveryCommitted.Status -eq 'Committed') "Recovery-only finalize did not commit: $($recoveryCommitted.Error)"
    $recoveryValidation = Test-GssCommittedBackupSnapshot -SnapshotPath $recoveryCommitted.SnapshotPath
    Assert-GssDriveBackupTest ($recoveryValidation.Manifest.snapshot_purpose -eq 'RecoveryOnly') 'Committed recovery snapshot omitted its purpose.'
    Assert-GssDriveBackupTest (@($recoveryValidation.Manifest.files).Count -eq 7) 'Committed recovery snapshot was not the exact closed recovery transaction.'
    $longRestoreLocalAppData = Join-Path $resolvedTestRoot ('restore-root-' + ('r' * 30))
    $recoveryRestore = Restore-GssDriveBackupForVerification -RunId $recoverySummary.RunId -SettingsPath $settingsPath -LocalAppDataPath $longRestoreLocalAppData -Phase Final
    $recoveryRestoreReceipt = Read-GssDriveBackupJson -Path $recoveryRestore.ReceiptPath
    $compactRestoreFiles = @($recoveryRestoreReceipt.files | Where-Object { $_.restored_path -match '^r/long-path/[0-9a-f]{64}(?:\.[^/]+)?$' })
    Assert-GssDriveBackupTest ($recoveryRestore.Status -eq 'Verified' -and $recoveryRestore.FileCount -eq 7) 'Long-path RecoveryOnly verify-only restore did not verify every file.'
    Assert-GssDriveBackupTest ($compactRestoreFiles.Count -gt 0) 'Verify-only restore reconstructed every original portable path instead of compacting an over-budget destination.'
    foreach ($restoredFile in $compactRestoreFiles) {
        Assert-GssDriveBackupTest (-not [string]::IsNullOrWhiteSpace([string]$restoredFile.portable_path)) 'Compact restore receipt lost the original portable_path metadata.'
        $restoredCompactPath = Join-Path $recoveryRestore.Destination ([string]$restoredFile.restored_path).Replace('/', '\')
        Assert-GssDriveBackupTest ($restoredCompactPath.Length -lt 248 -and (Test-Path -LiteralPath $restoredCompactPath -PathType Leaf)) 'Compact verify-only restore destination is missing or exceeds the safe path budget.'
        Assert-GssDriveBackupTest ((Get-GssDriveBackupSha256 -Path $restoredCompactPath) -eq [string]$restoredFile.sha256) 'Compact verify-only restore did not preserve byte-integrity verification.'
    }

    $tamperedPurposePath = Join-Path $resolvedTestRoot 'tampered-recovery-purpose'
    Copy-Item -LiteralPath $recoveryCommitted.SnapshotPath -Destination $tamperedPurposePath -Recurse
    $tamperedPurposeReceiptPath = Join-Path $tamperedPurposePath 'commit-receipt.json'
    $tamperedPurposeReceipt = Read-GssDriveBackupJson -Path $tamperedPurposeReceiptPath
    $tamperedPurposeReceipt.snapshot_purpose = 'WorkbookTransaction'
    Write-GssDriveBackupAtomicJson -Path $tamperedPurposeReceiptPath -Value $tamperedPurposeReceipt
    $purposeMismatchRefused = $false
    try {
        [void](Test-GssCommittedBackupSnapshot -SnapshotPath $tamperedPurposePath)
    }
    catch {
        $purposeMismatchRefused = $_.Exception.Message -match 'purpose does not match'
    }
    Assert-GssDriveBackupTest $purposeMismatchRefused 'Committed RecoveryOnly evidence accepted inconsistent purpose labels.'

    $tamperedSetPath = Join-Path $resolvedTestRoot 'tampered-recovery-set'
    Copy-Item -LiteralPath $recoveryCommitted.SnapshotPath -Destination $tamperedSetPath -Recurse
    $tamperedSetManifestPath = Join-Path $tamperedSetPath 'backup-manifest.json'
    $tamperedSetReceiptPath = Join-Path $tamperedSetPath 'commit-receipt.json'
    $tamperedSetManifest = Read-GssDriveBackupJson -Path $tamperedSetManifestPath
    $originalFinalFile = @($tamperedSetManifest.files)[0]
    $extraFinalFile = [ordered]@{
        role = 'recovery_qa'
        portable_path = 'recovery/evidence/recovery-qa.json'
        snapshot_path = [string]$originalFinalFile.snapshot_path
        byte_size = [long]$originalFinalFile.byte_size
        sha256 = [string]$originalFinalFile.sha256
        classification = 'restricted_operational'
    }
    $tamperedSetManifest.files = @($tamperedSetManifest.files) + @($extraFinalFile)
    $tamperedSetManifest.file_count = @($tamperedSetManifest.files).Count
    $tamperedSetManifest.total_bytes = [long](($tamperedSetManifest.files | Measure-Object -Property byte_size -Sum).Sum)
    Write-GssDriveBackupAtomicJson -Path $tamperedSetManifestPath -Value $tamperedSetManifest
    $tamperedSetReceipt = Read-GssDriveBackupJson -Path $tamperedSetReceiptPath
    $tamperedSetReceipt.backup_manifest_sha256 = Get-GssDriveBackupSha256 -Path $tamperedSetManifestPath
    $tamperedSetReceipt.file_count = $tamperedSetManifest.file_count
    $tamperedSetReceipt.total_bytes = $tamperedSetManifest.total_bytes
    Write-GssDriveBackupAtomicJson -Path $tamperedSetReceiptPath -Value $tamperedSetReceipt
    $setMismatchRefused = $false
    try {
        [void](Test-GssCommittedBackupSnapshot -SnapshotPath $tamperedSetPath)
    }
    catch {
        $setMismatchRefused = $_.Exception.Message -match 'count does not match the prepared inventory'
    }
    Assert-GssDriveBackupTest $setMismatchRefused 'Committed RecoveryOnly evidence accepted a widened final file set.'

    $releaseProgramRoot = Join-Path $gssRoot 'GSS Survey Workbook Automation'
    $releaseManifestDirectory = Join-Path $releaseProgramRoot 'release'
    $releaseStateDirectory = Join-Path $gssRoot '_automation_runs\state\release'
    $releaseStagingRoot = Join-Path $resolvedTestRoot 'release-archive-staging'
    foreach ($folder in @(
        $releaseManifestDirectory,
        $releaseStateDirectory,
        (Join-Path $releaseStagingRoot 'release'),
        (Join-Path $gssRoot '_automation_runs\test-output\release-copy-test')
    )) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $releaseReadmeText = "# Release fixture`r`nProgram source only.`r`n"
    $releaseReadmePath = Join-Path $releaseStagingRoot 'README.md'
    Write-GssDriveBackupAtomicText -Path $releaseReadmePath -Text $releaseReadmeText
    $releaseCanonicalBytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes(
        $releaseReadmeText.Replace("`r`n", "`n").Replace("`r", "`n")
    )
    $releaseVersion = '1.1.1'
    $releaseTag = "v$releaseVersion"
    $releaseArchiveName = "gss-survey-workbook-automation-$releaseTag.zip"
    $releaseManifestPath = Join-Path $releaseManifestDirectory 'release-manifest.json'
    $releaseManifest = [ordered]@{
        schema_version = 1
        release_version = $releaseVersion
        release_tag = $releaseTag
        commit_binding = 'exact_release_tag'
        generated_at_utc = '2026-07-24T12:00:00Z'
        classification = 'PROGRAM SOURCE ONLY - NO GSS WORKBOOKS, REPORTS, OR CUSTOMER DATA'
        runtime_contract = [ordered]@{
            require_clean_tree = $true
            reject_untracked_executables = $true
            require_exact_tag_at_head = $true
            automatic_sending = 'permanently_disabled'
            live_execution = 'manual_apply_only'
            excel_validation_receipt_required = $true
            excel_validation_receipt_name = 'local-excel-validation-receipt.json'
            excel_validation_receipt_relative_path = '_automation_runs/state/release/local-excel-validation-receipt.json'
        }
        archive_name = $releaseArchiveName
        files = @(
            [ordered]@{
                path = 'README.md'
                role = 'documentation'
                hash_mode = 'utf8_lf'
                canonical_size_bytes = $releaseCanonicalBytes.Length
                sha256 = Get-GssDriveBackupByteSha256 -Bytes $releaseCanonicalBytes
            }
        )
    }
    Write-GssDriveBackupAtomicJson -Path $releaseManifestPath -Value $releaseManifest
    $manifestUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    $releaseManifestText = [System.IO.File]::ReadAllText($releaseManifestPath, $manifestUtf8)
    $releaseManifestCanonicalText = $releaseManifestText.Replace("`r`n", "`n").Replace("`r", "`n")
    [System.IO.File]::WriteAllText(
        $releaseManifestPath,
        $releaseManifestCanonicalText.Replace("`n", "`r`n"),
        $manifestUtf8
    )
    $archivedReleaseManifestPath = Join-Path $releaseStagingRoot 'release\release-manifest.json'
    [System.IO.File]::WriteAllText($archivedReleaseManifestPath, $releaseManifestCanonicalText, $manifestUtf8)
    Assert-GssDriveBackupTest `
        ((Get-GssDriveBackupSha256 -Path $releaseManifestPath) -cne (Get-GssDriveBackupSha256 -Path $archivedReleaseManifestPath)) `
        'ReleaseOnly line-ending regression fixture did not produce distinct CRLF and LF manifest bytes.'

    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    }
    catch {
        if (-not ('System.IO.Compression.ZipFile' -as [type])) { throw }
    }
    $releaseArchivePath = Join-Path $releaseStateDirectory $releaseArchiveName
    [System.IO.Compression.ZipFile]::CreateFromDirectory($releaseStagingRoot, $releaseArchivePath)

    $releaseWorkbookPath = Join-Path $gssRoot '_automation_runs\test-output\release-copy-test\GSS Score Trends - Main.xlsx'
    Write-GssDriveBackupAtomicText -Path $releaseWorkbookPath -Text 'copy-only workbook evidence'
    $releaseWorkbookHash = Get-GssDriveBackupSha256 -Path $releaseWorkbookPath
    $releaseSourceLogPath = Join-Path $gssRoot '_automation_runs\logs\release-copy-test.json'
    $releaseWorkbookRelative = '_automation_runs/test-output/release-copy-test/GSS Score Trends - Main.xlsx'
    $releaseSourceRun = [pscustomobject][ordered]@{
        Mode = 'CopyTestOnly'
        TransactionStatus = 'Prepared'
        ProgramRelease = $releaseTag
        HostName = 'RELEASE-CERTIFICATION-HOST'
        RunId = '11111111-1111-1111-1111-111111111111'
        CurrentWeekEnding = '2026-07-24'
        StartingWorkbookSha256 = '1' * 64
        CurrentSourceSha256 = '2' * 64
        PriorYearSourceSha256 = '3' * 64
        StagedWorkbookRelativePath = $releaseWorkbookRelative
        StagedWorkbookSha256 = $releaseWorkbookHash
        StagedPdfSha256 = '4' * 64
        RunFingerprint = $null
    }
    $releaseSourceRun.RunFingerprint = Get-GssDriveBackupPreparedRunFingerprint -Run $releaseSourceRun
    $releaseRunFingerprint = $releaseSourceRun.RunFingerprint
    Write-GssDriveBackupAtomicJson -Path $releaseSourceLogPath -Value $releaseSourceRun
    $releaseReceiptPath = Join-Path $releaseStateDirectory 'local-excel-validation-receipt.json'
    Write-GssDriveBackupAtomicJson -Path $releaseReceiptPath -Value ([ordered]@{
        ReceiptSchemaVersion = 1
        TimestampUtc = '2026-07-24T12:10:00Z'
        Status = 'Passed'
        Error = $null
        GitHead = 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb'
        ReleaseTag = $releaseTag
        ExcelVersion = '16.0'
        WorkbookPath = $releaseWorkbookRelative
        WorkbookSha256 = $releaseWorkbookHash
        SourceRunFingerprint = $releaseRunFingerprint
        SourceRunLogPath = '_automation_runs/logs/release-copy-test.json'
        FormulaErrors = 0
        ConstantErrors = 0
    })

    $releaseBackupInventory = @(
        [pscustomobject][ordered]@{
            SourcePath = $releaseArchivePath
            PortablePath = "release/$releaseArchiveName"
            Role = 'release_archive'
            Classification = 'restricted_operational'
        }
        [pscustomobject][ordered]@{
            SourcePath = $releaseManifestPath
            PortablePath = 'release/release-manifest.json'
            Role = 'release_manifest'
            Classification = 'restricted_operational'
        }
        [pscustomobject][ordered]@{
            SourcePath = $releaseReceiptPath
            PortablePath = 'release/local-excel-validation-receipt.json'
            Role = 'release_excel_receipt'
            Classification = 'restricted_operational'
        }
    )
    $releaseInventoryEvidence = @(Get-GssDriveBackupInventory -GssRoot $gssRoot -AdditionalItems $releaseBackupInventory -InventoryMode ReleaseOnly)
    $releaseFingerprintHash = Get-GssReleaseOnlyInventoryFingerprint -Inventory $releaseInventoryEvidence -Release $releaseTag
    $releaseSummaryPath = Join-Path $releaseStateDirectory 'drive-release-summary-v1.1.1.json'
    $releaseSummary = [ordered]@{
        schema_version = 'gss-release-drive-summary/v1'
        GeneratedAtUtc = '2026-07-24T12:00:00Z'
        RunId = "gss-release-$releaseTag-$($releaseFingerprintHash.Substring(0, 12))"
        RunFingerprint = "sha256:$releaseFingerprintHash"
        Folder = $gssRoot
        CurrentWeekEnding = '2026-07-24'
        ProgramRelease = $releaseTag
        SnapshotPurpose = 'ReleaseOnly'
        BackupInventory = $releaseBackupInventory
    }
    Write-GssDriveBackupAtomicJson -Path $releaseSummaryPath -Value $releaseSummary

    $releaseInventory = & $invokeScript -Operation Inventory -RunSummaryPath $releaseSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($releaseInventory.SnapshotPurpose -eq 'ReleaseOnly') 'ReleaseOnly purpose was not propagated.'
    Assert-GssDriveBackupTest ($releaseInventory.InventoryMode -eq 'ReleaseOnly') 'ReleaseOnly inventory mode was not propagated.'
    Assert-GssDriveBackupTest ($releaseInventory.FileCount -eq 3 -and -not $releaseInventory.ContainsPersonalData) 'ReleaseOnly inventory was not the exact non-personal-data triad.'

    $releaseSourceRun = Read-GssDriveBackupJson -Path $releaseSourceLogPath
    $releaseSourceRun.HostName = ''
    Write-GssDriveBackupAtomicJson -Path $releaseSourceLogPath -Value $releaseSourceRun
    $releaseMissingAuditHostRefused = $false
    try {
        [void](& $invokeScript -Operation Inventory -RunSummaryPath $releaseSummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $releaseMissingAuditHostRefused = $_.Exception.Message -match 'does not match its Prepared copy-only source run'
    }
    Assert-GssDriveBackupTest $releaseMissingAuditHostRefused 'ReleaseOnly accepted a source run without an audit workstation.'
    $releaseSourceRun.HostName = 'RELEASE-CERTIFICATION-HOST'
    Write-GssDriveBackupAtomicJson -Path $releaseSourceLogPath -Value $releaseSourceRun

    $releaseSourceRun.HostName = 'TAMPERED-NONBLANK-HOST'
    Write-GssDriveBackupAtomicJson -Path $releaseSourceLogPath -Value $releaseSourceRun
    $releaseTamperedAuditHostRefused = $false
    try {
        [void](& $invokeScript -Operation Inventory -RunSummaryPath $releaseSummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $releaseTamperedAuditHostRefused = $_.Exception.Message -match 'does not match its Prepared copy-only source run'
    }
    Assert-GssDriveBackupTest $releaseTamperedAuditHostRefused 'ReleaseOnly accepted a nonblank audit workstation that did not match the run fingerprint.'
    $releaseSourceRun.HostName = 'RELEASE-CERTIFICATION-HOST'
    Write-GssDriveBackupAtomicJson -Path $releaseSourceLogPath -Value $releaseSourceRun

    $releaseExtraPath = Join-Path $releaseStateDirectory 'not-approved.txt'
    Write-GssDriveBackupAtomicText -Path $releaseExtraPath -Text 'must not enter ReleaseOnly'
    $releaseWidenedSummaryPath = Join-Path $resolvedTestRoot 'release-widened-summary.json'
    $releaseWidenedSummary = $releaseSummary | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    $releaseWidenedSummary.BackupInventory = @($releaseWidenedSummary.BackupInventory) + @(
        [pscustomobject]@{
            SourcePath = $releaseExtraPath
            PortablePath = 'release/not-approved.txt'
            Role = 'release_manifest'
            Classification = 'restricted_operational'
        }
    )
    Write-GssDriveBackupAtomicJson -Path $releaseWidenedSummaryPath -Value $releaseWidenedSummary
    $releaseWidenedRefused = $false
    try {
        [void](& $invokeScript -Operation Inventory -RunSummaryPath $releaseWidenedSummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $releaseWidenedRefused = $_.Exception.Message -match 'exactly three'
    }
    Assert-GssDriveBackupTest $releaseWidenedRefused 'ReleaseOnly accepted a fourth artifact.'

    $releaseJunctionTarget = Join-Path $resolvedTestRoot 'release-junction-target'
    New-Item -ItemType Directory -Path $releaseJunctionTarget -Force | Out-Null
    Copy-Item -LiteralPath $releaseArchivePath -Destination (Join-Path $releaseJunctionTarget $releaseArchiveName)
    $releaseJunctionPath = Join-Path $releaseStateDirectory 'linked'
    New-Item -ItemType Junction -Path $releaseJunctionPath -Target $releaseJunctionTarget | Out-Null
    $releaseJunctionSummaryPath = Join-Path $resolvedTestRoot 'release-junction-summary.json'
    $releaseJunctionSummary = $releaseSummary | ConvertTo-Json -Depth 12 | ConvertFrom-Json
    @($releaseJunctionSummary.BackupInventory | Where-Object Role -eq 'release_archive')[0].SourcePath =
        Join-Path $releaseJunctionPath $releaseArchiveName
    Write-GssDriveBackupAtomicJson -Path $releaseJunctionSummaryPath -Value $releaseJunctionSummary
    $releaseJunctionRefused = $false
    try {
        [void](& $invokeScript -Operation Inventory -RunSummaryPath $releaseJunctionSummaryPath -SettingsPath $settingsPath -OutputObject)
    }
    catch {
        $releaseJunctionRefused = $_.Exception.Message -match 'symbolic link or junction'
    }
    Assert-GssDriveBackupTest $releaseJunctionRefused 'ReleaseOnly traversed a junction to release bytes outside the GSS root.'

    $releasePrepared = & $invokeScript -Operation Prepare -RunSummaryPath $releaseSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($releasePrepared.Status -eq 'Prepared') 'ReleaseOnly preparation failed.'
    $releasePreparedManifest = Read-GssDriveBackupJson -Path $releasePrepared.PreparedManifestPath
    Assert-GssDriveBackupTest ($releasePreparedManifest.snapshot_purpose -eq 'ReleaseOnly') 'ReleaseOnly prepared manifest omitted its purpose.'
    Assert-GssDriveBackupTest ($releasePreparedManifest.scope.inventory_mode -eq 'ReleaseOnly') 'ReleaseOnly prepared manifest omitted its inventory mode.'
    Assert-GssDriveBackupTest (-not [bool]$releasePreparedManifest.data_classification.contains_personal_data) 'ReleaseOnly prepared manifest incorrectly claimed personal data.'
    $releasePreparedRetry = & $invokeScript -Operation Prepare -RunSummaryPath $releaseSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($releasePreparedRetry.Status -eq 'Prepared' -and $releasePreparedRetry.Idempotent) 'ReleaseOnly preparation was not idempotent.'

    $releaseArchiveOriginal = Join-Path $resolvedTestRoot 'release-archive-original.zip'
    Copy-Item -LiteralPath $releaseArchivePath -Destination $releaseArchiveOriginal
    Write-GssDriveBackupAtomicText -Path $releaseArchivePath -Text 'changed after preparation'
    $changedReleaseFinalize = & $invokeScript -Operation Finalize -RunSummaryPath $releaseSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($changedReleaseFinalize.Status -eq 'PendingFinalize') 'ReleaseOnly finalization accepted changed archive bytes.'
    Copy-Item -LiteralPath $releaseArchiveOriginal -Destination $releaseArchivePath -Force

    $releaseCommitted = & $invokeScript -Operation Finalize -RunSummaryPath $releaseSummaryPath -SettingsPath $settingsPath -OutputObject
    Assert-GssDriveBackupTest ($releaseCommitted.Status -eq 'Committed' -and $releaseCommitted.FileCount -eq 3) "ReleaseOnly finalize did not commit the exact triad: $($releaseCommitted.Error)"
    $releaseValidation = Test-GssCommittedBackupSnapshot -SnapshotPath $releaseCommitted.SnapshotPath
    Assert-GssDriveBackupTest ($releaseValidation.Manifest.snapshot_purpose -eq 'ReleaseOnly') 'Committed ReleaseOnly snapshot omitted its purpose.'
    Assert-GssDriveBackupTest ($releaseValidation.Manifest.scope.inventory_mode -eq 'ReleaseOnly') 'Committed ReleaseOnly snapshot omitted its narrow inventory mode.'
    Assert-GssDriveBackupTest (@($releaseValidation.Manifest.files).Count -eq 3) 'Committed ReleaseOnly snapshot was widened.'
    $releaseRestore = Restore-GssDriveBackupForVerification -RunId $releaseSummary.RunId -SettingsPath $settingsPath -LocalAppDataPath $localAppData -Phase Final
    Assert-GssDriveBackupTest ($releaseRestore.Status -eq 'Verified' -and $releaseRestore.FileCount -eq 3 -and -not $releaseRestore.LiveWorkbookOverwritten) 'ReleaseOnly verify-only restore failed.'

    Move-Item -LiteralPath $recoveredPathOne -Destination $recoveredHoldingRoot
    Move-Item -LiteralPath $recoveredPathTwo -Destination $recoveredHoldingRoot

    # A partial snapshot created by an older release can fit while staged but
    # exceed the budget after promotion. Finalize and RetryFinalize must both
    # reject that prepared-manifest path before moving the directory.
    $legacyBudgetRunId = 'gss-legacy-budget-20260719-abcdef'
    $legacyBudgetFingerprint = 'sha256:legacy-budget-0123456789abcdef0123456789abcdef0123456789abcdef'
    $legacyBudgetSource = Join-Path $recoveredHoldingRoot ([System.IO.Path]::GetFileName($recoveredPathOne))
    $legacyBudgetInventory = @(
        [pscustomobject]@{
            SourcePath = $legacyBudgetSource
            PortablePath = 'legacy/item.txt'
            Role = 'recovered_historical_detail'
            Classification = 'restricted_personal_data'
        }
    )
    $legacyBudgetPrepared = New-GssDriveBackupPreparedSnapshot -RunId $legacyBudgetRunId -Fingerprint $legacyBudgetFingerprint -ReportWeek ([datetime]'2026-07-19') -Inventory $legacyBudgetInventory -SnapshotPurpose WorkbookTransaction -SettingsPath $settingsPath
    $legacyPreparedManifest = Read-GssDriveBackupJson -Path $legacyBudgetPrepared.PreparedManifestPath
    $legacyPreparedEntry = @($legacyPreparedManifest.files)[0]
    $legacyOriginalPath = Join-Path $legacyBudgetPrepared.PreparedPath ([string]$legacyPreparedEntry.snapshot_path).Replace('/', '\')
    $legacyRelativePrefix = 'prepared-payload/'
    $legacyLeafLength = 246 - (Join-Path $legacyBudgetPrepared.PreparedPath $legacyRelativePrefix.Replace('/', '\')).Length
    Assert-GssDriveBackupTest ($legacyLeafLength -gt 4) 'Legacy promotion-budget fixture could not calibrate its staged leaf.'
    $legacySnapshotRelative = $legacyRelativePrefix + ('l' * ($legacyLeafLength - 4)) + '.txt'
    $legacyStagedPath = Join-Path $legacyBudgetPrepared.PreparedPath $legacySnapshotRelative.Replace('/', '\')
    $legacyFinalPath = Join-Path $driveRoot "snapshots\2026\2026-07\$legacyBudgetRunId"
    $legacyPromotedPath = Join-Path $legacyFinalPath $legacySnapshotRelative.Replace('/', '\')
    Assert-GssDriveBackupTest ($legacyStagedPath.Length -eq 246 -and $legacyPromotedPath.Length -ge 248) 'Legacy promotion-budget fixture did not fit staging while exceeding the promoted path budget.'
    $legacyStagedParent = Split-Path -Parent $legacyStagedPath
    New-Item -ItemType Directory -Path $legacyStagedParent -Force | Out-Null
    Move-Item -LiteralPath $legacyOriginalPath -Destination $legacyStagedPath
    $legacyPreparedEntry.snapshot_path = $legacySnapshotRelative
    Write-GssDriveBackupAtomicJson -Path $legacyBudgetPrepared.PreparedManifestPath -Value $legacyPreparedManifest

    foreach ($attempt in @('Finalize', 'RetryFinalize')) {
        $legacyBudgetRefused = $false
        try {
            [void](Complete-GssDriveBackupSnapshot -RunId $legacyBudgetRunId -Fingerprint $legacyBudgetFingerprint -FinalInventory $legacyBudgetInventory -SnapshotPurpose WorkbookTransaction -SettingsPath $settingsPath)
        }
        catch {
            $legacyBudgetRefused = $_.Exception.Message -match 'Prepared manifest snapshot path exceeds the safe Windows path budget at the promoted destination'
        }
        Assert-GssDriveBackupTest $legacyBudgetRefused "$attempt moved or accepted a legacy prepared payload that exceeds the promoted path budget."
        Assert-GssDriveBackupTest ((Test-Path -LiteralPath $legacyBudgetPrepared.PreparedPath -PathType Container) -and -not (Test-Path -LiteralPath $legacyFinalPath)) "$attempt moved the legacy partial snapshot before enforcing the path budget."
        Assert-GssDriveBackupTest (Test-Path -LiteralPath (Join-Path $legacyBudgetPrepared.PreparedPath 'backup-manifest.json') -PathType Leaf) "$attempt did not exercise the existing final-manifest retry path."
    }
    $legacyBudgetAborted = Stop-GssDriveBackupPreparedSnapshot -RunId $legacyBudgetRunId -Fingerprint $legacyBudgetFingerprint -SettingsPath $settingsPath -Confirm:$false
    Assert-GssDriveBackupTest ($legacyBudgetAborted.Status -eq 'Aborted') 'Legacy promotion-budget fixture did not retain safely aborted evidence.'

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
    Assert-GssDriveBackupTest ($hashOnlyDrillStatus.HashOnlyVerificationCount -eq 3 -and -not $hashOnlyDrillStatus.HashOnlyVerificationSatisfiesQuarterlyDrill) 'Hash-only restore evidence was not exposed separately.'

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

    # Exercise the long portable path through the complete Full lifecycle, not
    # only through the copy helper or the narrower RecoveryOnly workflow.
    $longLifecycleRunId = 'gss-long-full-20260719-abcdef'
    $longLifecycleFingerprint = 'sha256:long-full-0123456789abcdef0123456789abcdef0123456789abcdef'
    $longLifecycleSource = Join-Path $recoveredHoldingRoot ([System.IO.Path]::GetFileName($recoveredPathOne))
    $reservedRestorePortable = Get-GssDriveBackupCompactRelativePath -PortablePath $longPathPortable -Prefix 'r'
    $restoreCollisionFixtures = @(
        [pscustomobject]@{
            PortablePath = $reservedRestorePortable
            SourcePath = Join-Path $recoveredHoldingRoot 'reserved-restore-namespace.xlsx'
        },
        [pscustomobject]@{
            PortablePath = 'r'
            SourcePath = Join-Path $recoveredHoldingRoot 'reserved-restore-root.bin'
        },
        [pscustomobject]@{
            PortablePath = 'restore-verification.json'
            SourcePath = Join-Path $recoveredHoldingRoot 'reserved-restore-receipt.bin'
        },
        [pscustomobject]@{
            PortablePath = 'local-excel-validation-receipt.json'
            SourcePath = Join-Path $recoveredHoldingRoot 'reserved-excel-receipt.bin'
        },
        [pscustomobject]@{
            PortablePath = 'quarterly-restore-drill.json'
            SourcePath = Join-Path $recoveredHoldingRoot 'reserved-quarterly-receipt.bin'
        }
    )
    foreach ($fixture in $restoreCollisionFixtures) {
        Set-Content -LiteralPath $fixture.SourcePath -Value "restore collision fixture: $($fixture.PortablePath)" -Encoding UTF8
    }
    $longLifecycleInventory = @(
        [pscustomobject]@{
            SourcePath = $longLifecycleSource
            PortablePath = $longPathPortable
            Role = 'recovered_historical_detail'
            Classification = 'restricted_personal_data'
        }
    ) + @($restoreCollisionFixtures | ForEach-Object {
        [pscustomobject]@{
            SourcePath = $_.SourcePath
            PortablePath = $_.PortablePath
            Role = 'recovered_historical_detail'
            Classification = 'restricted_personal_data'
        }
    })
    $longLifecyclePrepared = New-GssDriveBackupPreparedSnapshot -RunId $longLifecycleRunId -Fingerprint $longLifecycleFingerprint -ReportWeek ([datetime]'2026-07-19') -Inventory $longLifecycleInventory -SnapshotPurpose WorkbookTransaction -SettingsPath $settingsPath
    $longPreparedManifest = Read-GssDriveBackupJson -Path $longLifecyclePrepared.PreparedManifestPath
    $longPreparedEntry = @($longPreparedManifest.files | Where-Object portable_path -eq $longPathPortable)[0]
    Assert-GssDriveBackupTest ($longPreparedEntry.portable_path -eq $longPathPortable) 'Full Prepare did not retain the original long portable_path metadata.'
    Assert-GssDriveBackupTest ($longPreparedEntry.snapshot_path -match '^prepared-payload/long-path/[0-9a-f]{64}\.xlsx$') 'Full Prepare did not persist the compact prepared snapshot_path.'
    $longLifecycleCommitted = Complete-GssDriveBackupSnapshot -RunId $longLifecycleRunId -Fingerprint $longLifecycleFingerprint -FinalInventory $longLifecycleInventory -SnapshotPurpose WorkbookTransaction -SettingsPath $settingsPath
    Assert-GssDriveBackupTest ($longLifecycleCommitted.Status -eq 'Committed') 'Full long-path lifecycle did not promote and commit.'
    $longCommittedValidation = Test-GssCommittedBackupSnapshot -SnapshotPath $longLifecycleCommitted.SnapshotPath
    $longFinalEntry = @($longCommittedValidation.Manifest.files | Where-Object portable_path -eq $longPathPortable)[0]
    $longPersistedPreparedEntry = @($longCommittedValidation.Manifest.prepared_files | Where-Object portable_path -eq $longPathPortable)[0]
    Assert-GssDriveBackupTest ($longFinalEntry.portable_path -eq $longPathPortable -and $longPersistedPreparedEntry.portable_path -eq $longPathPortable) 'Full Finalize did not retain original portable_path metadata in both inventories.'
    Assert-GssDriveBackupTest ($longFinalEntry.snapshot_path -match '^payload/long-path/[0-9a-f]{64}\.xlsx$' -and $longPersistedPreparedEntry.snapshot_path -eq $longPreparedEntry.snapshot_path) 'Full Finalize did not persist compact final and prepared snapshot paths.'
    $longLifecycleRestoreRoot = Join-Path $resolvedTestRoot ('full-restore-' + ('z' * 30))
    $longLifecycleRestore = Restore-GssDriveBackupForVerification -RunId $longLifecycleRunId -SettingsPath $settingsPath -LocalAppDataPath $longLifecycleRestoreRoot -Phase Final
    $longLifecycleRestoreReceipt = Read-GssDriveBackupJson -Path $longLifecycleRestore.ReceiptPath
    $longRestoredEntry = @($longLifecycleRestoreReceipt.files | Where-Object portable_path -eq $longPathPortable)[0]
    $longRestoredPath = Join-Path $longLifecycleRestore.Destination ([string]$longRestoredEntry.restored_path).Replace('/', '\')
    Assert-GssDriveBackupTest ($longRestoredEntry.portable_path -eq $longPathPortable -and $longRestoredEntry.restored_path -match '^r/long-path/[0-9a-f]{64}(?:\.[^/]+)?$') 'Full VerifyRestore did not retain portable metadata and map to a compact destination.'
    Assert-GssDriveBackupTest ($longRestoredPath.Length -lt 248 -and (Test-Path -LiteralPath $longRestoredPath -PathType Leaf)) 'Full VerifyRestore compact destination is absent or over budget.'
    Assert-GssDriveBackupTest ((Get-GssDriveBackupSha256 -Path $longRestoredPath) -eq (Get-GssDriveBackupSha256 -Path $longLifecycleSource) -and [string]$longRestoredEntry.sha256 -eq (Get-GssDriveBackupSha256 -Path $longLifecycleSource)) 'Full VerifyRestore did not preserve matching file bytes and hash evidence.'
    $resolvedLongMapping = Resolve-GssDriveBackupRestoredFile -ReceiptPath $longLifecycleRestore.ReceiptPath -PortablePath $longPathPortable -ExpectedDestination $longLifecycleRestore.Destination
    Assert-GssDriveBackupTest ($resolvedLongMapping.Path -eq $longRestoredPath) 'Restore receipt resolver did not return the compact long-path destination.'
    foreach ($fixture in $restoreCollisionFixtures) {
        $reservedRestoredEntry = @($longLifecycleRestoreReceipt.files | Where-Object portable_path -eq $fixture.PortablePath)[0]
        $reservedRestoredPath = Join-Path $longLifecycleRestore.Destination ([string]$reservedRestoredEntry.restored_path).Replace('/', '\')
        Assert-GssDriveBackupTest ($reservedRestoredEntry.restored_path -match '^r/long-path/[0-9a-f]{64}(?:\.[^/]+)?$' -and $reservedRestoredEntry.restored_path -ne $fixture.PortablePath) "VerifyRestore did not remap reserved portable path '$($fixture.PortablePath)'."
        Assert-GssDriveBackupTest ($reservedRestoredPath.Length -lt 248 -and (Test-Path -LiteralPath $reservedRestoredPath -PathType Leaf)) "Reserved restore destination for '$($fixture.PortablePath)' is absent or over budget."
        Assert-GssDriveBackupTest ((Get-GssDriveBackupSha256 -Path $reservedRestoredPath) -eq (Get-GssDriveBackupSha256 -Path $fixture.SourcePath) -and [string]$reservedRestoredEntry.sha256 -eq (Get-GssDriveBackupSha256 -Path $fixture.SourcePath)) "Reserved restore mapping for '$($fixture.PortablePath)' did not preserve matching file bytes and hash evidence."
        $resolvedReservedMapping = Resolve-GssDriveBackupRestoredFile -ReceiptPath $longLifecycleRestore.ReceiptPath -PortablePath $fixture.PortablePath -ExpectedDestination $longLifecycleRestore.Destination
        Assert-GssDriveBackupTest ($resolvedReservedMapping.Path -eq $reservedRestoredPath) "Restore receipt resolver did not return the remapped destination for '$($fixture.PortablePath)'."
    }

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
            -not ([System.IO.Path]::GetFileName($verifiedCleanupPath)).StartsWith('gdbt-', [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing unsafe recursive test cleanup: $verifiedCleanupPath"
        }
        Remove-Item -LiteralPath $verifiedCleanupPath -Recurse -Force
    }
}
