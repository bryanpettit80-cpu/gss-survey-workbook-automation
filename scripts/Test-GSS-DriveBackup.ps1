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

    Move-Item -LiteralPath $recoveredPathOne -Destination $recoveredHoldingRoot
    Move-Item -LiteralPath $recoveredPathTwo -Destination $recoveredHoldingRoot
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
            -not ([System.IO.Path]::GetFileName($verifiedCleanupPath)).StartsWith('gdbt-', [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing unsafe recursive test cleanup: $verifiedCleanupPath"
        }
        Remove-Item -LiteralPath $verifiedCleanupPath -Recurse -Force
    }
}
