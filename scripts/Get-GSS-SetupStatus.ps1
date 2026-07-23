[CmdletBinding()]
param(
    [string]$Folder
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$projectRoot = Split-Path -Parent $scriptRoot
. (Join-Path $scriptRoot 'Gss-DriveBackup.ps1')

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $Folder = Split-Path -Parent $projectRoot
}

$Folder = (Resolve-Path -LiteralPath $Folder).Path
$template = Join-Path $projectRoot 'templates\Run GSS Update After Upload.cmd'
$operatorLauncher = Join-Path $Folder 'Run GSS Update After Upload.cmd'
$guideTemplate = Join-Path $projectRoot 'templates\00 START HERE - GSS Survey Updates.txt'
$operatorGuide = Join-Path $Folder '00 START HERE - GSS Survey Updates.txt'
$dropGuideTemplate = Join-Path $projectRoot 'templates\00 DROP NEW WEEKLY WORKBOOK HERE.txt'
$dropGuide = Join-Path $Folder '02 Weekly Rolling Source Workbooks\00 DROP NEW WEEKLY WORKBOOK HERE.txt'
$uploadGuideTemplate = Join-Path $projectRoot 'templates\00 ABOUT THIS FOLDER.txt'
$uploadGuide = Join-Path $Folder '03 Uploaded Survey Workbooks\00 ABOUT THIS FOLDER.txt'
$task = Get-ScheduledTask -TaskName 'GSS Survey Main Workbook Weekly Update' -ErrorAction SilentlyContinue
$templateHash = if (Test-Path -LiteralPath $template) { (Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash } else { $null }
$operatorHash = if (Test-Path -LiteralPath $operatorLauncher) { (Get-FileHash -LiteralPath $operatorLauncher -Algorithm SHA256).Hash } else { $null }
$guideTemplateHash = if (Test-Path -LiteralPath $guideTemplate) { (Get-FileHash -LiteralPath $guideTemplate -Algorithm SHA256).Hash } else { $null }
$operatorGuideHash = if (Test-Path -LiteralPath $operatorGuide) { (Get-FileHash -LiteralPath $operatorGuide -Algorithm SHA256).Hash } else { $null }
$dropGuideTemplateHash = if (Test-Path -LiteralPath $dropGuideTemplate) { (Get-FileHash -LiteralPath $dropGuideTemplate -Algorithm SHA256).Hash } else { $null }
$dropGuideHash = if (Test-Path -LiteralPath $dropGuide) { (Get-FileHash -LiteralPath $dropGuide -Algorithm SHA256).Hash } else { $null }
$uploadGuideTemplateHash = if (Test-Path -LiteralPath $uploadGuideTemplate) { (Get-FileHash -LiteralPath $uploadGuideTemplate -Algorithm SHA256).Hash } else { $null }
$uploadGuideHash = if (Test-Path -LiteralPath $uploadGuide) { (Get-FileHash -LiteralPath $uploadGuide -Algorithm SHA256).Hash } else { $null }

$driveSettingsPath = Get-GssDriveBackupDefaultSettingsPath
$driveSettingsExists = Test-Path -LiteralPath $driveSettingsPath -PathType Leaf
$driveBackupConfigured = $false
$driveBackupAvailable = $false
$driveBackupRoot = $null
$driveBackupFolderId = $null
$driveBackupExpectedOwner = $null
$driveBackupMarkerId = $null
$driveBackupMarkerValid = $false
$driveBackupVerificationLevel = $null
$driveBackupRetention = $null
$driveBackupRequiredBeforeApply = $false
$driveBackupLatestCommittedRunId = $null
$driveBackupError = $null
$driveBackupCloudReadbackPath = $null
$driveBackupCloudReadbackValid = $false
$driveBackupCloudReadbackVerifiedAtUtc = $null
$restoreDrillStatus = Get-GssDriveBackupRestoreDrillStatus
if ($driveSettingsExists) {
    try {
        $driveContext = Get-GssDriveBackupRootContext -SettingsPath $driveSettingsPath
        $driveBackupConfigured = $true
        $driveBackupAvailable = $true
        $driveBackupRoot = $driveContext.RootPath
        $driveBackupFolderId = $driveContext.Settings.DriveFolderId
        $driveBackupExpectedOwner = $driveContext.Settings.ExpectedOwner
        $driveBackupMarkerId = $driveContext.Settings.MarkerId
        $driveBackupMarkerValid = $true
        $driveBackupVerificationLevel = $driveContext.Settings.VerificationLevel
        $driveBackupRetention = "$($driveContext.Settings.RetentionWeekly) weekly + $($driveContext.Settings.RetentionMonthly) monthly; report-only deletion candidates"
        $driveBackupRequiredBeforeApply = $driveContext.Settings.RequireBeforeApply
        $driveBackupCloudReadbackPath = Join-Path $driveContext.RootPath 'commissioning-readback.json'
        if (Test-Path -LiteralPath $driveBackupCloudReadbackPath -PathType Leaf) {
            $cloudReadback = Read-GssDriveBackupJson -Path $driveBackupCloudReadbackPath
            $driveBackupCloudReadbackValid = (
                [string](Get-GssDriveBackupProperty $cloudReadback @('drive_folder_id')) -eq $driveContext.Settings.DriveFolderId -and
                [string](Get-GssDriveBackupProperty $cloudReadback @('owner')) -eq $driveContext.Settings.ExpectedOwner -and
                -not [bool](Get-GssDriveBackupProperty $cloudReadback @('shared')) -and
                [int](Get-GssDriveBackupProperty $cloudReadback @('permission_count')) -eq 1 -and
                [string](Get-GssDriveBackupProperty $cloudReadback @('verification')) -eq 'cloud_metadata_readback_verified'
            )
            $driveBackupCloudReadbackVerifiedAtUtc = [string](Get-GssDriveBackupProperty $cloudReadback @('verified_at_utc'))
        }
        $chainHead = Get-GssDriveBackupChainHead -RootPath $driveContext.RootPath
        if ($chainHead) {
            $driveBackupLatestCommittedRunId = $chainHead.RunId
        }
    }
    catch {
        $driveBackupError = $_.Exception.Message
        try {
            $settingsOnly = Get-GssDriveBackupSetting -SettingsPath $driveSettingsPath
            $driveBackupConfigured = $true
            $driveBackupRoot = $settingsOnly.DriveRootPath
            $driveBackupFolderId = $settingsOnly.DriveFolderId
            $driveBackupExpectedOwner = $settingsOnly.ExpectedOwner
            $driveBackupMarkerId = $settingsOnly.MarkerId
            $driveBackupVerificationLevel = $settingsOnly.VerificationLevel
            $driveBackupRetention = "$($settingsOnly.RetentionWeekly) weekly + $($settingsOnly.RetentionMonthly) monthly; report-only deletion candidates"
            $driveBackupRequiredBeforeApply = $settingsOnly.RequireBeforeApply
        }
        catch {
            $driveBackupError = $_.Exception.Message
        }
    }
}

$status = [ordered]@{
    RepoRoot = $projectRoot
    DropboxFolder = $Folder
    ManualWorkflow = 'Enabled; use Run GSS Update After Upload.cmd after uploading the newest source workbook.'
    OperatorLauncher = $operatorLauncher
    OperatorLauncherExists = [bool](Test-Path -LiteralPath $operatorLauncher)
    OperatorLauncherMatchesTemplate = ($templateHash -and $operatorHash -and $templateHash -eq $operatorHash)
    OperatorGuide = $operatorGuide
    OperatorGuideExists = [bool](Test-Path -LiteralPath $operatorGuide)
    OperatorGuideMatchesTemplate = ($guideTemplateHash -and $operatorGuideHash -and $guideTemplateHash -eq $operatorGuideHash)
    WeeklyDropGuideExists = [bool](Test-Path -LiteralPath $dropGuide)
    WeeklyDropGuideMatchesTemplate = ($dropGuideTemplateHash -and $dropGuideHash -and $dropGuideTemplateHash -eq $dropGuideHash)
    UploadedSurveyGuideExists = [bool](Test-Path -LiteralPath $uploadGuide)
    UploadedSurveyGuideMatchesTemplate = ($uploadGuideTemplateHash -and $uploadGuideHash -and $uploadGuideTemplateHash -eq $uploadGuideHash)
    ScheduledTaskName = 'GSS Survey Main Workbook Weekly Update'
    ScheduledTaskExists = [bool]$task
    ScheduledTaskState = if ($task) { [string]$task.State } else { $null }
    ScheduledTaskEnabled = if ($task) { [bool]($task.Settings.Enabled) } else { $false }
    ScheduledTaskAction = if ($task) { ($task.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join '; ' } else { $null }
    DriveBackupSettingsPath = $driveSettingsPath
    DriveBackupSettingsExists = $driveSettingsExists
    DriveBackupConfigured = $driveBackupConfigured
    DriveBackupAvailable = $driveBackupAvailable
    DriveBackupRoot = $driveBackupRoot
    DriveBackupFolderId = $driveBackupFolderId
    DriveBackupExpectedOwner = $driveBackupExpectedOwner
    DriveBackupMarkerId = $driveBackupMarkerId
    DriveBackupMarkerValid = $driveBackupMarkerValid
    DriveBackupVerificationLevel = $driveBackupVerificationLevel
    DriveBackupRetention = $driveBackupRetention
    DriveBackupRequiredBeforeApply = $driveBackupRequiredBeforeApply
    DriveBackupLatestCommittedRunId = $driveBackupLatestCommittedRunId
    DriveBackupError = $driveBackupError
    DriveBackupCloudMetadataReadbackPath = $driveBackupCloudReadbackPath
    DriveBackupCloudMetadataReadbackValid = $driveBackupCloudReadbackValid
    DriveBackupCloudMetadataReadbackVerifiedAtUtc = $driveBackupCloudReadbackVerifiedAtUtc
    DriveBackupRestoreDrillFrequency = $restoreDrillStatus.Frequency
    DriveBackupRestoreDrillStatus = $restoreDrillStatus.Status
    DriveBackupRestoreDrillLastVerifiedAtUtc = $restoreDrillStatus.LastVerifiedAtUtc
    DriveBackupRestoreDrillLastReceiptPath = $restoreDrillStatus.LastReceiptPath
    DriveBackupRestoreDrillLastExcelValidationReceiptPath = $restoreDrillStatus.LastExcelValidationReceiptPath
    DriveBackupRestoreDrillNextDueDate = $restoreDrillStatus.NextDueDate
    DriveBackupRestoreDrillLiveOverwriteAllowed = $restoreDrillStatus.LiveWorkbookOverwriteAllowed
    DriveBackupHashOnlyVerificationCount = $restoreDrillStatus.HashOnlyVerificationCount
    DriveBackupLatestHashOnlyVerifiedAtUtc = $restoreDrillStatus.LatestHashOnlyVerifiedAtUtc
    DriveBackupHashOnlyVerificationSatisfiesQuarterlyDrill = $restoreDrillStatus.HashOnlyVerificationSatisfiesQuarterlyDrill
}

[pscustomobject]$status | Format-List
