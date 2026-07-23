$script:GssDriveBackupClassificationLabel = 'CONTAINS PERSONAL DATA ' + [char]0x2014 + ' RESTRICTED'

function Get-GssDriveBackupDefaultSettingsPath {
    [CmdletBinding()]
    param(
        [string]$LocalAppDataPath = $env:LOCALAPPDATA
    )

    if ([string]::IsNullOrWhiteSpace($LocalAppDataPath)) {
        throw 'LOCALAPPDATA is unavailable; the GSS Drive backup settings path cannot be resolved.'
    }

    return (Join-Path $LocalAppDataPath 'GSSSurveyWorkbookAutomation\settings.json')
}

function Write-GssDriveBackupAtomicJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [object]$Value,
        [int]$Depth = 12
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $parent = Split-Path -Parent $fullPath
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    # Keep atomic helper names short. Verify-only restore paths can approach the
    # legacy Windows MAX_PATH limit, and appending the full destination filename
    # plus a GUID makes an otherwise valid restore intermittently fail.
    $temporaryPath = Join-Path $parent ('.t-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $json = $Value | ConvertTo-Json -Depth $Depth
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($temporaryPath, ($json + [Environment]::NewLine), $encoding)

    $replacementBackupPath = Join-Path $parent ('.b-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    try {
        if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
            [System.IO.File]::Replace($temporaryPath, $fullPath, $replacementBackupPath, $true)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $fullPath)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
        if (Test-Path -LiteralPath $replacementBackupPath -PathType Leaf) {
            Remove-Item -LiteralPath $replacementBackupPath -Force
        }
    }
}

function Write-GssDriveBackupAtomicText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$Text
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $parent = Split-Path -Parent $fullPath
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $temporaryPath = Join-Path $parent ('.t-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($temporaryPath, $Text, $encoding)
    $replacementBackupPath = Join-Path $parent ('.b-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    try {
        if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
            [System.IO.File]::Replace($temporaryPath, $fullPath, $replacementBackupPath, $true)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $fullPath)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
        if (Test-Path -LiteralPath $replacementBackupPath -PathType Leaf) {
            Remove-Item -LiteralPath $replacementBackupPath -Force
        }
    }
}

function Read-GssDriveBackupJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required JSON file is missing: $Path"
    }

    try {
        return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    catch {
        throw "Required JSON file is not valid JSON: $Path. $($_.Exception.Message)"
    }
}

function Get-GssDriveBackupProperty {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$InputObject,
        [Parameter(Mandatory)]
        [string[]]$Names,
        [AllowNull()]
        [object]$Default = $null
    )

    if ($null -eq $InputObject) {
        return $Default
    }

    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties |
            Where-Object { $_.Name.Equals($name, [System.StringComparison]::OrdinalIgnoreCase) } |
            Select-Object -First 1
        if ($property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $Default
}

function Assert-GssDriveBackupSafeRelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $candidate = $Path.Replace('\', '/').Trim('/')
    if ([string]::IsNullOrWhiteSpace($candidate) -or
        [System.IO.Path]::IsPathRooted($candidate) -or
        $candidate -match '(^|/)\.\.?(/|$)' -or
        $candidate.Contains(':')) {
        throw "Unsafe portable backup path: $Path"
    }

    return $candidate
}

function Get-GssDriveBackupRelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$Root
    )

    $fullRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if (-not $fullPath.StartsWith("$fullRoot\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "File is outside the expected root '$fullRoot': $fullPath"
    }

    return (Assert-GssDriveBackupSafeRelativePath -Path $fullPath.Substring($fullRoot.Length + 1))
}

function Get-GssDriveBackupSha256 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Cannot hash missing backup file: $Path"
    }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Initialize-GssDriveBackupConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DriveRootPath,
        [Parameter(Mandatory)]
        [string]$DriveFolderId,
        [Parameter(Mandatory)]
        [string]$ExpectedOwner,
        [string]$MarkerId,
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath)
    )

    $root = [System.IO.Path]::GetFullPath($DriveRootPath).TrimEnd('\', '/')
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        throw "The dedicated Google Drive backup folder is unavailable: $root. No fallback destination is allowed."
    }
    if ([string]::IsNullOrWhiteSpace($DriveFolderId) -or
        [string]::IsNullOrWhiteSpace($ExpectedOwner)) {
        throw 'Drive folder ID and expected owner are required.'
    }

    $markerPath = Join-Path $root 'backup-root.json'
    if ([string]::IsNullOrWhiteSpace($MarkerId) -and (Test-Path -LiteralPath $markerPath -PathType Leaf)) {
        $existingMarkerForId = Read-GssDriveBackupJson -Path $markerPath
        if ([string](Get-GssDriveBackupProperty $existingMarkerForId @('drive_folder_id')) -ne $DriveFolderId -or
            [string](Get-GssDriveBackupProperty $existingMarkerForId @('expected_owner')) -ne $ExpectedOwner) {
            throw "Existing backup root marker belongs to a different destination identity: $markerPath"
        }
        $MarkerId = [string](Get-GssDriveBackupProperty $existingMarkerForId @('marker_id'))
    }
    if ([string]::IsNullOrWhiteSpace($MarkerId)) {
        $MarkerId = [guid]::NewGuid().ToString()
    }

    $marker = [ordered]@{
        schema_version = 1
        marker_id = $MarkerId
        drive_folder_id = $DriveFolderId
        expected_owner = $ExpectedOwner
        classification = $script:GssDriveBackupClassificationLabel
        contains_personal_data = $true
        purpose = 'GSS survey recovery backups'
        created_at_utc = [datetime]::UtcNow.ToString('o')
    }

    if (Test-Path -LiteralPath $markerPath -PathType Leaf) {
        $existing = Read-GssDriveBackupJson -Path $markerPath
        if ([string](Get-GssDriveBackupProperty $existing @('marker_id')) -ne $MarkerId -or
            [string](Get-GssDriveBackupProperty $existing @('drive_folder_id')) -ne $DriveFolderId -or
            [string](Get-GssDriveBackupProperty $existing @('expected_owner')) -ne $ExpectedOwner) {
            throw "Existing backup root marker does not match the requested private Drive destination: $markerPath"
        }
    }
    else {
        Write-GssDriveBackupAtomicJson -Path $markerPath -Value $marker
    }

    $noticePath = Join-Path $root 'PRIVATE GSS BACKUPS - READ ME.txt'
    $notice = @"
PRIVATE GSS BACKUPS

Classification: $script:GssDriveBackupClassificationLabel
Expected owner: $ExpectedOwner
Purpose: GSS survey recovery backups.

Do not share this folder. Some snapshots contain raw guest detail and other personal data.
Retention is report-only: this automation never deletes Drive snapshots automatically.
Restore drills are verify-only under LOCALAPPDATA and never overwrite the live workbook.
"@
    Write-GssDriveBackupAtomicText -Path $noticePath -Text $notice

    $settings = [ordered]@{
        schema_version = 1
        drive_root_path = $root
        drive_folder_id = $DriveFolderId
        expected_owner = $ExpectedOwner
        marker_id = $MarkerId
        verification_level = 'drivefs_hash_verified'
        retention = [ordered]@{
            weekly = 13
            monthly = 12
        }
        require_before_apply = $true
    }
    Write-GssDriveBackupAtomicJson -Path $SettingsPath -Value $settings

    $verifiedContext = Get-GssDriveBackupRootContext -SettingsPath $SettingsPath -SkipCloudMetadataReadback
    $verifiedMarkerHash = Get-GssDriveBackupSha256 -Path $markerPath
    $verifiedSettingsHash = Get-GssDriveBackupSha256 -Path $verifiedContext.Settings.SettingsPath
    $verifiedNoticeHash = Get-GssDriveBackupSha256 -Path $noticePath

    return [pscustomobject]@{
        Status = 'Configured'
        SettingsPath = [System.IO.Path]::GetFullPath($SettingsPath)
        DriveRootPath = $root
        MarkerPath = $markerPath
        DriveFolderId = $DriveFolderId
        ExpectedOwner = $ExpectedOwner
        MarkerId = $MarkerId
        VerificationLevel = 'drivefs_hash_verified'
        MarkerSha256 = $verifiedMarkerHash
        SettingsSha256 = $verifiedSettingsHash
        ClassificationNoticePath = $noticePath
        ClassificationNoticeSha256 = $verifiedNoticeHash
        ReadbackVerified = $true
    }
}

function Get-GssDriveBackupSetting {
    [CmdletBinding()]
    param(
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath)
    )

    $settings = Read-GssDriveBackupJson -Path $SettingsPath
    $schemaVersion = [int](Get-GssDriveBackupProperty $settings @('schema_version'))
    $driveRoot = [string](Get-GssDriveBackupProperty $settings @('drive_root_path'))
    $driveFolderId = [string](Get-GssDriveBackupProperty $settings @('drive_folder_id'))
    $expectedOwner = [string](Get-GssDriveBackupProperty $settings @('expected_owner'))
    $markerId = [string](Get-GssDriveBackupProperty $settings @('marker_id'))
    $verificationLevel = [string](Get-GssDriveBackupProperty $settings @('verification_level'))
    $retention = Get-GssDriveBackupProperty $settings @('retention')
    $weekly = [int](Get-GssDriveBackupProperty $retention @('weekly'))
    $monthly = [int](Get-GssDriveBackupProperty $retention @('monthly'))
    $requireBeforeApply = [bool](Get-GssDriveBackupProperty $settings @('require_before_apply'))

    if ($schemaVersion -ne 1) { throw "Unsupported GSS Drive backup settings schema: $schemaVersion" }
    if ([string]::IsNullOrWhiteSpace($driveRoot) -or -not [System.IO.Path]::IsPathRooted($driveRoot)) {
        throw 'GSS Drive backup settings must contain an exact absolute drive_root_path.'
    }
    if ([string]::IsNullOrWhiteSpace($driveFolderId) -or
        [string]::IsNullOrWhiteSpace($expectedOwner) -or
        [string]::IsNullOrWhiteSpace($markerId)) {
        throw 'GSS Drive backup settings are missing destination identity fields.'
    }
    if ($weekly -ne 13 -or $monthly -ne 12) {
        throw "GSS Drive backup retention must remain 13 weekly and 12 monthly; configured values are $weekly and $monthly."
    }
    if (-not $requireBeforeApply) {
        throw 'GSS Drive backup settings must require a successful preparation before live apply.'
    }
    if ($verificationLevel -ne 'drivefs_hash_verified') {
        throw "Unsupported Drive backup verification level: $verificationLevel"
    }

    return [pscustomobject]@{
        SchemaVersion = 1
        SettingsPath = [System.IO.Path]::GetFullPath($SettingsPath)
        DriveRootPath = [System.IO.Path]::GetFullPath($driveRoot).TrimEnd('\', '/')
        DriveFolderId = $driveFolderId
        ExpectedOwner = $expectedOwner
        MarkerId = $markerId
        VerificationLevel = $verificationLevel
        RetentionWeekly = $weekly
        RetentionMonthly = $monthly
        RequireBeforeApply = $requireBeforeApply
    }
}

function Get-GssDriveBackupRootContext {
    [CmdletBinding()]
    param(
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath),
        [switch]$SkipCloudMetadataReadback
    )

    $settings = Get-GssDriveBackupSetting -SettingsPath $SettingsPath
    if (-not (Test-Path -LiteralPath $settings.DriveRootPath -PathType Container)) {
        throw "Configured Google Drive backup root is unavailable: $($settings.DriveRootPath). No fallback destination is allowed."
    }

    $markerPath = Join-Path $settings.DriveRootPath 'backup-root.json'
    $marker = Read-GssDriveBackupJson -Path $markerPath
    if ([int](Get-GssDriveBackupProperty $marker @('schema_version')) -ne 1 -or
        [string](Get-GssDriveBackupProperty $marker @('marker_id')) -ne $settings.MarkerId -or
        [string](Get-GssDriveBackupProperty $marker @('drive_folder_id')) -ne $settings.DriveFolderId -or
        [string](Get-GssDriveBackupProperty $marker @('expected_owner')) -ne $settings.ExpectedOwner -or
        -not [bool](Get-GssDriveBackupProperty $marker @('contains_personal_data'))) {
        throw "Configured Google Drive backup root marker failed identity or classification validation: $markerPath"
    }

    $readbackPath = Join-Path $settings.DriveRootPath 'commissioning-readback.json'
    $readback = $null
    if (-not $SkipCloudMetadataReadback) {
        $readback = Read-GssDriveBackupJson -Path $readbackPath
        if ([int](Get-GssDriveBackupProperty $readback @('schema_version')) -ne 1 -or
            [string](Get-GssDriveBackupProperty $readback @('drive_folder_id')) -ne $settings.DriveFolderId -or
            [string](Get-GssDriveBackupProperty $readback @('owner')) -ne $settings.ExpectedOwner -or
            [bool](Get-GssDriveBackupProperty $readback @('shared')) -or
            [int](Get-GssDriveBackupProperty $readback @('permission_count')) -ne 1 -or
            [string](Get-GssDriveBackupProperty $readback @('verification')) -ne 'cloud_metadata_readback_verified' -or
            [string](Get-GssDriveBackupProperty $readback @('marker_id')) -ne $settings.MarkerId -or
            [string](Get-GssDriveBackupProperty $readback @('classification')) -ne $script:GssDriveBackupClassificationLabel) {
            throw "Configured Google Drive backup root failed its owner-only connector metadata evidence: $readbackPath"
        }
    }

    return [pscustomobject]@{
        Settings = $settings
        Marker = $marker
        MarkerPath = $markerPath
        CommissioningReadback = $readback
        CommissioningReadbackPath = $readbackPath
        RootPath = $settings.DriveRootPath
    }
}

function Write-GssDriveBackupCommissioningReadback {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DriveFolderId,
        [Parameter(Mandatory)]
        [string]$Owner,
        [Parameter(Mandatory)]
        [bool]$Shared,
        [Parameter(Mandatory)]
        [int]$PermissionCount,
        [datetime]$VerifiedAtUtc = [datetime]::UtcNow,
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath)
    )

    $context = Get-GssDriveBackupRootContext -SettingsPath $SettingsPath -SkipCloudMetadataReadback
    if ($DriveFolderId -ne $context.Settings.DriveFolderId) {
        throw 'Connector metadata readback folder ID does not match the commissioned settings.'
    }
    if ($Owner -ne $context.Settings.ExpectedOwner) {
        throw 'Connector metadata readback owner does not match the commissioned settings.'
    }
    if ($Shared) {
        throw 'Connector metadata readback says the GSS backup folder is shared; commissioning evidence was refused.'
    }
    if ($PermissionCount -ne 1) {
        throw "Connector metadata readback must show exactly one owner permission; observed $PermissionCount."
    }

    $readback = [ordered]@{
        schema_version = 1
        drive_folder_id = $DriveFolderId
        owner = $Owner
        shared = $false
        permission_count = 1
        verified_at_utc = $VerifiedAtUtc.ToUniversalTime().ToString('o')
        source = 'google_drive_connector'
        verification = 'cloud_metadata_readback_verified'
        marker_id = $context.Settings.MarkerId
        classification = $script:GssDriveBackupClassificationLabel
    }
    $path = Join-Path $context.RootPath 'commissioning-readback.json'
    Write-GssDriveBackupAtomicJson -Path $path -Value $readback
    $verified = Read-GssDriveBackupJson -Path $path
    if ([string](Get-GssDriveBackupProperty $verified @('drive_folder_id')) -ne $context.Settings.DriveFolderId -or
        [string](Get-GssDriveBackupProperty $verified @('owner')) -ne $context.Settings.ExpectedOwner -or
        [bool](Get-GssDriveBackupProperty $verified @('shared')) -or
        [int](Get-GssDriveBackupProperty $verified @('permission_count')) -ne 1 -or
        [string](Get-GssDriveBackupProperty $verified @('verification')) -ne 'cloud_metadata_readback_verified') {
        throw 'Commissioning metadata readback failed its atomic write/readback verification.'
    }

    return [pscustomobject]@{
        Status = 'Recorded'
        Path = $path
        Sha256 = Get-GssDriveBackupSha256 -Path $path
        DriveFolderId = $DriveFolderId
        Owner = $Owner
        Shared = $false
        PermissionCount = 1
        Verification = 'cloud_metadata_readback_verified'
    }
}

function Test-GssDriveBackupExcludedPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath
    )

    $segments = @($RelativePath.Replace('\', '/') -split '/' | Where-Object { $_ })
    foreach ($segment in $segments) {
        $lower = $segment.ToLowerInvariant()
        if ($lower -in @(
            'gss survey workbook automation',
            '.git',
            'test-output',
            'backups',
            'quarantine',
            'staging',
            'temp',
            'tmp'
        ) -or
        $lower.StartsWith('.staging') -or
        $lower.StartsWith('.partial-') -or
        $lower.EndsWith('.partial') -or
        $lower.EndsWith('.tmp')) {
            return $true
        }
    }
    return $false
}

function ConvertTo-GssDriveBackupInventoryRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,
        [Parameter(Mandatory)]
        [string]$PortablePath,
        [Parameter(Mandatory)]
        [string]$Role,
        [Parameter(Mandatory)]
        [string]$Classification
    )

    $source = [System.IO.Path]::GetFullPath($SourcePath)
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
        throw "Backup inventory source is missing or not locally available: $source"
    }

    return [pscustomobject]@{
        SourcePath = $source
        PortablePath = Assert-GssDriveBackupSafeRelativePath -Path $PortablePath
        Role = $Role
        Classification = $Classification
    }
}

function Get-GssDriveBackupClassification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath
    )

    if ($RelativePath -match '(?i)(^|[\\/])03 Uploaded Survey Workbooks([\\/]|$)|guest[ _-]*detail|raw[ _-]*detail') {
        return 'restricted_personal_data'
    }
    return 'restricted_operational'
}

function Get-GssDriveBackupInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GssRoot,
        [object[]]$AdditionalItems = @(),
        [string[]]$TransactionArtifactPaths = @(),
        [string[]]$ReleaseArchivePaths = @()
    )

    $root = [System.IO.Path]::GetFullPath($GssRoot).TrimEnd('\', '/')
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        throw "GSS source root is unavailable: $root"
    }

    $records = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $addRecord = {
        param($record)
        $key = $record.PortablePath.ToLowerInvariant()
        if ($seen.ContainsKey($key)) {
            $existing = $seen[$key]
            if (-not $existing.SourcePath.Equals($record.SourcePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Backup inventory has a portable path collision at '$($record.PortablePath)'."
            }
            return
        }
        $seen[$key] = $record
        $records.Add($record)
    }

    foreach ($rootFileName in @('Run GSS Update After Upload.cmd', '00 START HERE - GSS Survey Updates.txt')) {
        $source = Join-Path $root $rootFileName
        if (Test-Path -LiteralPath $source -PathType Leaf) {
            & $addRecord (ConvertTo-GssDriveBackupInventoryRecord -SourcePath $source -PortablePath "gss/$rootFileName" -Role 'operator_control' -Classification 'restricted_operational')
        }
    }

    $operationalFolders = @(
        '01 Main Workbook',
        '02 Weekly Rolling Source Workbooks',
        '03 Uploaded Survey Workbooks',
        '04 Email Comparison PDFs',
        '05 Reference Materials',
        '06 Exports and Images'
    )
    foreach ($folderName in $operationalFolders) {
        $folder = Join-Path $root $folderName
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        foreach ($file in Get-ChildItem -LiteralPath $folder -Recurse -File) {
            $relative = Get-GssDriveBackupRelativePath -Path $file.FullName -Root $root
            if (Test-GssDriveBackupExcludedPath -RelativePath $relative) { continue }
            $role = 'operational_' + ($folderName.Substring(0, 2))
            $classification = Get-GssDriveBackupClassification -RelativePath $relative
            & $addRecord (ConvertTo-GssDriveBackupInventoryRecord -SourcePath $file.FullName -PortablePath "gss/$relative" -Role $role -Classification $classification)
        }
    }

    foreach ($runFolderName in @('qa', 'logs', 'state')) {
        $folder = Join-Path $root "_automation_runs\$runFolderName"
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        foreach ($file in Get-ChildItem -LiteralPath $folder -Recurse -File) {
            $relative = Get-GssDriveBackupRelativePath -Path $file.FullName -Root $root
            if (Test-GssDriveBackupExcludedPath -RelativePath $relative) { continue }
            & $addRecord (ConvertTo-GssDriveBackupInventoryRecord -SourcePath $file.FullName -PortablePath "gss/$relative" -Role $runFolderName -Classification (Get-GssDriveBackupClassification $relative))
        }
    }

    $outbox = Join-Path $root '_automation_runs\email_outbox'
    if (Test-Path -LiteralPath $outbox -PathType Container) {
        foreach ($packageFolder in Get-ChildItem -LiteralPath $outbox -Directory) {
            if ($packageFolder.Name.StartsWith('.staging', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if (-not (Test-Path -LiteralPath (Join-Path $packageFolder.FullName 'READY') -PathType Leaf)) { continue }
            foreach ($file in Get-ChildItem -LiteralPath $packageFolder.FullName -Recurse -File) {
                $relative = Get-GssDriveBackupRelativePath -Path $file.FullName -Root $root
                if (Test-GssDriveBackupExcludedPath -RelativePath $relative) { continue }
                & $addRecord (ConvertTo-GssDriveBackupInventoryRecord -SourcePath $file.FullName -PortablePath "gss/$relative" -Role 'ready_package' -Classification (Get-GssDriveBackupClassification $relative))
            }
        }
    }

    $transactionIndex = 0
    foreach ($artifactPath in @($TransactionArtifactPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $transactionIndex++
        $source = [System.IO.Path]::GetFullPath($artifactPath)
        $portable = "transaction/{0:d2}-{1}" -f $transactionIndex, [System.IO.Path]::GetFileName($source)
        & $addRecord (ConvertTo-GssDriveBackupInventoryRecord -SourcePath $source -PortablePath $portable -Role 'transaction_artifact' -Classification (Get-GssDriveBackupClassification $source))
    }

    $releaseIndex = 0
    foreach ($archivePath in @($ReleaseArchivePaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $releaseIndex++
        $source = [System.IO.Path]::GetFullPath($archivePath)
        $portable = "release/{0:d2}-{1}" -f $releaseIndex, [System.IO.Path]::GetFileName($source)
        & $addRecord (ConvertTo-GssDriveBackupInventoryRecord -SourcePath $source -PortablePath $portable -Role 'release_archive' -Classification 'restricted_operational')
    }

    foreach ($item in @($AdditionalItems)) {
        if ($null -eq $item) { continue }
        if ($item -is [string]) {
            $source = [System.IO.Path]::GetFullPath([string]$item)
            $portable = "additional/$([System.IO.Path]::GetFileName($source))"
            $role = 'additional_artifact'
            $classification = Get-GssDriveBackupClassification $source
        }
        else {
            $source = [string](Get-GssDriveBackupProperty $item @('SourcePath', 'source_path', 'Path', 'path'))
            $portable = [string](Get-GssDriveBackupProperty $item @('PortablePath', 'portable_path'))
            $role = [string](Get-GssDriveBackupProperty $item @('Role', 'role') 'additional_artifact')
            $classification = [string](Get-GssDriveBackupProperty $item @('Classification', 'classification') (Get-GssDriveBackupClassification $source))
            if ([string]::IsNullOrWhiteSpace($portable)) {
                $portable = "additional/$([System.IO.Path]::GetFileName($source))"
            }
        }
        & $addRecord (ConvertTo-GssDriveBackupInventoryRecord -SourcePath $source -PortablePath $portable -Role $role -Classification $classification)
    }

    return @($records | Sort-Object PortablePath)
}

function Enter-GssDriveBackupMutex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$MarkerId
    )

    $safeName = [regex]::Replace($MarkerId, '[^A-Za-z0-9_-]', '_')
    $mutex = New-Object System.Threading.Mutex($false, "Local\GSSDriveBackup-$safeName")
    try {
        if (-not $mutex.WaitOne(0)) {
            $mutex.Dispose()
            throw 'Another GSS Drive backup operation is already running on this workstation.'
        }
    }
    catch [System.Threading.AbandonedMutexException] {
        # Ownership is granted when an abandoned mutex is observed.
        Write-Verbose 'Acquired an abandoned GSS Drive backup mutex.'
    }
    return $mutex
}

function Exit-GssDriveBackupMutex {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [System.Threading.Mutex]$Mutex
    )

    if ($null -ne $Mutex) {
        try { $Mutex.ReleaseMutex() } catch { Write-Verbose "Mutex release was already completed: $($_.Exception.Message)" }
        $Mutex.Dispose()
    }
}

function Write-GssDriveBackupStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Directory,
        [Parameter(Mandatory)]
        [ValidateSet('Prepared', 'Committed', 'PendingFinalize', 'Blocked', 'Aborted')]
        [string]$Status,
        [Parameter(Mandatory)]
        [string]$RunId,
        [string]$Fingerprint,
        [string]$Message
    )

    $value = [ordered]@{
        schema_version = 1
        status = $Status
        run_id = $RunId
        fingerprint = $Fingerprint
        updated_at_utc = [datetime]::UtcNow.ToString('o')
        message = $Message
    }
    Write-GssDriveBackupAtomicJson -Path (Join-Path $Directory 'backup-status.json') -Value $value
}

function Copy-GssDriveBackupInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Inventory,
        [Parameter(Mandatory)]
        [string]$SnapshotDirectory,
        [Parameter(Mandatory)]
        [string]$PayloadPrefix
    )

    $safePayloadPrefix = Assert-GssDriveBackupSafeRelativePath -Path $PayloadPrefix
    $results = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($item in @($Inventory | Sort-Object PortablePath)) {
        $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string]$item.PortablePath)
        $snapshotRelative = Assert-GssDriveBackupSafeRelativePath -Path "$safePayloadPrefix/$portable"
        $key = $snapshotRelative.ToLowerInvariant()
        if ($seen.ContainsKey($key)) {
            throw "Duplicate snapshot path in backup inventory: $snapshotRelative"
        }
        $seen[$key] = $true

        $source = [System.IO.Path]::GetFullPath([string]$item.SourcePath)
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
            throw "Backup source disappeared before copy: $source"
        }
        $sourceInfo = Get-Item -LiteralPath $source
        $sourceHashBefore = Get-GssDriveBackupSha256 -Path $source
        $destination = Join-Path $SnapshotDirectory $snapshotRelative.Replace('/', '\')
        $destinationParent = Split-Path -Parent $destination
        if (-not (Test-Path -LiteralPath $destinationParent -PathType Container)) {
            New-Item -ItemType Directory -Path $destinationParent -Force | Out-Null
        }

        $partial = Join-Path $destinationParent ('.t-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        if (Test-Path -LiteralPath $partial) {
            throw "A stale partial backup file blocks safe preparation: $partial"
        }
        Copy-Item -LiteralPath $source -Destination $partial
        $destinationHash = Get-GssDriveBackupSha256 -Path $partial
        $sourceHashAfter = Get-GssDriveBackupSha256 -Path $source
        if ($sourceHashBefore -ne $sourceHashAfter) {
            throw "Backup source changed during copy: $source"
        }
        if ($sourceHashBefore -ne $destinationHash) {
            throw "DriveFS destination hash mismatch for: $snapshotRelative"
        }

        if (Test-Path -LiteralPath $destination -PathType Leaf) {
            $replacementBackupPath = Join-Path $destinationParent ('.b-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            try {
                [System.IO.File]::Replace($partial, $destination, $replacementBackupPath, $true)
            }
            finally {
                if (Test-Path -LiteralPath $replacementBackupPath -PathType Leaf) {
                    Remove-Item -LiteralPath $replacementBackupPath -Force
                }
            }
        }
        else {
            [System.IO.File]::Move($partial, $destination)
        }
        $verifiedHash = Get-GssDriveBackupSha256 -Path $destination
        if ($verifiedHash -ne $sourceHashBefore) {
            throw "DriveFS destination changed after promotion: $snapshotRelative"
        }

        $results.Add([pscustomobject][ordered]@{
            role = [string]$item.Role
            portable_path = $portable
            snapshot_path = $snapshotRelative
            byte_size = [long]$sourceInfo.Length
            sha256 = $sourceHashBefore
            classification = [string]$item.Classification
        })
    }
    return ($results | ForEach-Object { $_ })
}

function Test-GssDriveBackupPayload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SnapshotDirectory,
        [Parameter(Mandatory)]
        [object[]]$Files
    )

    $root = [System.IO.Path]::GetFullPath($SnapshotDirectory).TrimEnd('\', '/')
    $validated = 0
    foreach ($file in @($Files)) {
        $relative = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $file @('snapshot_path')))
        $path = [System.IO.Path]::GetFullPath((Join-Path $root $relative.Replace('/', '\')))
        if (-not $path.StartsWith("$root\", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Backup manifest path escaped its snapshot root: $relative"
        }
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Backup payload file is missing: $relative"
        }
        $expectedBytes = [long](Get-GssDriveBackupProperty $file @('byte_size'))
        if ((Get-Item -LiteralPath $path).Length -ne $expectedBytes) {
            throw "Backup payload size mismatch: $relative"
        }
        $expectedHash = [string](Get-GssDriveBackupProperty $file @('sha256'))
        if ((Get-GssDriveBackupSha256 -Path $path) -ne $expectedHash) {
            throw "Backup payload hash mismatch: $relative"
        }
        $validated++
    }
    return $validated
}

function Get-GssDriveBackupChainHead {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RootPath
    )

    $headPath = Join-Path $RootPath 'chain-head.json'
    if (-not (Test-Path -LiteralPath $headPath -PathType Leaf)) {
        return $null
    }
    $head = Read-GssDriveBackupJson -Path $headPath
    $relative = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $head @('backup_manifest_relative_path')))
    $manifestPath = [System.IO.Path]::GetFullPath((Join-Path $RootPath $relative.Replace('/', '\')))
    $root = [System.IO.Path]::GetFullPath($RootPath).TrimEnd('\', '/')
    if (-not $manifestPath.StartsWith("$root\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Drive backup chain head escaped the configured root.'
    }
    $expected = [string](Get-GssDriveBackupProperty $head @('backup_manifest_sha256'))
    if ((Get-GssDriveBackupSha256 -Path $manifestPath) -ne $expected) {
        throw 'Drive backup chain head does not match its referenced manifest.'
    }
    return [pscustomobject]@{
        ManifestPath = $manifestPath
        ManifestRelativePath = $relative
        ManifestSha256 = $expected
        RunId = [string](Get-GssDriveBackupProperty $head @('run_id'))
    }
}

function Set-GssDriveBackupChainHead {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)]
        [string]$RootPath,
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [string]$Fingerprint,
        [Parameter(Mandatory)]
        [string]$CommittedAtUtc,
        [Parameter(Mandatory)]
        [string]$ManifestPath,
        [Parameter(Mandatory)]
        [string]$ManifestSha256
    )

    $manifestRelative = Get-GssDriveBackupRelativePath -Path $ManifestPath -Root $RootPath
    $chainHead = [ordered]@{
        schema_version = 1
        run_id = $RunId
        fingerprint = $Fingerprint
        committed_at_utc = $CommittedAtUtc
        backup_manifest_relative_path = $manifestRelative
        backup_manifest_sha256 = $ManifestSha256
    }
    $headPath = Join-Path $RootPath 'chain-head.json'
    if ($PSCmdlet.ShouldProcess($headPath, "Record committed backup chain head for run '$RunId'")) {
        Write-GssDriveBackupAtomicJson -Path $headPath -Value $chainHead
    }
}

function New-GssDriveBackupPreparedSnapshot {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [string]$Fingerprint,
        [Parameter(Mandatory)]
        [datetime]$ReportWeek,
        [Parameter(Mandatory)]
        [object[]]$Inventory,
        [string]$Release = 'unversioned',
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath)
    )

    if ($RunId -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]{5,127}$') {
        throw "Run ID is not safe for a Drive snapshot path: $RunId"
    }
    if ([string]::IsNullOrWhiteSpace($Fingerprint)) {
        throw 'A non-empty run fingerprint is required before Drive backup preparation.'
    }
    if (@($Inventory).Count -eq 0) {
        throw 'The curated Drive backup inventory is empty.'
    }

    $context = Get-GssDriveBackupRootContext -SettingsPath $SettingsPath
    $mutex = $null
    try {
        $mutex = Enter-GssDriveBackupMutex -MarkerId $context.Settings.MarkerId
        $partialPath = Join-Path $context.RootPath ".partial-$RunId"
        $preparedPath = Join-Path $partialPath 'prepared-manifest.json'
        if (Test-Path -LiteralPath $partialPath -PathType Container) {
            $statusPath = Join-Path $partialPath 'backup-status.json'
            if (Test-Path -LiteralPath $statusPath -PathType Leaf) {
                $existingStatus = [string](Get-GssDriveBackupProperty (Read-GssDriveBackupJson $statusPath) @('status'))
                if ($existingStatus -eq 'Aborted') {
                    throw "Prepared Drive backup was explicitly aborted and cannot be reused: $partialPath"
                }
            }
            if (-not (Test-Path -LiteralPath $preparedPath -PathType Leaf)) {
                throw "Incomplete backup staging directory blocks preparation: $partialPath"
            }
            $existing = Read-GssDriveBackupJson -Path $preparedPath
            if ([string](Get-GssDriveBackupProperty $existing @('run_id')) -ne $RunId -or
                [string](Get-GssDriveBackupProperty $existing @('fingerprint')) -ne $Fingerprint) {
                throw "Existing prepared backup does not match the requested run: $partialPath"
            }
            $files = @(Get-GssDriveBackupProperty $existing @('files') @())
            [void](Test-GssDriveBackupPayload -SnapshotDirectory $partialPath -Files $files)
            return [pscustomobject]@{
                Status = 'Prepared'
                BackupStatus = 'Prepared'
                RunId = $RunId
                Fingerprint = $Fingerprint
                PreparedPath = $partialPath
                PreparedManifestPath = $preparedPath
                PreparedManifestSha256 = Get-GssDriveBackupSha256 $preparedPath
                VerificationLevel = $context.Settings.VerificationLevel
                Idempotent = $true
            }
        }

        if (-not $PSCmdlet.ShouldProcess($partialPath, "Prepare hash-verified Drive backup for run '$RunId'")) {
            return [pscustomobject]@{
                Status = 'Blocked'
                BackupStatus = 'Blocked'
                RunId = $RunId
                Fingerprint = $Fingerprint
                PreparedPath = $partialPath
                VerificationLevel = $context.Settings.VerificationLevel
                WhatIf = $true
            }
        }
        New-Item -ItemType Directory -Path $partialPath | Out-Null
        try {
            $files = @(Copy-GssDriveBackupInventory -Inventory $Inventory -SnapshotDirectory $partialPath -PayloadPrefix 'prepared-payload')
            $chainHead = Get-GssDriveBackupChainHead -RootPath $context.RootPath
            $prepared = [ordered]@{
                schema_version = 1
                manifest_type = 'prepared'
                status = 'Prepared'
                run_id = $RunId
                fingerprint = $Fingerprint
                report_week = $ReportWeek.Date.ToString('yyyy-MM-dd')
                prepared_at_utc = [datetime]::UtcNow.ToString('o')
                host = [Environment]::MachineName
                release = $Release
                prior_manifest_sha256 = if ($chainHead) { $chainHead.ManifestSha256 } else { $null }
                drive = [ordered]@{
                    folder_id = $context.Settings.DriveFolderId
                    marker_id = $context.Settings.MarkerId
                    expected_owner = $context.Settings.ExpectedOwner
                    verification_level = $context.Settings.VerificationLevel
                }
                data_classification = [ordered]@{
                    label = $script:GssDriveBackupClassificationLabel
                    contains_personal_data = $true
                }
                scope = [ordered]@{
                    included_roles = @($files.role | Sort-Object -Unique)
                    excluded = @('test-output', '_automation_runs/backups except explicit transaction artifacts', 'quarantine', 'staging', 'temp', 'repository working tree', '.git')
                }
                file_count = @($files).Count
                total_bytes = [long](($files | Measure-Object -Property byte_size -Sum).Sum)
                files = $files
            }
            Write-GssDriveBackupStatus -Directory $partialPath -Status Prepared -RunId $RunId -Fingerprint $Fingerprint -Message 'DriveFS payload hashes verified; prepared manifest is the preparation commit point.'
            Write-GssDriveBackupAtomicJson -Path $preparedPath -Value $prepared
        }
        catch {
            try {
                Write-GssDriveBackupStatus -Directory $partialPath -Status Aborted -RunId $RunId -Fingerprint $Fingerprint -Message $_.Exception.Message
            }
            catch { Write-Verbose "Could not record Aborted status after preparation failure: $($_.Exception.Message)" }
            throw
        }

        return [pscustomobject]@{
            Status = 'Prepared'
            BackupStatus = 'Prepared'
            RunId = $RunId
            Fingerprint = $Fingerprint
            PreparedPath = $partialPath
            PreparedManifestPath = $preparedPath
            PreparedManifestSha256 = Get-GssDriveBackupSha256 $preparedPath
            FileCount = @($files).Count
            TotalBytes = [long](($files | Measure-Object -Property byte_size -Sum).Sum)
            VerificationLevel = $context.Settings.VerificationLevel
            Idempotent = $false
        }
    }
    finally {
        Exit-GssDriveBackupMutex -Mutex $mutex
    }
}

function Get-GssDriveBackupRunStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RunId,
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath)
    )

    $context = Get-GssDriveBackupRootContext -SettingsPath $SettingsPath
    $location = Find-GssDriveBackupSnapshot -RootPath $context.RootPath -RunId $RunId
    if ($null -eq $location) { return $null }
    $statusPath = Join-Path $location.Path 'backup-status.json'
    if (-not (Test-Path -LiteralPath $statusPath -PathType Leaf)) { return $null }
    return [string](Get-GssDriveBackupProperty (Read-GssDriveBackupJson -Path $statusPath) @('status'))
}

function Stop-GssDriveBackupPreparedSnapshot {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [string]$Fingerprint,
        [string]$Message = 'Transaction ended before a committed Drive snapshot was required.',
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath)
    )

    $context = Get-GssDriveBackupRootContext -SettingsPath $SettingsPath
    $mutex = $null
    try {
        $mutex = Enter-GssDriveBackupMutex -MarkerId $context.Settings.MarkerId
        $location = Find-GssDriveBackupSnapshot -RootPath $context.RootPath -RunId $RunId
        if ($null -eq $location) {
            throw "No prepared Drive backup exists to abort for run '$RunId'."
        }
        if (-not $location.IsPartial) {
            throw "Run '$RunId' has a committed snapshot and cannot be marked Aborted."
        }
        $preparedPath = Join-Path $location.Path 'prepared-manifest.json'
        $prepared = Read-GssDriveBackupJson -Path $preparedPath
        if ([string](Get-GssDriveBackupProperty $prepared @('run_id')) -ne $RunId -or
            [string](Get-GssDriveBackupProperty $prepared @('fingerprint')) -ne $Fingerprint) {
            throw 'Prepared snapshot identity does not match the requested abort.'
        }
        [void](Test-GssDriveBackupPayload -SnapshotDirectory $location.Path -Files @(Get-GssDriveBackupProperty $prepared @('files') @()))

        $statusPath = Join-Path $location.Path 'backup-status.json'
        if (Test-Path -LiteralPath $statusPath -PathType Leaf) {
            $status = [string](Get-GssDriveBackupProperty (Read-GssDriveBackupJson -Path $statusPath) @('status'))
            if ($status -eq 'Committed') {
                throw "Run '$RunId' is already committed and cannot be marked Aborted."
            }
            if ($status -eq 'Aborted') {
                return [pscustomobject]@{
                    Status = 'Aborted'
                    BackupStatus = 'Aborted'
                    RunId = $RunId
                    Fingerprint = $Fingerprint
                    PreparedPath = $location.Path
                    EvidenceRetained = $true
                    Idempotent = $true
                }
            }
        }

        if (-not $PSCmdlet.ShouldProcess($location.Path, "Mark prepared Drive backup Aborted for run '$RunId' without deleting evidence")) {
            return [pscustomobject]@{
                Status = 'Prepared'
                BackupStatus = 'Prepared'
                RunId = $RunId
                Fingerprint = $Fingerprint
                PreparedPath = $location.Path
                EvidenceRetained = $true
                WhatIf = $true
            }
        }
        Write-GssDriveBackupStatus -Directory $location.Path -Status Aborted -RunId $RunId -Fingerprint $Fingerprint -Message $Message
        return [pscustomobject]@{
            Status = 'Aborted'
            BackupStatus = 'Aborted'
            RunId = $RunId
            Fingerprint = $Fingerprint
            PreparedPath = $location.Path
            EvidenceRetained = $true
            Idempotent = $false
        }
    }
    finally {
        Exit-GssDriveBackupMutex -Mutex $mutex
    }
}

function Find-GssDriveBackupSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RootPath,
        [Parameter(Mandatory)]
        [string]$RunId
    )

    $partial = Join-Path $RootPath ".partial-$RunId"
    $committedMatches = @()
    $snapshotsRoot = Join-Path $RootPath 'snapshots'
    if (Test-Path -LiteralPath $snapshotsRoot -PathType Container) {
        $committedMatches = @(Get-ChildItem -LiteralPath $snapshotsRoot -Directory -Recurse |
            Where-Object {
                $_.Name -eq $RunId -and
                (Test-Path -LiteralPath (Join-Path $_.FullName 'prepared-manifest.json') -PathType Leaf)
            })
    }

    if ($committedMatches.Count -gt 1) {
        throw "More than one committed snapshot directory was found for run '$RunId'."
    }
    if ((Test-Path -LiteralPath $partial -PathType Container) -and $committedMatches.Count -eq 1) {
        throw "Both prepared and committed-looking directories exist for run '$RunId'; manual review is required."
    }
    if ($committedMatches.Count -eq 1) {
        return [pscustomobject]@{ Path = $committedMatches[0].FullName; IsPartial = $false }
    }
    if (Test-Path -LiteralPath $partial -PathType Container) {
        return [pscustomobject]@{ Path = $partial; IsPartial = $true }
    }
    return $null
}

function Test-GssCommittedBackupSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SnapshotPath
    )

    $receiptPath = Join-Path $SnapshotPath 'commit-receipt.json'
    $manifestPath = Join-Path $SnapshotPath 'backup-manifest.json'
    $preparedManifestPath = Join-Path $SnapshotPath 'prepared-manifest.json'
    $receipt = Read-GssDriveBackupJson -Path $receiptPath
    $manifest = Read-GssDriveBackupJson -Path $manifestPath
    $preparedManifest = Read-GssDriveBackupJson -Path $preparedManifestPath
    $expectedManifestHash = [string](Get-GssDriveBackupProperty $receipt @('backup_manifest_sha256'))
    $actualManifestHash = Get-GssDriveBackupSha256 -Path $manifestPath
    if ($actualManifestHash -ne $expectedManifestHash) {
        throw "Committed backup manifest hash does not match its receipt: $SnapshotPath"
    }
    $actualPreparedManifestHash = Get-GssDriveBackupSha256 -Path $preparedManifestPath
    $receiptPreparedHash = [string](Get-GssDriveBackupProperty $receipt @('prepared_manifest_sha256'))
    $manifestPreparedHash = [string](Get-GssDriveBackupProperty $manifest @('prepared_manifest_sha256'))
    if ($actualPreparedManifestHash -ne $receiptPreparedHash -or
        $actualPreparedManifestHash -ne $manifestPreparedHash) {
        throw "Prepared backup manifest hash does not match the committed evidence: $SnapshotPath"
    }
    $files = @(Get-GssDriveBackupProperty $manifest @('files') @())
    $validatedCount = Test-GssDriveBackupPayload -SnapshotDirectory $SnapshotPath -Files $files
    $preparedFiles = @(Get-GssDriveBackupProperty $preparedManifest @('files') @())
    $validatedPreparedCount = Test-GssDriveBackupPayload -SnapshotDirectory $SnapshotPath -Files $preparedFiles
    return [pscustomobject]@{
        Receipt = $receipt
        Manifest = $manifest
        PreparedManifest = $preparedManifest
        ReceiptPath = $receiptPath
        ManifestPath = $manifestPath
        PreparedManifestPath = $preparedManifestPath
        ManifestSha256 = $actualManifestHash
        PreparedManifestSha256 = $actualPreparedManifestHash
        ValidatedFileCount = $validatedCount
        ValidatedPreparedFileCount = $validatedPreparedCount
    }
}

function Complete-GssDriveBackupSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [string]$Fingerprint,
        [object[]]$FinalInventory,
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath)
    )

    $context = Get-GssDriveBackupRootContext -SettingsPath $SettingsPath
    $mutex = $null
    $activePath = $null
    try {
        $mutex = Enter-GssDriveBackupMutex -MarkerId $context.Settings.MarkerId
        $location = Find-GssDriveBackupSnapshot -RootPath $context.RootPath -RunId $RunId
        if ($null -eq $location) {
            throw "No prepared Drive backup exists for run '$RunId'."
        }
        $activePath = $location.Path
        $statusPath = Join-Path $activePath 'backup-status.json'
        if (Test-Path -LiteralPath $statusPath -PathType Leaf) {
            $currentStatus = [string](Get-GssDriveBackupProperty (Read-GssDriveBackupJson -Path $statusPath) @('status'))
            if ($currentStatus -eq 'Aborted') {
                throw "Prepared Drive backup was explicitly aborted and cannot be finalized: $activePath"
            }
        }

        $receiptPath = Join-Path $activePath 'commit-receipt.json'
        if (Test-Path -LiteralPath $receiptPath -PathType Leaf) {
            $validated = Test-GssCommittedBackupSnapshot -SnapshotPath $activePath
            if ([string](Get-GssDriveBackupProperty $validated.Manifest @('run_id')) -ne $RunId -or
                [string](Get-GssDriveBackupProperty $validated.Manifest @('fingerprint')) -ne $Fingerprint) {
                throw 'Committed snapshot identity does not match the requested run.'
            }
            $head = Get-GssDriveBackupChainHead -RootPath $context.RootPath
            $priorHash = [string](Get-GssDriveBackupProperty $validated.Manifest @('prior_manifest_sha256'))
            $headHash = if ($head) { $head.ManifestSha256 } else { '' }
            if ($headHash -ne $validated.ManifestSha256 -and
                (($headHash -eq '' -and $priorHash -eq '') -or $headHash -eq $priorHash)) {
                Set-GssDriveBackupChainHead `
                    -RootPath $context.RootPath `
                    -RunId $RunId `
                    -Fingerprint $Fingerprint `
                    -CommittedAtUtc ([string](Get-GssDriveBackupProperty $validated.Receipt @('committed_at_utc'))) `
                    -ManifestPath $validated.ManifestPath `
                    -ManifestSha256 $validated.ManifestSha256
            }
            return [pscustomobject]@{
                Status = 'Committed'
                BackupStatus = 'Committed'
                RunId = $RunId
                Fingerprint = $Fingerprint
                SnapshotPath = $activePath
                BackupManifestPath = $validated.ManifestPath
                BackupManifestSha256 = $validated.ManifestSha256
                CommitReceiptPath = $validated.ReceiptPath
                FileCount = $validated.ValidatedFileCount
                VerificationLevel = $context.Settings.VerificationLevel
                Idempotent = $true
            }
        }

        $preparedPath = Join-Path $activePath 'prepared-manifest.json'
        $prepared = Read-GssDriveBackupJson -Path $preparedPath
        if ([string](Get-GssDriveBackupProperty $prepared @('run_id')) -ne $RunId -or
            [string](Get-GssDriveBackupProperty $prepared @('fingerprint')) -ne $Fingerprint) {
            throw 'Prepared snapshot identity does not match the requested run.'
        }
        $preparedFiles = @(Get-GssDriveBackupProperty $prepared @('files') @())
        [void](Test-GssDriveBackupPayload -SnapshotDirectory $activePath -Files $preparedFiles)

        $manifestPath = Join-Path $activePath 'backup-manifest.json'
        if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
            $backupManifest = Read-GssDriveBackupJson -Path $manifestPath
            $finalFiles = @(Get-GssDriveBackupProperty $backupManifest @('files') @())
            [void](Test-GssDriveBackupPayload -SnapshotDirectory $activePath -Files $finalFiles)
            $manifestPreparedFiles = @(Get-GssDriveBackupProperty $backupManifest @('prepared_files') @())
            [void](Test-GssDriveBackupPayload -SnapshotDirectory $activePath -Files $manifestPreparedFiles)
        }
        else {
            if ($null -eq $FinalInventory -or @($FinalInventory).Count -eq 0) {
                throw 'Finalization requires the current curated inventory so the post-apply state can be captured.'
            }

            $finalFiles = @(Copy-GssDriveBackupInventory -Inventory $FinalInventory -SnapshotDirectory $activePath -PayloadPrefix 'payload')
            [void](Test-GssDriveBackupPayload -SnapshotDirectory $activePath -Files $finalFiles)
            $expectedPrior = [string](Get-GssDriveBackupProperty $prepared @('prior_manifest_sha256'))

            $backupManifest = [ordered]@{
                schema_version = 1
                manifest_type = 'backup'
                status = 'Committed'
                run_id = $RunId
                fingerprint = $Fingerprint
                report_week = [string](Get-GssDriveBackupProperty $prepared @('report_week'))
                prepared_at_utc = [string](Get-GssDriveBackupProperty $prepared @('prepared_at_utc'))
                finalized_at_utc = [datetime]::UtcNow.ToString('o')
                host = [string](Get-GssDriveBackupProperty $prepared @('host'))
                release = [string](Get-GssDriveBackupProperty $prepared @('release'))
                prior_manifest_sha256 = if ([string]::IsNullOrWhiteSpace($expectedPrior)) { $null } else { $expectedPrior }
                prepared_manifest_sha256 = Get-GssDriveBackupSha256 -Path $preparedPath
                drive = Get-GssDriveBackupProperty $prepared @('drive')
                data_classification = Get-GssDriveBackupProperty $prepared @('data_classification')
                scope = Get-GssDriveBackupProperty $prepared @('scope')
                file_count = @($finalFiles).Count
                total_bytes = [long](($finalFiles | Measure-Object -Property byte_size -Sum).Sum)
                files = @($finalFiles | Sort-Object snapshot_path)
                prepared_file_count = @($preparedFiles).Count
                prepared_total_bytes = [long](($preparedFiles | Measure-Object -Property byte_size -Sum).Sum)
                prepared_files = @($preparedFiles | Sort-Object snapshot_path)
            }
            Write-GssDriveBackupStatus -Directory $activePath -Status PendingFinalize -RunId $RunId -Fingerprint $Fingerprint -Message 'Post-apply payload captured; waiting for snapshot promotion and commit receipt.'
            Write-GssDriveBackupAtomicJson -Path $manifestPath -Value $backupManifest
            $backupManifest = Read-GssDriveBackupJson -Path $manifestPath
        }

        $chainHead = Get-GssDriveBackupChainHead -RootPath $context.RootPath
        $expectedPrior = [string](Get-GssDriveBackupProperty $backupManifest @('prior_manifest_sha256'))
        $actualPrior = if ($chainHead) { $chainHead.ManifestSha256 } else { '' }
        if ($expectedPrior -ne $actualPrior) {
            Write-GssDriveBackupStatus -Directory $activePath -Status Blocked -RunId $RunId -Fingerprint $Fingerprint -Message 'Backup chain head changed after preparation; retry requires manual review.'
            throw 'Backup chain head changed after preparation; refusing to finalize an ambiguous manifest chain.'
        }

        if ($location.IsPartial) {
            $reportWeekText = [string](Get-GssDriveBackupProperty $prepared @('report_week'))
            $reportWeek = [datetime]::ParseExact($reportWeekText, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
            $snapshotParent = Join-Path $context.RootPath ("snapshots\{0}\{1}" -f $reportWeek.ToString('yyyy'), $reportWeek.ToString('yyyy-MM'))
            if (-not (Test-Path -LiteralPath $snapshotParent -PathType Container)) {
                New-Item -ItemType Directory -Path $snapshotParent -Force | Out-Null
            }
            $finalPath = Join-Path $snapshotParent $RunId
            if (Test-Path -LiteralPath $finalPath) {
                throw "Final snapshot path already exists and will not be overwritten: $finalPath"
            }
            Move-Item -LiteralPath $activePath -Destination $finalPath
            $activePath = $finalPath
            $manifestPath = Join-Path $activePath 'backup-manifest.json'
            $preparedPath = Join-Path $activePath 'prepared-manifest.json'
        }

        $manifestHash = Get-GssDriveBackupSha256 -Path $manifestPath
        $manifest = Read-GssDriveBackupJson -Path $manifestPath
        $receiptPath = Join-Path $activePath 'commit-receipt.json'
        $receipt = [ordered]@{
            schema_version = 1
            status = 'Committed'
            run_id = $RunId
            fingerprint = $Fingerprint
            committed_at_utc = [datetime]::UtcNow.ToString('o')
            backup_manifest_sha256 = $manifestHash
            prepared_manifest_sha256 = Get-GssDriveBackupSha256 -Path $preparedPath
            file_count = [int](Get-GssDriveBackupProperty $manifest @('file_count'))
            total_bytes = [long](Get-GssDriveBackupProperty $manifest @('total_bytes'))
            verification_level = $context.Settings.VerificationLevel
        }
        Write-GssDriveBackupAtomicJson -Path $receiptPath -Value $receipt

        Set-GssDriveBackupChainHead `
            -RootPath $context.RootPath `
            -RunId $RunId `
            -Fingerprint $Fingerprint `
            -CommittedAtUtc ([string]$receipt.committed_at_utc) `
            -ManifestPath $manifestPath `
            -ManifestSha256 $manifestHash
        try {
            Write-GssDriveBackupStatus -Directory $activePath -Status Committed -RunId $RunId -Fingerprint $Fingerprint -Message 'Commit receipt and chain head verified.'
        }
        catch {
            # The immutable receipt remains the authority if the convenience status file cannot be refreshed.
            Write-Verbose "Commit receipt is authoritative; convenience status refresh failed: $($_.Exception.Message)"
        }

        $validated = Test-GssCommittedBackupSnapshot -SnapshotPath $activePath
        return [pscustomobject]@{
            Status = 'Committed'
            BackupStatus = 'Committed'
            RunId = $RunId
            Fingerprint = $Fingerprint
            SnapshotPath = $activePath
            BackupManifestPath = $validated.ManifestPath
            BackupManifestSha256 = $validated.ManifestSha256
            CommitReceiptPath = $validated.ReceiptPath
            FileCount = $validated.ValidatedFileCount
            TotalBytes = [long](Get-GssDriveBackupProperty $validated.Manifest @('total_bytes'))
            VerificationLevel = $context.Settings.VerificationLevel
            Idempotent = $false
        }
    }
    catch {
        if ($activePath -and (Test-Path -LiteralPath $activePath -PathType Container)) {
            try {
                $existingStatus = $null
                $statusPath = Join-Path $activePath 'backup-status.json'
                if (Test-Path -LiteralPath $statusPath -PathType Leaf) {
                    $existingStatus = [string](Get-GssDriveBackupProperty (Read-GssDriveBackupJson $statusPath) @('status'))
                }
                if ($existingStatus -ne 'Blocked') {
                    Write-GssDriveBackupStatus -Directory $activePath -Status PendingFinalize -RunId $RunId -Fingerprint $Fingerprint -Message $_.Exception.Message
                }
            }
            catch { Write-Verbose "Could not record finalization failure status: $($_.Exception.Message)" }
        }
        throw
    }
    finally {
        Exit-GssDriveBackupMutex -Mutex $mutex
    }
}

function Get-GssDriveBackupIsoWeekKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [datetime]$Date
    )

    $dayOffset = (([int]$Date.DayOfWeek + 6) % 7)
    $thursday = $Date.Date.AddDays(3 - $dayOffset)
    $week = [Globalization.CultureInfo]::InvariantCulture.Calendar.GetWeekOfYear(
        $Date.Date,
        [Globalization.CalendarWeekRule]::FirstFourDayWeek,
        [DayOfWeek]::Monday
    )
    return ('{0}-W{1:d2}' -f $thursday.Year, $week)
}

function Get-GssCommittedBackupSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RootPath
    )

    $snapshotsRoot = Join-Path $RootPath 'snapshots'
    if (-not (Test-Path -LiteralPath $snapshotsRoot -PathType Container)) {
        return @()
    }

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($receiptFile in Get-ChildItem -LiteralPath $snapshotsRoot -Filter 'commit-receipt.json' -File -Recurse) {
        $snapshotPath = Split-Path -Parent $receiptFile.FullName
        $receipt = Read-GssDriveBackupJson -Path $receiptFile.FullName
        $manifestPath = Join-Path $snapshotPath 'backup-manifest.json'
        $manifestHash = Get-GssDriveBackupSha256 -Path $manifestPath
        if ($manifestHash -ne [string](Get-GssDriveBackupProperty $receipt @('backup_manifest_sha256'))) {
            throw "Retention inventory found a committed snapshot with an invalid manifest receipt: $snapshotPath"
        }
        $manifest = Read-GssDriveBackupJson -Path $manifestPath
        $reportWeek = [datetime]::ParseExact([string](Get-GssDriveBackupProperty $manifest @('report_week')), 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $results.Add([pscustomobject]@{
            RunId = [string](Get-GssDriveBackupProperty $manifest @('run_id'))
            Fingerprint = [string](Get-GssDriveBackupProperty $manifest @('fingerprint'))
            ReportWeek = $reportWeek.Date
            IsoWeek = Get-GssDriveBackupIsoWeekKey $reportWeek
            Month = $reportWeek.ToString('yyyy-MM')
            CommittedAtUtc = [datetime](Get-GssDriveBackupProperty $receipt @('committed_at_utc'))
            SnapshotPath = $snapshotPath
            ManifestPath = $manifestPath
            ManifestSha256 = $manifestHash
            TotalBytes = [long](Get-GssDriveBackupProperty $manifest @('total_bytes'))
        })
    }
    return ($results | ForEach-Object { $_ })
}

function Write-GssDriveBackupRetentionReport {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    param(
        [datetime]$AsOfDate = (Get-Date),
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath),
        [string]$ReportPath
    )

    $context = Get-GssDriveBackupRootContext -SettingsPath $SettingsPath
    $snapshots = @(Get-GssCommittedBackupSnapshot -RootPath $context.RootPath)
    $weeklyKeys = @{}
    for ($index = 0; $index -lt $context.Settings.RetentionWeekly; $index++) {
        $weeklyKeys[(Get-GssDriveBackupIsoWeekKey $AsOfDate.Date.AddDays(-7 * $index))] = $true
    }
    $monthlyKeys = @{}
    for ($index = 1; $index -le $context.Settings.RetentionMonthly; $index++) {
        $monthlyKeys[$AsOfDate.Date.AddMonths(-$index).ToString('yyyy-MM')] = $true
    }

    $keep = @{}
    foreach ($group in @($snapshots | Where-Object { $weeklyKeys.ContainsKey($_.IsoWeek) } | Group-Object IsoWeek)) {
        $winner = $group.Group | Sort-Object CommittedAtUtc -Descending | Select-Object -First 1
        $keep[$winner.RunId] = @($keep[$winner.RunId]) + "newest_in_$($group.Name)"
    }
    foreach ($group in @($snapshots | Where-Object { $monthlyKeys.ContainsKey($_.Month) } | Group-Object Month)) {
        $winner = $group.Group | Sort-Object CommittedAtUtc -Descending | Select-Object -First 1
        $keep[$winner.RunId] = @($keep[$winner.RunId]) + "newest_in_$($group.Name)"
    }

    $kept = @()
    $candidates = @()
    foreach ($snapshot in $snapshots | Sort-Object CommittedAtUtc -Descending) {
        $entry = [ordered]@{
            run_id = $snapshot.RunId
            report_week = $snapshot.ReportWeek.ToString('yyyy-MM-dd')
            committed_at_utc = $snapshot.CommittedAtUtc.ToUniversalTime().ToString('o')
            snapshot_relative_path = Get-GssDriveBackupRelativePath -Path $snapshot.SnapshotPath -Root $context.RootPath
            total_bytes = $snapshot.TotalBytes
            reason = if ($keep.ContainsKey($snapshot.RunId)) { @($keep[$snapshot.RunId]) } else { @('outside_retention_union_or_superseded') }
        }
        if ($keep.ContainsKey($snapshot.RunId)) { $kept += $entry } else { $candidates += $entry }
    }

    $report = [ordered]@{
        schema_version = 1
        report_type = 'retention_candidates_only'
        generated_at_utc = [datetime]::UtcNow.ToString('o')
        as_of_date = $AsOfDate.Date.ToString('yyyy-MM-dd')
        policy = [ordered]@{
            weekly = $context.Settings.RetentionWeekly
            monthly = $context.Settings.RetentionMonthly
            action = 'report_only'
            automatic_deletion = $false
        }
        completed_snapshot_count = $snapshots.Count
        kept = $kept
        deletion_candidates = $candidates
    }

    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $reportFolder = Join-Path $context.RootPath 'retention-reports'
        $ReportPath = Join-Path $reportFolder ("retention-candidates-{0}.json" -f [datetime]::UtcNow.ToString('yyyyMMddTHHmmssZ'))
    }
    if ($PSCmdlet.ShouldProcess($ReportPath, 'Write report-only retention candidate inventory without deleting snapshots')) {
        Write-GssDriveBackupAtomicJson -Path $ReportPath -Value $report
    }
    return [pscustomobject]@{
        Status = 'Reported'
        ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
        CompletedSnapshotCount = $snapshots.Count
        KeptCount = $kept.Count
        CandidateCount = $candidates.Count
        AutomaticDeletion = $false
        Report = $report
    }
}

function Restore-GssDriveBackupForVerification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RunId,
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath),
        [string]$LocalAppDataPath = $env:LOCALAPPDATA,
        [ValidateSet('Final', 'Prepared')]
        [string]$Phase = 'Final'
    )

    if ([string]::IsNullOrWhiteSpace($LocalAppDataPath)) {
        throw 'LOCALAPPDATA is unavailable; verify-only restore cannot choose a safe destination.'
    }
    $context = Get-GssDriveBackupRootContext -SettingsPath $SettingsPath
    $location = Find-GssDriveBackupSnapshot -RootPath $context.RootPath -RunId $RunId
    if ($null -eq $location -or $location.IsPartial) {
        throw "A committed Drive snapshot was not found for verify-only restore: $RunId"
    }
    $validated = Test-GssCommittedBackupSnapshot -SnapshotPath $location.Path
    if ($Phase -eq 'Prepared') {
        $restoreManifest = $validated.PreparedManifest
        $sourceManifestHash = $validated.PreparedManifestSha256
    }
    else {
        $restoreManifest = $validated.Manifest
        $sourceManifestHash = $validated.ManifestSha256
    }

    # Keep the verification root compact for Windows PowerShell 5.1 MAX_PATH
    # compatibility. The status reader also recognizes the former long path.
    $verificationRoot = [System.IO.Path]::GetFullPath((Join-Path $LocalAppDataPath 'GSSSurveyWorkbookAutomation\rv')).TrimEnd('\', '/')
    $localRoot = [System.IO.Path]::GetFullPath($LocalAppDataPath).TrimEnd('\', '/')
    if (-not $verificationRoot.StartsWith("$localRoot\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Verify-only restore destination escaped LOCALAPPDATA.'
    }
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $runIdDigestBytes = $sha256.ComputeHash([Text.Encoding]::UTF8.GetBytes($RunId))
    }
    finally {
        $sha256.Dispose()
    }
    $runIdDigest = ([BitConverter]::ToString($runIdDigestBytes)).Replace('-', '').ToLowerInvariant().Substring(0, 8)
    $phaseCode = if ($Phase -eq 'Prepared') { 'p' } else { 'f' }
    $destination = Join-Path $verificationRoot ("{0}-{1}-{2}-{3}" -f [datetime]::UtcNow.ToString('yyyyMMddHHmmss'), $runIdDigest, $phaseCode, [guid]::NewGuid().ToString('N').Substring(0, 6))
    if (Test-Path -LiteralPath $destination) {
        throw "Verify-only restore destination already exists and will not be overwritten: $destination"
    }
    New-Item -ItemType Directory -Path $destination -Force | Out-Null

    $restored = New-Object System.Collections.Generic.List[object]
    foreach ($file in @(Get-GssDriveBackupProperty $restoreManifest @('files') @())) {
        $snapshotRelative = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $file @('snapshot_path')))
        $source = Join-Path $location.Path $snapshotRelative.Replace('/', '\')
        $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $file @('portable_path')))
        $target = Join-Path $destination $portable.Replace('/', '\')
        $targetParent = Split-Path -Parent $target
        if (-not (Test-Path -LiteralPath $targetParent -PathType Container)) {
            [void][System.IO.Directory]::CreateDirectory($targetParent)
        }
        if (-not (Test-Path -LiteralPath $targetParent -PathType Container)) {
            throw "Verify-only restore could not create target directory: $targetParent"
        }
        [System.IO.File]::Copy($source, $target, $false)
        $hash = Get-GssDriveBackupSha256 -Path $target
        if ($hash -ne [string](Get-GssDriveBackupProperty $file @('sha256'))) {
            throw "Verify-only restore hash mismatch: $portable"
        }
        $restored.Add([pscustomobject][ordered]@{
            portable_path = $portable
            byte_size = [long](Get-GssDriveBackupProperty $file @('byte_size'))
            sha256 = $hash
        })
    }

    $receiptPath = Join-Path $destination 'restore-verification.json'
    $receipt = [ordered]@{
        schema_version = 1
        operation = 'verify_only_restore'
        status = 'Verified'
        run_id = $RunId
        phase = $Phase
        verified_at_utc = [datetime]::UtcNow.ToString('o')
        source_snapshot_manifest_sha256 = $sourceManifestHash
        destination = $destination
        live_workbook_overwritten = $false
        verification_level = $context.Settings.VerificationLevel
        file_count = $restored.Count
        files = @($restored | ForEach-Object { $_ })
    }
    Write-GssDriveBackupAtomicJson -Path $receiptPath -Value $receipt

    return [pscustomobject]@{
        Status = 'Verified'
        RunId = $RunId
        Phase = $Phase
        Destination = $destination
        ReceiptPath = $receiptPath
        FileCount = $restored.Count
        LiveWorkbookOverwritten = $false
        VerificationLevel = $context.Settings.VerificationLevel
    }
}

function Get-GssDriveBackupRestoreDrillStatus {
    [CmdletBinding()]
    param(
        [string]$LocalAppDataPath = $env:LOCALAPPDATA,
        [datetime]$AsOfDate = (Get-Date)
    )

    if ([string]::IsNullOrWhiteSpace($LocalAppDataPath)) {
        throw 'LOCALAPPDATA is unavailable; restore-drill status cannot be resolved.'
    }
    $roots = @(
        (Join-Path $LocalAppDataPath 'GSSSurveyWorkbookAutomation\rv'),
        (Join-Path $LocalAppDataPath 'GSSSurveyWorkbookAutomation\restore-verification')
    ) | Select-Object -Unique
    $hashOnlyReceipts = @()
    $quarterlyReceipts = @()
    foreach ($root in $roots) {
        $fullRoot = [System.IO.Path]::GetFullPath($root).TrimEnd('\', '/')
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {
            continue
        }
        foreach ($file in Get-ChildItem -LiteralPath $root -Filter 'restore-verification.json' -File -Recurse) {
            try {
                $receipt = Read-GssDriveBackupJson -Path $file.FullName
                if ([string](Get-GssDriveBackupProperty $receipt @('operation')) -eq 'verify_only_restore' -and
                    [string](Get-GssDriveBackupProperty $receipt @('status')) -eq 'Verified' -and
                    -not [bool](Get-GssDriveBackupProperty $receipt @('live_workbook_overwritten') $true)) {
                    $hashOnlyReceipts += [pscustomobject]@{
                        Path = $file.FullName
                        VerifiedAtUtc = [datetime](Get-GssDriveBackupProperty $receipt @('verified_at_utc'))
                        RunId = [string](Get-GssDriveBackupProperty $receipt @('run_id'))
                    }
                }
            }
            catch {
                # An invalid receipt is not evidence of a completed drill.
                Write-Verbose "Ignored invalid restore verification receipt '$($file.FullName)': $($_.Exception.Message)"
            }
        }
        foreach ($file in Get-ChildItem -LiteralPath $root -Filter 'quarterly-restore-drill.json' -File -Recurse) {
            try {
                $receipt = Read-GssDriveBackupJson -Path $file.FullName
                if ([string](Get-GssDriveBackupProperty $receipt @('operation')) -ne 'quarterly_verify_only_restore_drill' -or
                    [string](Get-GssDriveBackupProperty $receipt @('status')) -ne 'Passed' -or
                    [bool](Get-GssDriveBackupProperty $receipt @('live_workbook_overwritten') $true)) {
                    continue
                }

                $excelReceiptPath = [System.IO.Path]::GetFullPath([string](Get-GssDriveBackupProperty $receipt @('excel_validation_receipt')))
                $driveReceiptPath = [System.IO.Path]::GetFullPath([string](Get-GssDriveBackupProperty $receipt @('drive_restore_receipt')))
                if (-not $excelReceiptPath.StartsWith("$fullRoot\", [System.StringComparison]::OrdinalIgnoreCase) -or
                    -not $driveReceiptPath.StartsWith("$fullRoot\", [System.StringComparison]::OrdinalIgnoreCase) -or
                    -not (Test-Path -LiteralPath $excelReceiptPath -PathType Leaf) -or
                    -not (Test-Path -LiteralPath $driveReceiptPath -PathType Leaf)) {
                    continue
                }
                $excelReceipt = Read-GssDriveBackupJson -Path $excelReceiptPath
                $driveReceipt = Read-GssDriveBackupJson -Path $driveReceiptPath
                if ([string](Get-GssDriveBackupProperty $excelReceipt @('Status', 'status')) -ne 'Passed' -or
                    [string](Get-GssDriveBackupProperty $driveReceipt @('operation')) -ne 'verify_only_restore' -or
                    [string](Get-GssDriveBackupProperty $driveReceipt @('status')) -ne 'Verified' -or
                    [bool](Get-GssDriveBackupProperty $driveReceipt @('live_workbook_overwritten') $true)) {
                    continue
                }
                $quarterlyReceipts += [pscustomobject]@{
                    Path = $file.FullName
                    VerifiedAtUtc = [datetime](Get-GssDriveBackupProperty $receipt @('completed_at_utc'))
                    RunId = [string](Get-GssDriveBackupProperty $receipt @('run_id'))
                    ExcelReceiptPath = $excelReceiptPath
                    DriveReceiptPath = $driveReceiptPath
                }
            }
            catch {
                Write-Verbose "Ignored invalid quarterly restore-drill receipt '$($file.FullName)': $($_.Exception.Message)"
            }
        }
    }

    $latestHashOnly = $hashOnlyReceipts | Sort-Object VerifiedAtUtc -Descending | Select-Object -First 1
    $latest = $quarterlyReceipts | Sort-Object VerifiedAtUtc -Descending | Select-Object -First 1
    $nextDue = if ($latest) { $latest.VerifiedAtUtc.ToLocalTime().Date.AddMonths(3) } else { $AsOfDate.Date }
    return [pscustomobject]@{
        Status = if ($latest -and $AsOfDate.Date -lt $nextDue) { 'Current' } else { 'Due' }
        Frequency = 'Quarterly'
        LastVerifiedAtUtc = if ($latest) { $latest.VerifiedAtUtc.ToUniversalTime().ToString('o') } else { $null }
        LastRunId = if ($latest) { $latest.RunId } else { $null }
        LastReceiptPath = if ($latest) { $latest.Path } else { $null }
        LastExcelValidationReceiptPath = if ($latest) { $latest.ExcelReceiptPath } else { $null }
        HashOnlyVerificationCount = $hashOnlyReceipts.Count
        LatestHashOnlyVerifiedAtUtc = if ($latestHashOnly) { $latestHashOnly.VerifiedAtUtc.ToUniversalTime().ToString('o') } else { $null }
        HashOnlyVerificationSatisfiesQuarterlyDrill = $false
        NextDueDate = $nextDue.ToString('yyyy-MM-dd')
        LiveWorkbookOverwriteAllowed = $false
    }
}

function Get-GssDriveBackupCapacityProjection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [long]$ProjectedWeeklyBytes,
        [Nullable[long]]$FreeBytes,
        [string]$SettingsPath = (Get-GssDriveBackupDefaultSettingsPath)
    )

    if ($ProjectedWeeklyBytes -le 0) {
        throw 'Projected weekly backup bytes must be greater than zero.'
    }
    $context = Get-GssDriveBackupRootContext -SettingsPath $SettingsPath
    if ($null -eq $FreeBytes) {
        $driveRoot = [System.IO.Path]::GetPathRoot($context.RootPath)
        $driveInfo = New-Object System.IO.DriveInfo($driveRoot)
        if (-not $driveInfo.IsReady) {
            throw "DriveFS volume is not ready for capacity projection: $driveRoot"
        }
        $availableFreeBytes = [long]$driveInfo.AvailableFreeSpace
    }
    else {
        $availableFreeBytes = [long]$FreeBytes
    }

    $projectedSnapshotBytes = [long]($ProjectedWeeklyBytes * 2)
    $retainedLoads = $context.Settings.RetentionWeekly + $context.Settings.RetentionMonthly
    $retainedProjection = [long]($projectedSnapshotBytes * $retainedLoads)
    $worstCaseWeeklyLoadsAvailable = [long][math]::Floor($availableFreeBytes / [double]$projectedSnapshotBytes)
    return [pscustomobject]@{
        Status = 'Projected'
        ProjectedWeeklyBytes = $ProjectedWeeklyBytes
        ProjectedPreparedAndFinalSnapshotBytes = $projectedSnapshotBytes
        AvailableFreeBytes = $availableFreeBytes
        RetentionWeekly = $context.Settings.RetentionWeekly
        RetentionMonthly = $context.Settings.RetentionMonthly
        WorstCaseRetainedLoads = $retainedLoads
        ProjectedRetainedBytes = $retainedProjection
        WorstCaseWeeklyLoadsAvailable = $worstCaseWeeklyLoadsAvailable
        StructuralRedesignThresholdLoads = 104
        StructuralRedesignStatus = if ($worstCaseWeeklyLoadsAvailable -lt 104) { 'ReviewRequired' } else { 'Deferred' }
        VerificationLevel = $context.Settings.VerificationLevel
    }
}
