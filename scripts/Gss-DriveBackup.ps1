$script:GssDriveBackupClassificationLabel = 'CONTAINS PERSONAL DATA ' + [char]0x2014 + ' RESTRICTED'
$script:GssDriveBackupLegacySafePathLength = 248
$script:GssDriveBackupMaxRetainedExtensionLength = 32

function Get-GssDriveBackupCompactRelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PortablePath,
        [Parameter(Mandatory)]
        [string]$Prefix,
        [switch]$OmitExtension
    )

    $portable = Assert-GssDriveBackupSafeRelativePath -Path $PortablePath
    $safePrefix = Assert-GssDriveBackupSafeRelativePath -Path $Prefix
    $portableBytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($portable)
    $portableDigest = Get-GssDriveBackupByteSha256 -Bytes $portableBytes
    $extension = [System.IO.Path]::GetExtension($portable)
    if ($OmitExtension) {
        $extension = ''
    }
    elseif ($extension.Length -gt $script:GssDriveBackupMaxRetainedExtensionLength) {
        $extension = $extension.Substring(0, $script:GssDriveBackupMaxRetainedExtensionLength)
    }
    if ($extension -eq '.') {
        $extension = ''
    }
    return (Assert-GssDriveBackupSafeRelativePath -Path "$safePrefix/long-path/$portableDigest$extension")
}

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

    $json = $Value | ConvertTo-Json -Depth $Depth
    $encoding = New-Object System.Text.UTF8Encoding($false)
    $content = $json + [Environment]::NewLine
    $contentBytes = $encoding.GetBytes($content)
    if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
        $existingBytes = [System.IO.File]::ReadAllBytes($fullPath)
        if ($contentBytes.Length -eq $existingBytes.Length -and
            (Get-GssDriveBackupByteSha256 -Bytes $contentBytes) -ceq
                (Get-GssDriveBackupByteSha256 -Bytes $existingBytes)) {
            return
        }
    }

    # Keep atomic helper names short. Verify-only restore paths can approach the
    # legacy Windows MAX_PATH limit, and appending the full destination filename
    # plus a GUID makes an otherwise valid restore intermittently fail.
    $temporaryPath = Join-Path $parent ('.t-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [System.IO.File]::WriteAllText($temporaryPath, $content, $encoding)
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

function Test-GssDriveBackupNoLinkTraversal {
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
    $cursor = $fullPath
    while (-not [string]::IsNullOrWhiteSpace($cursor) -and
        ($cursor.Equals($fullRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
            $cursor.StartsWith("$fullRoot\", [System.StringComparison]::OrdinalIgnoreCase))) {
        $item = Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        $linkTypeProperty = $item.PSObject.Properties['LinkType']
        $targetProperty = $item.PSObject.Properties['Target']
        $linkType = if ($linkTypeProperty) { [string]$linkTypeProperty.Value } else { '' }
        $targets = if ($targetProperty -and $null -ne $targetProperty.Value) { @($targetProperty.Value) } else { @() }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
            (-not [string]::IsNullOrWhiteSpace($linkType) -or $targets.Count -gt 0)) {
            throw "Narrow backup source paths cannot traverse a symbolic link or junction: $cursor"
        }
        if ($cursor.Equals($fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        $cursor = Split-Path -Parent $cursor
    }
    return $true
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
        [string[]]$ReleaseArchivePaths = @(),
        [ValidateSet('Full', 'RecoveryOnly', 'ReleaseOnly')]
        [string]$InventoryMode = 'Full'
    )

    $root = [System.IO.Path]::GetFullPath($GssRoot).TrimEnd('\', '/')
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        throw "GSS source root is unavailable: $root"
    }

    if ($InventoryMode -eq 'RecoveryOnly') {
        if (@($TransactionArtifactPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0 -or
            @($ReleaseArchivePaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
            throw 'RecoveryOnly inventories require closed, explicit BackupInventory records and cannot use transaction or release path shortcuts.'
        }
        $allowedRoles = @(
            'recovered_historical_detail',
            'recovery_ledger',
            'recovery_manifest',
            'recovery_receipt',
            'recovery_qa',
            'recovery_run_summary'
        )
        $recoveryTransactionHash = ''
        foreach ($item in @($AdditionalItems)) {
            if ($null -eq $item -or $item -is [string]) {
                throw 'RecoveryOnly BackupInventory entries must be structured records.'
            }
            $source = [string](Get-GssDriveBackupProperty $item @('SourcePath', 'source_path', 'Path', 'path'))
            $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $item @('PortablePath', 'portable_path')))
            $role = [string](Get-GssDriveBackupProperty $item @('Role', 'role'))
            $classification = [string](Get-GssDriveBackupProperty $item @('Classification', 'classification'))
            if ($role -notin $allowedRoles) {
                throw "RecoveryOnly BackupInventory role is not allowed: $role"
            }
            if (-not $portable.StartsWith('recovery/', [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "RecoveryOnly portable paths must stay under recovery/: $portable"
            }

            [void](Test-GssDriveBackupNoLinkTraversal -Path $source -Root $root)
            $relativeSource = (Get-GssDriveBackupRelativePath -Path $source -Root $root).Replace('/', '\')
            $historicalArchivePrefix = '03 Uploaded Survey Workbooks\Archive - Previous Uploads\Recovered Historical Detail\'
            $ledgerRelativePath = '_automation_runs\state\gss_feedback_first_seen.json'
            switch ($role) {
                'recovered_historical_detail' {
                    if (-not $relativeSource.StartsWith($historicalArchivePrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
                        [System.IO.Path]::GetExtension($relativeSource) -ne '.xlsx' -or
                        $classification -ne 'restricted_personal_data') {
                        throw 'Recovered historical detail must be an XLSX under the recovered archive and classified restricted_personal_data.'
                    }
                    $expectedPortable = 'recovery/recovered-detail/' + [System.IO.Path]::GetFileName($relativeSource)
                    if (-not $portable.Equals($expectedPortable, [System.StringComparison]::OrdinalIgnoreCase)) {
                        throw "Recovered historical detail must use its exact closed portable path: $expectedPortable"
                    }
                }
                'recovery_ledger' {
                    if (-not $relativeSource.Equals($ledgerRelativePath, [System.StringComparison]::OrdinalIgnoreCase) -or
                        $classification -ne 'restricted_operational') {
                        throw 'Recovery ledger must be the exact first-seen ledger and classified restricted_operational.'
                    }
                    if (-not $portable.Equals('recovery/state/gss_feedback_first_seen.json', [System.StringComparison]::OrdinalIgnoreCase)) {
                        throw 'Recovery ledger must use the exact recovery/state/gss_feedback_first_seen.json portable path.'
                    }
                }
                default {
                    if ($relativeSource -notmatch '^_automation_runs\\historical-recovery\\([a-f0-9]{64})\\([^\\]+)$' -or
                        $classification -ne 'restricted_operational') {
                        throw "Recovery evidence role '$role' must be a direct file in a manifest-hash transaction directory and be classified restricted_operational."
                    }
                    $transactionHash = $matches[1].ToLowerInvariant()
                    $leaf = $matches[2]
                    if ([string]::IsNullOrWhiteSpace($recoveryTransactionHash)) {
                        $recoveryTransactionHash = $transactionHash
                    }
                    elseif ($recoveryTransactionHash -cne $transactionHash) {
                        throw 'Recovery evidence files must all come from the same manifest-hash transaction directory.'
                    }
                    $expectedLeaf = switch ($role) {
                        'recovery_manifest' { 'recovery-manifest.json' }
                        'recovery_receipt' { 'transaction-receipt.json' }
                        'recovery_qa' { 'recovery-qa.json' }
                        'recovery_run_summary' { 'drive-recovery-summary.json' }
                        default { throw "Recovery evidence role is not bound to an approved filename: $role" }
                    }
                    if (-not $leaf.Equals($expectedLeaf, [System.StringComparison]::OrdinalIgnoreCase)) {
                        throw "Recovery evidence role '$role' must use the exact filename '$expectedLeaf'."
                    }
                    $expectedPortable = "recovery/evidence/$expectedLeaf"
                    if (-not $portable.Equals($expectedPortable, [System.StringComparison]::OrdinalIgnoreCase)) {
                        throw "Recovery evidence role '$role' must use the exact closed portable path: $expectedPortable"
                    }
                }
            }
        }
        if ([string]::IsNullOrWhiteSpace($recoveryTransactionHash)) {
            throw 'RecoveryOnly inventories must include at least one manifest-bound transaction evidence file.'
        }
    }

    if ($InventoryMode -eq 'ReleaseOnly') {
        if (@($TransactionArtifactPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0 -or
            @($ReleaseArchivePaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
            throw 'ReleaseOnly inventories require a closed, explicit BackupInventory and cannot use transaction or release path shortcuts.'
        }
        if (@($AdditionalItems).Count -ne 3) {
            throw 'ReleaseOnly inventories require exactly three structured release artifacts.'
        }

        $allowedRoles = @('release_archive', 'release_manifest', 'release_excel_receipt')
        foreach ($item in @($AdditionalItems)) {
            if ($null -eq $item -or $item -is [string]) {
                throw 'ReleaseOnly BackupInventory entries must be structured records.'
            }
            $source = [string](Get-GssDriveBackupProperty $item @('SourcePath', 'source_path', 'Path', 'path'))
            $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $item @('PortablePath', 'portable_path')))
            $role = [string](Get-GssDriveBackupProperty $item @('Role', 'role'))
            $classification = [string](Get-GssDriveBackupProperty $item @('Classification', 'classification'))
            if ($role -notin $allowedRoles) {
                throw "ReleaseOnly BackupInventory role is not allowed: $role"
            }
            if ($classification -cne 'restricted_operational') {
                throw "ReleaseOnly artifact '$role' must be classified restricted_operational."
            }

            [void](Test-GssDriveBackupNoLinkTraversal -Path $source -Root $root)
            $relativeSource = (Get-GssDriveBackupRelativePath -Path $source -Root $root).Replace('/', '\')
            switch ($role) {
                'release_archive' {
                    if ($relativeSource -notmatch '^_automation_runs\\state\\release\\(gss-survey-workbook-automation-v\d+\.\d+\.\d+\.zip)$') {
                        throw 'ReleaseOnly archive must use the exact versioned path under _automation_runs\state\release.'
                    }
                    $expectedPortable = "release/$($matches[1])"
                }
                'release_manifest' {
                    if (-not $relativeSource.Equals('GSS Survey Workbook Automation\release\release-manifest.json', [System.StringComparison]::OrdinalIgnoreCase)) {
                        throw 'ReleaseOnly manifest must be the exact program release\release-manifest.json file.'
                    }
                    $expectedPortable = 'release/release-manifest.json'
                }
                'release_excel_receipt' {
                    if (-not $relativeSource.Equals('_automation_runs\state\release\local-excel-validation-receipt.json', [System.StringComparison]::OrdinalIgnoreCase)) {
                        throw 'ReleaseOnly Excel receipt must use the exact _automation_runs\state\release path.'
                    }
                    $expectedPortable = 'release/local-excel-validation-receipt.json'
                }
            }
            if (-not $portable.Equals($expectedPortable, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "ReleaseOnly artifact '$role' must use the exact portable path '$expectedPortable'."
            }
        }
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

    if ($InventoryMode -eq 'Full') {
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

    $result = @($records | Sort-Object PortablePath)
    if ($InventoryMode -eq 'RecoveryOnly') {
        [void](Assert-GssRecoveryOnlyInventoryContract -GssRoot $root -Inventory $result)
    }
    elseif ($InventoryMode -eq 'ReleaseOnly') {
        [void](Assert-GssReleaseOnlyInventoryContract -GssRoot $root -Inventory $result)
    }
    return $result
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
        [string]$PayloadPrefix,
        [string[]]$PathBudgetDirectories = @()
    )

    $safePayloadPrefix = Assert-GssDriveBackupSafeRelativePath -Path $PayloadPrefix
    $results = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($item in @($Inventory | Sort-Object PortablePath)) {
        $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string]$item.PortablePath)
        $snapshotRelative = Assert-GssDriveBackupSafeRelativePath -Path "$safePayloadPrefix/$portable"
        $budgetDirectories = @($SnapshotDirectory) + @($PathBudgetDirectories)
        $usesReservedCompactNamespace = $portable.Equals('r', [System.StringComparison]::OrdinalIgnoreCase) -or
            $portable.Equals('long-path', [System.StringComparison]::OrdinalIgnoreCase) -or
            $portable.StartsWith('long-path/', [System.StringComparison]::OrdinalIgnoreCase)
        $requiresCompaction = $usesReservedCompactNamespace -or @($budgetDirectories | Where-Object {
            (Join-Path $_ $snapshotRelative.Replace('/', '\')).Length -ge $script:GssDriveBackupLegacySafePathLength
        }).Count -gt 0
        if ($requiresCompaction) {
            $snapshotRelative = Get-GssDriveBackupCompactRelativePath -PortablePath $portable -Prefix $safePayloadPrefix
            $overBudgetCompactDestination = @($budgetDirectories | ForEach-Object {
                Join-Path $_ $snapshotRelative.Replace('/', '\')
            } | Where-Object {
                $_.Length -ge $script:GssDriveBackupLegacySafePathLength
            } | Select-Object -First 1)
            if ($overBudgetCompactDestination.Count -gt 0) {
                $snapshotRelative = Get-GssDriveBackupCompactRelativePath -PortablePath $portable -Prefix $safePayloadPrefix -OmitExtension
                $overBudgetCompactDestination = @($budgetDirectories | ForEach-Object {
                    Join-Path $_ $snapshotRelative.Replace('/', '\')
                } | Where-Object {
                    $_.Length -ge $script:GssDriveBackupLegacySafePathLength
                } | Select-Object -First 1)
            }
            if ($overBudgetCompactDestination.Count -gt 0) {
                throw "Compacted snapshot destination still exceeds the safe Windows path budget before copy. Shorten the Drive root or RunId: $($overBudgetCompactDestination[0])"
            }
        }
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

function ConvertTo-GssDriveBackupInventoryEvidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Inventory
    )

    return @($Inventory | ForEach-Object {
        $source = [System.IO.Path]::GetFullPath([string]$_.SourcePath)
        $sourceInfo = Get-Item -LiteralPath $source -Force -ErrorAction Stop
        [pscustomobject][ordered]@{
            portable_path = Assert-GssDriveBackupSafeRelativePath -Path ([string]$_.PortablePath)
            role = [string]$_.Role
            classification = [string]$_.Classification
            byte_size = [long]$sourceInfo.Length
            sha256 = Get-GssDriveBackupSha256 -Path $source
        }
    })
}

function Test-GssDriveBackupExactFileSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$ExpectedFiles,
        [Parameter(Mandatory)]
        [object[]]$ActualFiles,
        [string]$Label = 'RecoveryOnly inventory'
    )

    $expectedByPath = @{}
    foreach ($file in @($ExpectedFiles)) {
        $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $file @('portable_path', 'PortablePath')))
        $key = $portable.ToLowerInvariant()
        if ($expectedByPath.ContainsKey($key)) {
            throw "$Label contains a duplicate prepared path: $portable"
        }
        $expectedByPath[$key] = $file
    }
    if (@($ActualFiles).Count -ne $expectedByPath.Count) {
        throw "$Label count does not match the prepared inventory."
    }
    $seen = @{}
    foreach ($file in @($ActualFiles)) {
        $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $file @('portable_path', 'PortablePath')))
        $key = $portable.ToLowerInvariant()
        if ($seen.ContainsKey($key) -or -not $expectedByPath.ContainsKey($key)) {
            throw "$Label contains an unprepared or duplicate path: $portable"
        }
        $seen[$key] = $true
        $expected = $expectedByPath[$key]
        foreach ($field in @('role', 'classification', 'sha256', 'byte_size')) {
            $expectedValue = [string](Get-GssDriveBackupProperty $expected @($field))
            $actualValue = [string](Get-GssDriveBackupProperty $file @($field))
            if ($actualValue -cne $expectedValue) {
                throw "$Label changed $field after preparation: $portable"
            }
        }
    }
    return $true
}

function Assert-GssDriveBackupObjectShape {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Value,
        [Parameter(Mandatory)]
        [string[]]$Required,
        [Parameter(Mandatory)]
        [string[]]$Allowed,
        [Parameter(Mandatory)]
        [string]$Label
    )

    if ($null -eq $Value -or $Value -is [string]) {
        throw "$Label must be a structured JSON object."
    }
    $propertyNames = @($Value.PSObject.Properties.Name)
    foreach ($requiredName in $Required) {
        if ($propertyNames -notcontains $requiredName) {
            throw "$Label is missing required property '$requiredName'."
        }
    }
    foreach ($propertyName in $propertyNames) {
        if ($Allowed -notcontains $propertyName) {
            throw "$Label contains unsupported property '$propertyName'."
        }
    }
    return $true
}

function Test-GssDriveBackupExactInventoryRecordSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Expected,
        [Parameter(Mandatory)]
        [object[]]$Actual,
        [Parameter(Mandatory)]
        [string]$Label
    )

    $expectedByPortablePath = @{}
    foreach ($item in @($Expected)) {
        $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $item @('PortablePath', 'portable_path')))
        $key = $portable.ToLowerInvariant()
        if ($expectedByPortablePath.ContainsKey($key)) {
            throw "$Label expected inventory contains duplicate portable path '$portable'."
        }
        $expectedByPortablePath[$key] = $item
    }
    if (@($Actual).Count -ne $expectedByPortablePath.Count) {
        throw "$Label count does not match the requested closed inventory."
    }

    $seen = @{}
    foreach ($item in @($Actual)) {
        $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $item @('PortablePath', 'portable_path')))
        $key = $portable.ToLowerInvariant()
        if ($seen.ContainsKey($key) -or -not $expectedByPortablePath.ContainsKey($key)) {
            throw "$Label contains an unexpected or duplicate portable path '$portable'."
        }
        $seen[$key] = $true
        $expectedItem = $expectedByPortablePath[$key]
        $expectedSource = [System.IO.Path]::GetFullPath([string](Get-GssDriveBackupProperty $expectedItem @('SourcePath', 'source_path')))
        $actualSource = [System.IO.Path]::GetFullPath([string](Get-GssDriveBackupProperty $item @('SourcePath', 'source_path')))
        if (-not $actualSource.Equals($expectedSource, [System.StringComparison]::OrdinalIgnoreCase) -or
            [string](Get-GssDriveBackupProperty $item @('Role', 'role')) -cne [string](Get-GssDriveBackupProperty $expectedItem @('Role', 'role')) -or
            [string](Get-GssDriveBackupProperty $item @('Classification', 'classification')) -cne [string](Get-GssDriveBackupProperty $expectedItem @('Classification', 'classification'))) {
            throw "$Label changed the source, role, or classification for '$portable'."
        }
    }
    return $true
}

function Get-GssDriveBackupByteSha256 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha256.ComputeHash($Bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-GssDriveBackupZipEntryByteArray {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Entry
    )

    $stream = $null
    $memory = $null
    try {
        $stream = $Entry.Open()
        $memory = New-Object System.IO.MemoryStream
        $stream.CopyTo($memory)
        return [byte[]]$memory.ToArray()
    }
    finally {
        if ($memory) { $memory.Dispose() }
        if ($stream) { $stream.Dispose() }
    }
}

function Get-GssReleaseOnlyInventoryFingerprint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Inventory,
        [Parameter(Mandatory)]
        [string]$Release
    )

    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add('gss-release-only-inventory/v1')
    $parts.Add($Release.Trim())
    foreach ($evidence in @(ConvertTo-GssDriveBackupInventoryEvidence -Inventory $Inventory | Sort-Object portable_path)) {
        $parts.Add(([string]$evidence.portable_path).Trim())
        $parts.Add(([string]$evidence.role).Trim())
        $parts.Add(([string]$evidence.classification).Trim())
        $parts.Add(([string]$evidence.byte_size).Trim())
        $parts.Add(([string]$evidence.sha256).Trim().ToLowerInvariant())
    }
    $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes(($parts -join "`n"))
    return Get-GssDriveBackupByteSha256 -Bytes $bytes
}

function Get-GssDriveBackupPreparedRunFingerprint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Run
    )

    $parts = @(
        'gss-transaction-v1',
        ([string](Get-GssDriveBackupProperty $Run @('RunId'))).Trim().ToLowerInvariant(),
        ([string](Get-GssDriveBackupProperty $Run @('HostName'))).Trim().ToLowerInvariant(),
        ([string](Get-GssDriveBackupProperty $Run @('CurrentWeekEnding'))).Trim(),
        ([string](Get-GssDriveBackupProperty $Run @('StartingWorkbookSha256'))).Trim().ToLowerInvariant(),
        ([string](Get-GssDriveBackupProperty $Run @('CurrentSourceSha256'))).Trim().ToLowerInvariant(),
        ([string](Get-GssDriveBackupProperty $Run @('PriorYearSourceSha256'))).Trim().ToLowerInvariant(),
        ([string](Get-GssDriveBackupProperty $Run @('StagedWorkbookSha256'))).Trim().ToLowerInvariant(),
        ([string](Get-GssDriveBackupProperty $Run @('StagedPdfSha256'))).Trim().ToLowerInvariant(),
        ([string](Get-GssDriveBackupProperty $Run @('ProgramRelease'))).Trim().ToLowerInvariant()
    )
    $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes(($parts -join "`n"))
    return Get-GssDriveBackupByteSha256 -Bytes $bytes
}

function Assert-GssReleaseOnlyArtifactSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ArchivePath,
        [Parameter(Mandatory)]
        [string]$ManifestPath,
        [Parameter(Mandatory)]
        [string]$ExcelReceiptPath,
        [string]$GssRoot
    )

    $archive = [System.IO.Path]::GetFullPath($ArchivePath)
    $manifestSource = [System.IO.Path]::GetFullPath($ManifestPath)
    $receiptSource = [System.IO.Path]::GetFullPath($ExcelReceiptPath)
    foreach ($requiredPath in @($archive, $manifestSource, $receiptSource)) {
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            throw "ReleaseOnly required artifact is missing: $requiredPath"
        }
    }

    $manifest = Read-GssDriveBackupJson -Path $manifestSource
    $version = [string](Get-GssDriveBackupProperty $manifest @('release_version'))
    $tag = [string](Get-GssDriveBackupProperty $manifest @('release_tag'))
    $archiveName = [string](Get-GssDriveBackupProperty $manifest @('archive_name'))
    if ($version -notmatch '^\d+\.\d+\.\d+$' -or
        $tag -cne "v$version" -or
        $archiveName -cne "gss-survey-workbook-automation-$tag.zip" -or
        [System.IO.Path]::GetFileName($archive) -cne $archiveName -or
        [string](Get-GssDriveBackupProperty $manifest @('commit_binding')) -cne 'exact_release_tag' -or
        [string](Get-GssDriveBackupProperty $manifest @('classification')) -cne 'PROGRAM SOURCE ONLY - NO GSS WORKBOOKS, REPORTS, OR CUSTOMER DATA') {
        throw 'ReleaseOnly manifest version, tag, archive name, commit binding, or classification is invalid.'
    }
    $runtime = Get-GssDriveBackupProperty $manifest @('runtime_contract')
    if ($null -eq $runtime -or
        (Get-GssDriveBackupProperty $runtime @('require_clean_tree')) -isnot [bool] -or
        -not [bool](Get-GssDriveBackupProperty $runtime @('require_clean_tree')) -or
        (Get-GssDriveBackupProperty $runtime @('reject_untracked_executables')) -isnot [bool] -or
        -not [bool](Get-GssDriveBackupProperty $runtime @('reject_untracked_executables')) -or
        (Get-GssDriveBackupProperty $runtime @('require_exact_tag_at_head')) -isnot [bool] -or
        -not [bool](Get-GssDriveBackupProperty $runtime @('require_exact_tag_at_head')) -or
        [string](Get-GssDriveBackupProperty $runtime @('automatic_sending')) -cne 'permanently_disabled' -or
        [string](Get-GssDriveBackupProperty $runtime @('live_execution')) -cne 'manual_apply_only' -or
        (Get-GssDriveBackupProperty $runtime @('excel_validation_receipt_required')) -isnot [bool] -or
        -not [bool](Get-GssDriveBackupProperty $runtime @('excel_validation_receipt_required')) -or
        [string](Get-GssDriveBackupProperty $runtime @('excel_validation_receipt_name')) -cne 'local-excel-validation-receipt.json' -or
        [string](Get-GssDriveBackupProperty $runtime @('excel_validation_receipt_relative_path')) -cne '_automation_runs/state/release/local-excel-validation-receipt.json') {
        throw 'ReleaseOnly manifest does not preserve the approved clean-tag, manual-apply, disabled-send, and Excel-receipt controls.'
    }

    $receipt = Read-GssDriveBackupJson -Path $receiptSource
    $receiptError = [string](Get-GssDriveBackupProperty $receipt @('Error'))
    $workbookPortable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $receipt @('WorkbookPath')))
    $sourceLogPortable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $receipt @('SourceRunLogPath')))
    if ([int](Get-GssDriveBackupProperty $receipt @('ReceiptSchemaVersion')) -ne 2 -or
        [string](Get-GssDriveBackupProperty $receipt @('Status')) -cne 'Passed' -or
        -not [string]::IsNullOrWhiteSpace($receiptError) -or
        [string](Get-GssDriveBackupProperty $receipt @('ReleaseTag')) -cne $tag -or
        [string](Get-GssDriveBackupProperty $receipt @('GitHead')) -notmatch '^[a-fA-F0-9]{40}$' -or
        [string]::IsNullOrWhiteSpace([string](Get-GssDriveBackupProperty $receipt @('ExcelVersion'))) -or
        [string](Get-GssDriveBackupProperty $receipt @('WorkbookSha256')) -notmatch '^[a-fA-F0-9]{64}$' -or
        [string](Get-GssDriveBackupProperty $receipt @('SourceRunFingerprint')) -notmatch '^[a-fA-F0-9]{64}$' -or
        [string]::IsNullOrWhiteSpace([string](Get-GssDriveBackupProperty $receipt @('SourceRunHostName'))) -or
        [string]::IsNullOrWhiteSpace([string](Get-GssDriveBackupProperty $receipt @('CertificationHostName'))) -or
        -not $workbookPortable.StartsWith('_automation_runs/test-output/', [System.StringComparison]::OrdinalIgnoreCase) -or
        -not $sourceLogPortable.StartsWith('_automation_runs/logs/', [System.StringComparison]::OrdinalIgnoreCase) -or
        [int](Get-GssDriveBackupProperty $receipt @('FormulaErrors') -1) -ne 0 -or
        [int](Get-GssDriveBackupProperty $receipt @('ConstantErrors') -1) -ne 0) {
        throw 'ReleaseOnly Excel receipt is not a passed copy-only validation for the exact release tag.'
    }

    if (-not [string]::IsNullOrWhiteSpace($GssRoot)) {
        $root = [System.IO.Path]::GetFullPath($GssRoot).TrimEnd('\', '/')
        $workbookPath = [System.IO.Path]::GetFullPath((Join-Path $root $workbookPortable.Replace('/', '\')))
        $sourceLogPath = [System.IO.Path]::GetFullPath((Join-Path $root $sourceLogPortable.Replace('/', '\')))
        foreach ($evidencePath in @($workbookPath, $sourceLogPath)) {
            [void](Get-GssDriveBackupRelativePath -Path $evidencePath -Root $root)
            [void](Test-GssDriveBackupNoLinkTraversal -Path $evidencePath -Root $root)
        }
        if ((Get-GssDriveBackupSha256 -Path $workbookPath) -cne ([string](Get-GssDriveBackupProperty $receipt @('WorkbookSha256'))).ToLowerInvariant()) {
            throw 'ReleaseOnly Excel receipt workbook hash does not match its copy-test artifact.'
        }
        $sourceRun = Read-GssDriveBackupJson -Path $sourceLogPath
        if ([string](Get-GssDriveBackupProperty $sourceRun @('Mode')) -cne 'CopyTestOnly' -or
            [string](Get-GssDriveBackupProperty $sourceRun @('TransactionStatus')) -cne 'Prepared' -or
            [string](Get-GssDriveBackupProperty $sourceRun @('ProgramRelease')) -cne $tag -or
            [string]::IsNullOrWhiteSpace([string](Get-GssDriveBackupProperty $sourceRun @('HostName'))) -or
            [string](Get-GssDriveBackupProperty $sourceRun @('HostName')) -cne [string](Get-GssDriveBackupProperty $receipt @('SourceRunHostName')) -or
            [string](Get-GssDriveBackupProperty $sourceRun @('RunFingerprint')) -cne [string](Get-GssDriveBackupProperty $receipt @('SourceRunFingerprint')) -or
            (Get-GssDriveBackupPreparedRunFingerprint -Run $sourceRun) -cne ([string](Get-GssDriveBackupProperty $sourceRun @('RunFingerprint'))).ToLowerInvariant() -or
            ([string](Get-GssDriveBackupProperty $sourceRun @('StagedWorkbookRelativePath'))).Replace('\', '/') -cne $workbookPortable -or
            [string](Get-GssDriveBackupProperty $sourceRun @('StagedWorkbookSha256')) -cne ([string](Get-GssDriveBackupProperty $receipt @('WorkbookSha256'))).ToLowerInvariant()) {
            throw 'ReleaseOnly Excel receipt does not match its Prepared copy-only source run.'
        }
    }

    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    }
    catch {
        if (-not ('System.IO.Compression.ZipFile' -as [type])) {
            throw "ReleaseOnly archive validation requires System.IO.Compression: $($_.Exception.Message)"
        }
    }

    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($archive)
        $entriesByPath = @{}
        foreach ($entry in @($zip.Entries)) {
            if ([string]::IsNullOrWhiteSpace([string]$entry.Name)) { continue }
            $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string]$entry.FullName)
            $key = $portable.ToLowerInvariant()
            if ($entriesByPath.ContainsKey($key)) {
                throw "ReleaseOnly archive contains duplicate path '$portable'."
            }
            $entriesByPath[$key] = $entry
        }

        $expectedPaths = @{}
        $manifestFiles = @(Get-GssDriveBackupProperty $manifest @('files') @())
        if ($manifestFiles.Count -lt 1) {
            throw 'ReleaseOnly manifest contains no program files.'
        }
        $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
        $canonicalUtf8 = New-Object System.Text.UTF8Encoding($false)
        foreach ($file in $manifestFiles) {
            $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $file @('path')))
            $key = $portable.ToLowerInvariant()
            if ($expectedPaths.ContainsKey($key) -or $portable.Equals('release/release-manifest.json', [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "ReleaseOnly manifest contains a duplicate or self-referential path '$portable'."
            }
            $expectedPaths[$key] = $true
            if (-not $entriesByPath.ContainsKey($key)) {
                throw "ReleaseOnly archive is missing manifest-bound program file '$portable'."
            }
            if ([string](Get-GssDriveBackupProperty $file @('hash_mode')) -cne 'utf8_lf') {
                throw "ReleaseOnly manifest uses an unsupported hash mode for '$portable'."
            }
            $entryBytes = Get-GssDriveBackupZipEntryByteArray -Entry $entriesByPath[$key]
            try {
                $entryText = $strictUtf8.GetString($entryBytes)
            }
            catch {
                throw "ReleaseOnly archive program file is not strict UTF-8 text: $portable"
            }
            $canonicalBytes = $canonicalUtf8.GetBytes($entryText.Replace("`r`n", "`n").Replace("`r", "`n"))
            if ([long](Get-GssDriveBackupProperty $file @('canonical_size_bytes')) -ne $canonicalBytes.Length -or
                [string](Get-GssDriveBackupProperty $file @('sha256')) -cne (Get-GssDriveBackupByteSha256 -Bytes $canonicalBytes)) {
                throw "ReleaseOnly archive program file does not match the release manifest: $portable"
            }
        }

        $manifestArchivePath = 'release/release-manifest.json'
        $manifestArchiveKey = $manifestArchivePath.ToLowerInvariant()
        if (-not $entriesByPath.ContainsKey($manifestArchiveKey)) {
            throw 'ReleaseOnly archive does not contain release/release-manifest.json.'
        }
        $expectedPaths[$manifestArchiveKey] = $true
        $archivedManifestBytes = Get-GssDriveBackupZipEntryByteArray -Entry $entriesByPath[$manifestArchiveKey]
        $sourceManifestBytes = [System.IO.File]::ReadAllBytes($manifestSource)
        try {
            $archivedManifestText = $strictUtf8.GetString($archivedManifestBytes)
            $sourceManifestText = $strictUtf8.GetString($sourceManifestBytes)
        }
        catch {
            throw 'ReleaseOnly archive and inventoried release manifests must both be strict UTF-8 text.'
        }
        $archivedManifestCanonicalBytes = $canonicalUtf8.GetBytes(
            $archivedManifestText.Replace("`r`n", "`n").Replace("`r", "`n")
        )
        $sourceManifestCanonicalBytes = $canonicalUtf8.GetBytes(
            $sourceManifestText.Replace("`r`n", "`n").Replace("`r", "`n")
        )
        if ($archivedManifestCanonicalBytes.Length -ne $sourceManifestCanonicalBytes.Length -or
            (Get-GssDriveBackupByteSha256 -Bytes $archivedManifestCanonicalBytes) -cne
                (Get-GssDriveBackupByteSha256 -Bytes $sourceManifestCanonicalBytes)) {
            throw 'ReleaseOnly archive manifest content does not match the inventoried release manifest after canonical line-ending normalization.'
        }

        foreach ($entryKey in @($entriesByPath.Keys)) {
            if (-not $expectedPaths.ContainsKey($entryKey) -and $entryKey -cne '.gitignore') {
                throw "ReleaseOnly archive contains an unmanifested program file: $($entriesByPath[$entryKey].FullName)"
            }
        }
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }

    return [pscustomobject]@{
        Version = $version
        Tag = $tag
        ArchiveName = $archiveName
        ManifestSha256 = Get-GssDriveBackupSha256 -Path $manifestSource
        ArchiveSha256 = Get-GssDriveBackupSha256 -Path $archive
        ExcelReceiptSha256 = Get-GssDriveBackupSha256 -Path $receiptSource
    }
}

function Assert-GssReleaseOnlyInventoryContract {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GssRoot,
        [Parameter(Mandatory)]
        [object[]]$Inventory
    )

    if (@($Inventory).Count -ne 3) {
        throw 'ReleaseOnly inventory must contain exactly three artifacts.'
    }
    $byRole = @{}
    foreach ($item in @($Inventory)) {
        $role = [string](Get-GssDriveBackupProperty $item @('Role', 'role'))
        if ($role -notin @('release_archive', 'release_manifest', 'release_excel_receipt') -or $byRole.ContainsKey($role)) {
            throw "ReleaseOnly inventory contains an unsupported or duplicate role '$role'."
        }
        if ([string](Get-GssDriveBackupProperty $item @('Classification', 'classification')) -cne 'restricted_operational') {
            throw "ReleaseOnly inventory role '$role' has an invalid classification."
        }
        $byRole[$role] = $item
    }
    foreach ($role in @('release_archive', 'release_manifest', 'release_excel_receipt')) {
        if (-not $byRole.ContainsKey($role)) {
            throw "ReleaseOnly inventory is missing required role '$role'."
        }
    }

    [void](Assert-GssReleaseOnlyArtifactSet `
        -ArchivePath ([string]$byRole['release_archive'].SourcePath) `
        -ManifestPath ([string]$byRole['release_manifest'].SourcePath) `
        -ExcelReceiptPath ([string]$byRole['release_excel_receipt'].SourcePath) `
        -GssRoot $GssRoot)
    return $true
}

function Assert-GssReleaseOnlySnapshotContract {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SnapshotDirectory,
        [Parameter(Mandatory)]
        [object[]]$Files
    )

    if (@($Files).Count -ne 3) {
        throw 'ReleaseOnly snapshot must contain exactly three artifacts.'
    }
    $root = [System.IO.Path]::GetFullPath($SnapshotDirectory).TrimEnd('\', '/')
    $byRole = @{}
    $expectedPortable = @{
        release_manifest = 'release/release-manifest.json'
        release_excel_receipt = 'release/local-excel-validation-receipt.json'
    }
    foreach ($file in @($Files)) {
        $role = [string](Get-GssDriveBackupProperty $file @('role', 'Role'))
        $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $file @('portable_path', 'PortablePath')))
        if ($role -notin @('release_archive', 'release_manifest', 'release_excel_receipt') -or $byRole.ContainsKey($role)) {
            throw "ReleaseOnly snapshot contains an unsupported or duplicate role '$role'."
        }
        if ([string](Get-GssDriveBackupProperty $file @('classification', 'Classification')) -cne 'restricted_operational') {
            throw "ReleaseOnly snapshot role '$role' has an invalid classification."
        }
        if ($role -eq 'release_archive') {
            if ($portable -notmatch '^release/gss-survey-workbook-automation-v\d+\.\d+\.\d+\.zip$') {
                throw 'ReleaseOnly snapshot archive portable path is invalid.'
            }
        }
        elseif (-not $portable.Equals($expectedPortable[$role], [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "ReleaseOnly snapshot role '$role' has an invalid portable path."
        }
        $snapshotRelative = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $file @('snapshot_path')))
        $sourcePath = [System.IO.Path]::GetFullPath((Join-Path $root $snapshotRelative.Replace('/', '\')))
        if (-not $sourcePath.StartsWith("$root\", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "ReleaseOnly snapshot role '$role' escaped the snapshot root."
        }
        $byRole[$role] = $sourcePath
    }
    foreach ($role in @('release_archive', 'release_manifest', 'release_excel_receipt')) {
        if (-not $byRole.ContainsKey($role)) {
            throw "ReleaseOnly snapshot is missing required role '$role'."
        }
    }

    [void](Assert-GssReleaseOnlyArtifactSet `
        -ArchivePath $byRole['release_archive'] `
        -ManifestPath $byRole['release_manifest'] `
        -ExcelReceiptPath $byRole['release_excel_receipt'])
    return $true
}

function Assert-GssRecoveryOnlyInventoryContract {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GssRoot,
        [Parameter(Mandatory)]
        [object[]]$Inventory
    )

    $root = [System.IO.Path]::GetFullPath($GssRoot).TrimEnd('\', '/')
    $roles = @(
        'recovered_historical_detail',
        'recovery_ledger',
        'recovery_manifest',
        'recovery_receipt',
        'recovery_qa',
        'recovery_run_summary'
    )
    $byRole = @{}
    foreach ($role in $roles) { $byRole[$role] = @() }
    foreach ($item in @($Inventory)) {
        $role = [string](Get-GssDriveBackupProperty $item @('Role', 'role'))
        if (-not $byRole.ContainsKey($role)) {
            throw "RecoveryOnly inventory contains unsupported role '$role'."
        }
        $byRole[$role] = @($byRole[$role]) + @($item)
    }

    foreach ($role in @('recovery_ledger', 'recovery_manifest', 'recovery_receipt', 'recovery_qa', 'recovery_run_summary')) {
        if (@($byRole[$role]).Count -ne 1) {
            throw "RecoveryOnly inventory requires exactly one '$role' record."
        }
    }
    if (@($byRole['recovered_historical_detail']).Count -lt 1) {
        throw "RecoveryOnly inventory requires at least one 'recovered_historical_detail' XLSX."
    }

    $manifestRecord = @($byRole['recovery_manifest'])[0]
    $receiptRecord = @($byRole['recovery_receipt'])[0]
    $ledgerRecord = @($byRole['recovery_ledger'])[0]
    $qaRecord = @($byRole['recovery_qa'])[0]
    $summaryRecord = @($byRole['recovery_run_summary'])[0]
    $manifestPath = [System.IO.Path]::GetFullPath([string]$manifestRecord.SourcePath)
    $receiptPath = [System.IO.Path]::GetFullPath([string]$receiptRecord.SourcePath)
    $ledgerPath = [System.IO.Path]::GetFullPath([string]$ledgerRecord.SourcePath)
    $qaPath = [System.IO.Path]::GetFullPath([string]$qaRecord.SourcePath)
    $summaryPath = [System.IO.Path]::GetFullPath([string]$summaryRecord.SourcePath)
    $transactionRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $manifestPath)).TrimEnd('\', '/')
    $transactionHash = [System.IO.Path]::GetFileName($transactionRoot).ToLowerInvariant()
    $manifestHash = Get-GssDriveBackupSha256 -Path $manifestPath
    if ($transactionHash -notmatch '^[a-f0-9]{64}$' -or $transactionHash -cne $manifestHash) {
        throw 'Recovery manifest SHA-256 must exactly match its manifest-hash transaction directory.'
    }
    foreach ($evidencePath in @($receiptPath, $qaPath, $summaryPath)) {
        $evidenceRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $evidencePath)).TrimEnd('\', '/')
        if (-not $evidenceRoot.Equals($transactionRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw 'Recovery transaction manifest, receipt, QA, and run summary must share one manifest-hash directory.'
        }
    }

    $manifest = Read-GssDriveBackupJson -Path $manifestPath
    $receipt = Read-GssDriveBackupJson -Path $receiptPath
    if ([string](Get-GssDriveBackupProperty $manifest @('schema_version')) -cne 'gss-historical-recovery/v1') {
        throw 'Recovery manifest schema_version must be gss-historical-recovery/v1.'
    }
    if ([string](Get-GssDriveBackupProperty $receipt @('schema_version')) -cne 'gss-historical-recovery-receipt/v1' -or
        [string](Get-GssDriveBackupProperty $receipt @('state')) -cne 'Committed') {
        throw 'Recovery transaction receipt must use gss-historical-recovery-receipt/v1 and be Committed.'
    }
    if ([string](Get-GssDriveBackupProperty $receipt @('manifest_sha256')) -cne $manifestHash -or
        [string](Get-GssDriveBackupProperty $receipt @('transaction_id')) -cne "historical-recovery:$manifestHash") {
        throw 'Recovery transaction receipt identity does not match the manifest SHA-256.'
    }
    $receiptManifestPath = [System.IO.Path]::GetFullPath([string](Get-GssDriveBackupProperty $receipt @('manifest_snapshot_path')))
    if (-not $receiptManifestPath.Equals($manifestPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Recovery receipt manifest_snapshot_path does not identify the inventoried manifest.'
    }

    $manifestSources = @(Get-GssDriveBackupProperty $manifest @('sources') @())
    $receiptFiles = @(Get-GssDriveBackupProperty $receipt @('files') @())
    $recoveredItems = @($byRole['recovered_historical_detail'])
    if ($manifestSources.Count -lt 1 -or
        $receiptFiles.Count -ne $manifestSources.Count -or
        $recoveredItems.Count -ne $manifestSources.Count) {
        throw 'Recovery manifest, receipt, and recovered XLSX inventory counts must match exactly.'
    }

    $recoveredBySourcePath = @{}
    foreach ($recoveredItem in $recoveredItems) {
        $sourcePath = [System.IO.Path]::GetFullPath([string]$recoveredItem.SourcePath)
        $key = $sourcePath.ToLowerInvariant()
        if ($recoveredBySourcePath.ContainsKey($key)) {
            throw "RecoveryOnly inventory repeats recovered XLSX '$sourcePath'."
        }
        $recoveredBySourcePath[$key] = $recoveredItem
    }
    $seenDestinations = @{}
    $totalRows = [long]0
    for ($sourceIndex = 0; $sourceIndex -lt $manifestSources.Count; $sourceIndex++) {
        $source = $manifestSources[$sourceIndex]
        $destinationRelative = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $source @('destination_path')))
        if (-not $destinationRelative.StartsWith('03 Uploaded Survey Workbooks/Archive - Previous Uploads/Recovered Historical Detail/', [System.StringComparison]::OrdinalIgnoreCase) -or
            [System.IO.Path]::GetExtension($destinationRelative) -ne '.xlsx') {
            throw "Recovery manifest source $($sourceIndex + 1) has an invalid recovered archive destination."
        }
        $destinationFullPath = [System.IO.Path]::GetFullPath((Join-Path $root $destinationRelative.Replace('/', '\')))
        [void](Get-GssDriveBackupRelativePath -Path $destinationFullPath -Root $root)
        $destinationKey = $destinationFullPath.ToLowerInvariant()
        if ($seenDestinations.ContainsKey($destinationKey) -or -not $recoveredBySourcePath.ContainsKey($destinationKey)) {
            throw "Recovery manifest destination is duplicated or absent from the closed inventory: $destinationRelative"
        }
        $seenDestinations[$destinationKey] = $true

        $sourceHash = [string](Get-GssDriveBackupProperty $source @('sha256'))
        $sourceBytes = [long](Get-GssDriveBackupProperty $source @('byte_size'))
        $sourceRows = [long](Get-GssDriveBackupProperty $source @('row_count'))
        if ($sourceHash -notmatch '^[a-f0-9]{64}$' -or
            $sourceBytes -lt 1 -or
            $sourceRows -lt 0 -or
            (Get-GssDriveBackupSha256 -Path $destinationFullPath) -cne $sourceHash -or
            [long](Get-Item -LiteralPath $destinationFullPath -Force).Length -ne $sourceBytes) {
            throw "Recovered XLSX bytes do not match manifest source $($sourceIndex + 1)."
        }
        $totalRows += $sourceRows

        $receiptFile = @($receiptFiles | Where-Object { [int](Get-GssDriveBackupProperty $_ @('source_index')) -eq ($sourceIndex + 1) })
        if ($receiptFile.Count -ne 1) {
            throw "Recovery receipt must identify manifest source $($sourceIndex + 1) exactly once."
        }
        $receiptDestinationFullPath = [System.IO.Path]::GetFullPath([string](Get-GssDriveBackupProperty $receiptFile[0] @('destination_full_path')))
        $receiptDestinationRelative = ([string](Get-GssDriveBackupProperty $receiptFile[0] @('destination_path'))).Replace('\', '/').Trim('/')
        if (-not $receiptDestinationFullPath.Equals($destinationFullPath, [System.StringComparison]::OrdinalIgnoreCase) -or
            -not $receiptDestinationRelative.Equals($destinationRelative, [System.StringComparison]::OrdinalIgnoreCase) -or
            [string](Get-GssDriveBackupProperty $receiptFile[0] @('sha256')) -cne $sourceHash -or
            [long](Get-GssDriveBackupProperty $receiptFile[0] @('byte_size')) -ne $sourceBytes -or
            [long](Get-GssDriveBackupProperty $receiptFile[0] @('row_count')) -ne $sourceRows) {
            throw "Recovery receipt file $($sourceIndex + 1) does not exactly match its manifest destination."
        }
    }
    if ($seenDestinations.Count -ne $recoveredBySourcePath.Count) {
        throw 'RecoveryOnly inventory contains an unrelated recovered XLSX not named by the manifest and receipt.'
    }

    $receiptLedgerPath = [System.IO.Path]::GetFullPath([string](Get-GssDriveBackupProperty $receipt @('ledger_path')))
    $receiptLedgerHash = [string](Get-GssDriveBackupProperty $receipt @('ledger_sha256_after'))
    if (-not $receiptLedgerPath.Equals($ledgerPath, [System.StringComparison]::OrdinalIgnoreCase) -or
        $receiptLedgerHash -notmatch '^[a-f0-9]{64}$' -or
        (Get-GssDriveBackupSha256 -Path $ledgerPath) -cne $receiptLedgerHash) {
        throw 'Recovery ledger path or SHA-256 does not match the Committed receipt.'
    }
    $ledger = Read-GssDriveBackupJson -Path $ledgerPath
    if ([string](Get-GssDriveBackupProperty $ledger @('schema_version')) -cne 'gss-feedback-first-seen/v1') {
        throw 'Recovery ledger schema_version must be gss-feedback-first-seen/v1.'
    }
    $ledgerEntryCount = @(Get-GssDriveBackupProperty $ledger @('entries') @()).Count
    $plannedEntryCount = @(Get-GssDriveBackupProperty $receipt @('planned_entries') @()).Count
    $insertedResponseCount = @(Get-GssDriveBackupProperty $receipt @('inserted_response_hashes') @()).Count
    $publishedFileCount = [int](Get-GssDriveBackupProperty $receipt @('published_file_count') 0)
    foreach ($receiptFileControl in $receiptFiles) {
        $publishedValue = Get-GssDriveBackupProperty $receiptFileControl @('published_by_transaction') $null
        if ($publishedValue -isnot [bool]) {
            throw 'Recovery receipt published_by_transaction controls must be JSON booleans.'
        }
    }
    $publishedReceiptCount = @($receiptFiles | Where-Object {
        (Get-GssDriveBackupProperty $_ @('published_by_transaction') $false) -eq $true
    }).Count
    if ($publishedFileCount -ne $publishedReceiptCount) {
        throw 'Recovery receipt published_file_count does not match its file controls.'
    }

    $qa = Read-GssDriveBackupJson -Path $qaPath
    $qaRequired = @(
        'schema_version',
        'status',
        'manifest_sha256',
        'transaction_id',
        'source_count',
        'recovered_file_count',
        'row_count',
        'unique_response_count',
        'inserted_response_count',
        'published_file_count',
        'ledger_entry_count_after',
        'controls'
    )
    [void](Assert-GssDriveBackupObjectShape -Value $qa -Required $qaRequired -Allowed ($qaRequired + @('generated_at_utc')) -Label 'Recovery QA')
    if ([string]$qa.schema_version -cne 'gss-historical-recovery-qa/v1' -or
        [string]$qa.status -cne 'Passed' -or
        [string]$qa.manifest_sha256 -cne $manifestHash -or
        [string]$qa.transaction_id -cne "historical-recovery:$manifestHash" -or
        [int]$qa.source_count -ne $manifestSources.Count -or
        [int]$qa.recovered_file_count -ne $recoveredItems.Count -or
        [long]$qa.row_count -ne $totalRows -or
        [int]$qa.unique_response_count -ne $plannedEntryCount -or
        [int]$qa.inserted_response_count -ne $insertedResponseCount -or
        [int]$qa.published_file_count -ne $publishedFileCount -or
        [int]$qa.ledger_entry_count_after -ne $ledgerEntryCount) {
        throw 'Recovery QA identity or aggregate counts do not match the manifest, receipt, and ledger.'
    }
    $qaControlNames = @(
        'manifest_hash_verified',
        'receipt_committed',
        'destinations_verified',
        'ledger_hash_verified',
        'no_unrelated_recovered_xlsx',
        'live_workbook_unchanged',
        'email_package_unchanged',
        'automatic_sending_disabled',
        'scheduled_task_disabled',
        'contains_row_level_data'
    )
    [void](Assert-GssDriveBackupObjectShape -Value $qa.controls -Required $qaControlNames -Allowed $qaControlNames -Label 'Recovery QA controls')
    foreach ($controlName in $qaControlNames) {
        $controlValue = Get-GssDriveBackupProperty $qa.controls @($controlName)
        if ($controlValue -isnot [bool]) {
            throw "Recovery QA control '$controlName' must be a JSON boolean."
        }
        if ($controlName -eq 'contains_row_level_data') {
            if ($controlValue) { throw 'Recovery QA must remain aggregate-only and cannot contain row-level data.' }
        }
        elseif (-not $controlValue) {
            throw "Recovery QA control '$controlName' must be true."
        }
    }

    $summary = Read-GssDriveBackupJson -Path $summaryPath
    $summaryRequired = @(
        'schema_version',
        'RunId',
        'RunFingerprint',
        'Folder',
        'CurrentWeekEnding',
        'ProgramRelease',
        'SnapshotPurpose',
        'BackupInventory'
    )
    [void](Assert-GssDriveBackupObjectShape -Value $summary -Required $summaryRequired -Allowed ($summaryRequired + @('GeneratedAtUtc')) -Label 'Recovery Drive run summary')
    if ([string]$summary.schema_version -cne 'gss-recovery-drive-summary/v1' -or
        [string]$summary.SnapshotPurpose -cne 'RecoveryOnly' -or
        [string]$summary.RunFingerprint -cne "sha256:$manifestHash" -or
        [string]::IsNullOrWhiteSpace([string]$summary.RunId)) {
        throw 'Recovery Drive run summary schema, purpose, run ID, or manifest fingerprint is invalid.'
    }
    $summaryRoot = [System.IO.Path]::GetFullPath([string]$summary.Folder).TrimEnd('\', '/')
    if (-not $summaryRoot.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Recovery Drive run summary Folder does not match the requested GSS root.'
    }
    $embeddedInventory = @($summary.BackupInventory)
    foreach ($embeddedItem in $embeddedInventory) {
        [void](Assert-GssDriveBackupObjectShape `
            -Value $embeddedItem `
            -Required @('SourcePath', 'PortablePath', 'Role', 'Classification') `
            -Allowed @('SourcePath', 'PortablePath', 'Role', 'Classification') `
            -Label 'Recovery Drive run summary BackupInventory record')
    }
    [void](Test-GssDriveBackupExactInventoryRecordSet -Expected $Inventory -Actual $embeddedInventory -Label 'Recovery Drive run summary embedded inventory')
    return $true
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

function Assert-GssDriveBackupSnapshotPathBudget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DestinationDirectory,
        [Parameter(Mandatory)]
        [object[]]$Files,
        [Parameter(Mandatory)]
        [string]$Label
    )

    $destinationRoot = [System.IO.Path]::GetFullPath($DestinationDirectory).TrimEnd('\', '/')
    $validated = 0
    foreach ($file in @($Files)) {
        $relative = Assert-GssDriveBackupSafeRelativePath -Path ([string](Get-GssDriveBackupProperty $file @('snapshot_path')))
        $plannedDestination = [System.IO.Path]::GetFullPath((Join-Path $destinationRoot $relative.Replace('/', '\')))
        if (-not $plannedDestination.StartsWith("$destinationRoot\", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "$Label path escaped its promoted snapshot root: $relative"
        }
        if ($plannedDestination.Length -ge $script:GssDriveBackupLegacySafePathLength) {
            throw "$Label path exceeds the safe Windows path budget at the promoted destination: $plannedDestination"
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
        [ValidateSet('WorkbookTransaction', 'RecoveryOnly', 'ReleaseOnly')]
        [string]$SnapshotPurpose = 'WorkbookTransaction',
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
                [string](Get-GssDriveBackupProperty $existing @('fingerprint')) -ne $Fingerprint -or
                [string](Get-GssDriveBackupProperty $existing @('snapshot_purpose') 'WorkbookTransaction') -ne $SnapshotPurpose) {
                throw "Existing prepared backup does not match the requested run: $partialPath"
            }
            $files = @(Get-GssDriveBackupProperty $existing @('files') @())
            [void](Test-GssDriveBackupPayload -SnapshotDirectory $partialPath -Files $files)
            if ($SnapshotPurpose -eq 'ReleaseOnly') {
                [void](Assert-GssReleaseOnlySnapshotContract -SnapshotDirectory $partialPath -Files $files)
            }
            if ($SnapshotPurpose -in @('RecoveryOnly', 'ReleaseOnly')) {
                $requestedFiles = @(ConvertTo-GssDriveBackupInventoryEvidence -Inventory $Inventory)
                [void](Test-GssDriveBackupExactFileSet -ExpectedFiles $files -ActualFiles $requestedFiles -Label "Idempotent $SnapshotPurpose preparation inventory")
            }
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
            # Budget against both staging and the eventual promoted location. The
            # latter is longer and is the path that must remain usable after commit.
            $promotedPath = Join-Path $context.RootPath ("snapshots\{0}\{1}\{2}" -f $ReportWeek.ToString('yyyy'), $ReportWeek.ToString('yyyy-MM'), $RunId)
            $files = @(Copy-GssDriveBackupInventory -Inventory $Inventory -SnapshotDirectory $partialPath -PayloadPrefix 'prepared-payload' -PathBudgetDirectories @($promotedPath))
            [void](Test-GssDriveBackupPayload -SnapshotDirectory $partialPath -Files $files)
            if ($SnapshotPurpose -eq 'ReleaseOnly') {
                [void](Assert-GssReleaseOnlySnapshotContract -SnapshotDirectory $partialPath -Files $files)
            }
            $chainHead = Get-GssDriveBackupChainHead -RootPath $context.RootPath
            $containsPersonalData = [bool](@($files | Where-Object classification -eq 'restricted_personal_data').Count -gt 0)
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
                snapshot_purpose = $SnapshotPurpose
                prior_manifest_sha256 = if ($chainHead) { $chainHead.ManifestSha256 } else { $null }
                drive = [ordered]@{
                    folder_id = $context.Settings.DriveFolderId
                    marker_id = $context.Settings.MarkerId
                    expected_owner = $context.Settings.ExpectedOwner
                    verification_level = $context.Settings.VerificationLevel
                }
                data_classification = [ordered]@{
                    label = if ($containsPersonalData) {
                        $script:GssDriveBackupClassificationLabel
                    }
                    else {
                        'RESTRICTED OPERATIONAL - NO PERSONAL DATA IN SNAPSHOT INVENTORY'
                    }
                    contains_personal_data = $containsPersonalData
                }
                scope = [ordered]@{
                    inventory_mode = if ($SnapshotPurpose -in @('RecoveryOnly', 'ReleaseOnly')) { $SnapshotPurpose } else { 'Full' }
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
    $receiptPurpose = [string](Get-GssDriveBackupProperty $receipt @('snapshot_purpose') 'WorkbookTransaction')
    $manifestPurpose = [string](Get-GssDriveBackupProperty $manifest @('snapshot_purpose') 'WorkbookTransaction')
    $preparedPurpose = [string](Get-GssDriveBackupProperty $preparedManifest @('snapshot_purpose') 'WorkbookTransaction')
    if ($receiptPurpose -notin @('WorkbookTransaction', 'RecoveryOnly', 'ReleaseOnly')) {
        throw "Committed snapshot purpose is unsupported: $receiptPurpose"
    }
    if ($receiptPurpose -ne $manifestPurpose -or $receiptPurpose -ne $preparedPurpose) {
        throw "Committed snapshot purpose does not match across its prepared manifest, final manifest, and receipt: $SnapshotPath"
    }
    $expectedInventoryMode = if ($receiptPurpose -in @('RecoveryOnly', 'ReleaseOnly')) { $receiptPurpose } else { 'Full' }
    $manifestInventoryMode = [string](Get-GssDriveBackupProperty (Get-GssDriveBackupProperty $manifest @('scope')) @('inventory_mode'))
    $preparedInventoryMode = [string](Get-GssDriveBackupProperty (Get-GssDriveBackupProperty $preparedManifest @('scope')) @('inventory_mode'))
    if ($manifestInventoryMode -cne $expectedInventoryMode -or $preparedInventoryMode -cne $expectedInventoryMode) {
        throw "Committed snapshot inventory mode does not match purpose '$receiptPurpose'."
    }
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
    if ($receiptPurpose -in @('RecoveryOnly', 'ReleaseOnly')) {
        [void](Test-GssDriveBackupExactFileSet -ExpectedFiles $preparedFiles -ActualFiles $files -Label "Committed $receiptPurpose inventory")
    }
    if ($receiptPurpose -eq 'ReleaseOnly') {
        [void](Assert-GssReleaseOnlySnapshotContract -SnapshotDirectory $SnapshotPath -Files $preparedFiles)
        [void](Assert-GssReleaseOnlySnapshotContract -SnapshotDirectory $SnapshotPath -Files $files)
    }
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
        [ValidateSet('WorkbookTransaction', 'RecoveryOnly', 'ReleaseOnly')]
        [string]$SnapshotPurpose = 'WorkbookTransaction',
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
                [string](Get-GssDriveBackupProperty $validated.Manifest @('fingerprint')) -ne $Fingerprint -or
                [string](Get-GssDriveBackupProperty $validated.Manifest @('snapshot_purpose') 'WorkbookTransaction') -ne $SnapshotPurpose) {
                throw 'Committed snapshot identity does not match the requested run.'
            }
            if ($SnapshotPurpose -in @('RecoveryOnly', 'ReleaseOnly') -and $null -ne $FinalInventory -and @($FinalInventory).Count -gt 0) {
                $requestedFiles = @(ConvertTo-GssDriveBackupInventoryEvidence -Inventory $FinalInventory)
                $committedFiles = @(Get-GssDriveBackupProperty $validated.Manifest @('files') @())
                [void](Test-GssDriveBackupExactFileSet -ExpectedFiles $committedFiles -ActualFiles $requestedFiles -Label "Idempotent committed $SnapshotPurpose inventory")
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
            [string](Get-GssDriveBackupProperty $prepared @('fingerprint')) -ne $Fingerprint -or
            [string](Get-GssDriveBackupProperty $prepared @('snapshot_purpose') 'WorkbookTransaction') -ne $SnapshotPurpose) {
            throw 'Prepared snapshot identity does not match the requested run.'
        }
        $preparedFiles = @(Get-GssDriveBackupProperty $prepared @('files') @())
        [void](Test-GssDriveBackupPayload -SnapshotDirectory $activePath -Files $preparedFiles)
        if ($SnapshotPurpose -eq 'ReleaseOnly') {
            [void](Assert-GssReleaseOnlySnapshotContract -SnapshotDirectory $activePath -Files $preparedFiles)
        }

        $manifestPath = Join-Path $activePath 'backup-manifest.json'
        if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
            $backupManifest = Read-GssDriveBackupJson -Path $manifestPath
            if ([string](Get-GssDriveBackupProperty $backupManifest @('run_id')) -ne $RunId -or
                [string](Get-GssDriveBackupProperty $backupManifest @('fingerprint')) -ne $Fingerprint -or
                [string](Get-GssDriveBackupProperty $backupManifest @('snapshot_purpose') 'WorkbookTransaction') -ne $SnapshotPurpose -or
                [string](Get-GssDriveBackupProperty $backupManifest @('prepared_manifest_sha256')) -ne (Get-GssDriveBackupSha256 -Path $preparedPath)) {
                throw 'Existing final manifest identity does not match the prepared snapshot.'
            }
            $finalFiles = @(Get-GssDriveBackupProperty $backupManifest @('files') @())
            [void](Test-GssDriveBackupPayload -SnapshotDirectory $activePath -Files $finalFiles)
            $manifestPreparedFiles = @(Get-GssDriveBackupProperty $backupManifest @('prepared_files') @())
            [void](Test-GssDriveBackupPayload -SnapshotDirectory $activePath -Files $manifestPreparedFiles)
            [void](Test-GssDriveBackupExactFileSet -ExpectedFiles $preparedFiles -ActualFiles $manifestPreparedFiles -Label 'Final manifest prepared inventory')
            if ($SnapshotPurpose -in @('RecoveryOnly', 'ReleaseOnly')) {
                [void](Test-GssDriveBackupExactFileSet -ExpectedFiles $preparedFiles -ActualFiles $finalFiles -Label "Existing $SnapshotPurpose final inventory")
            }
            if ($SnapshotPurpose -eq 'ReleaseOnly') {
                [void](Assert-GssReleaseOnlySnapshotContract -SnapshotDirectory $activePath -Files $finalFiles)
            }
        }
        else {
            if ($null -eq $FinalInventory -or @($FinalInventory).Count -eq 0) {
                throw 'Finalization requires the current curated inventory so the post-apply state can be captured.'
            }

            if ($SnapshotPurpose -in @('RecoveryOnly', 'ReleaseOnly')) {
                $preparedByPortablePath = @{}
                foreach ($preparedFile in @($preparedFiles)) {
                    $preparedKey = ([string](Get-GssDriveBackupProperty $preparedFile @('portable_path'))).ToLowerInvariant()
                    $preparedByPortablePath[$preparedKey] = $preparedFile
                }
                if (@($FinalInventory).Count -ne $preparedByPortablePath.Count) {
                    throw "$SnapshotPurpose final inventory does not match the prepared inventory count."
                }
                foreach ($item in @($FinalInventory)) {
                    $portable = Assert-GssDriveBackupSafeRelativePath -Path ([string]$item.PortablePath)
                    $key = $portable.ToLowerInvariant()
                    if (-not $preparedByPortablePath.ContainsKey($key)) {
                        throw "$SnapshotPurpose final inventory added an unprepared path: $portable"
                    }
                    $expected = $preparedByPortablePath[$key]
                    $actualHash = Get-GssDriveBackupSha256 -Path ([string]$item.SourcePath)
                    if ($actualHash -ne [string](Get-GssDriveBackupProperty $expected @('sha256')) -or
                        [string]$item.Role -ne [string](Get-GssDriveBackupProperty $expected @('role')) -or
                        [string]$item.Classification -ne [string](Get-GssDriveBackupProperty $expected @('classification'))) {
                        throw "$SnapshotPurpose final inventory changed after preparation: $portable"
                    }
                }
            }

            $reportWeekForBudget = [datetime]::ParseExact([string](Get-GssDriveBackupProperty $prepared @('report_week')), 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
            $promotedPathForBudget = Join-Path $context.RootPath ("snapshots\{0}\{1}\{2}" -f $reportWeekForBudget.ToString('yyyy'), $reportWeekForBudget.ToString('yyyy-MM'), $RunId)
            $finalFiles = @(Copy-GssDriveBackupInventory -Inventory $FinalInventory -SnapshotDirectory $activePath -PayloadPrefix 'payload' -PathBudgetDirectories @($promotedPathForBudget))
            [void](Test-GssDriveBackupPayload -SnapshotDirectory $activePath -Files $finalFiles)
            if ($SnapshotPurpose -in @('RecoveryOnly', 'ReleaseOnly')) {
                [void](Test-GssDriveBackupExactFileSet -ExpectedFiles $preparedFiles -ActualFiles $finalFiles -Label "Copied $SnapshotPurpose final inventory")
            }
            if ($SnapshotPurpose -eq 'ReleaseOnly') {
                [void](Assert-GssReleaseOnlySnapshotContract -SnapshotDirectory $activePath -Files $finalFiles)
            }
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
                snapshot_purpose = [string](Get-GssDriveBackupProperty $prepared @('snapshot_purpose') 'WorkbookTransaction')
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

        $manifestPreparedFiles = @(Get-GssDriveBackupProperty $backupManifest @('prepared_files') @())
        $reportWeekText = [string](Get-GssDriveBackupProperty $prepared @('report_week'))
        $reportWeek = [datetime]::ParseExact($reportWeekText, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $snapshotParent = Join-Path $context.RootPath ("snapshots\{0}\{1}" -f $reportWeek.ToString('yyyy'), $reportWeek.ToString('yyyy-MM'))
        $finalPath = if ($location.IsPartial) {
            Join-Path $snapshotParent $RunId
        }
        else {
            $activePath
        }
        [void](Assert-GssDriveBackupSnapshotPathBudget -DestinationDirectory $finalPath -Files $preparedFiles -Label 'Prepared manifest snapshot')
        [void](Assert-GssDriveBackupSnapshotPathBudget -DestinationDirectory $finalPath -Files $manifestPreparedFiles -Label 'Final manifest prepared snapshot')
        [void](Assert-GssDriveBackupSnapshotPathBudget -DestinationDirectory $finalPath -Files $finalFiles -Label 'Final manifest snapshot')

        if ($location.IsPartial) {
            if (-not (Test-Path -LiteralPath $snapshotParent -PathType Container)) {
                New-Item -ItemType Directory -Path $snapshotParent -Force | Out-Null
            }
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
            snapshot_purpose = [string](Get-GssDriveBackupProperty $manifest @('snapshot_purpose') 'WorkbookTransaction')
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
        $restoreRelative = $portable
        $target = Join-Path $destination $restoreRelative.Replace('/', '\')
        $usesReservedCompactNamespace = $portable.Equals('r', [System.StringComparison]::OrdinalIgnoreCase) -or
            $portable.Equals('r/long-path', [System.StringComparison]::OrdinalIgnoreCase) -or
            $portable.StartsWith('r/long-path/', [System.StringComparison]::OrdinalIgnoreCase)
        foreach ($reservedRootPath in @(
            'restore-verification.json',
            'local-excel-validation-receipt.json',
            'quarterly-restore-drill.json'
        )) {
            if ($portable.Equals($reservedRootPath, [System.StringComparison]::OrdinalIgnoreCase) -or
                $portable.StartsWith("$reservedRootPath/", [System.StringComparison]::OrdinalIgnoreCase)) {
                $usesReservedCompactNamespace = $true
                break
            }
        }
        if ($target.Length -ge $script:GssDriveBackupLegacySafePathLength -or $usesReservedCompactNamespace) {
            # The restore prefix is deliberately minimal: LOCALAPPDATA and the
            # isolated receipt directory already consume substantial MAX_PATH.
            $restoreRelative = Get-GssDriveBackupCompactRelativePath -PortablePath $portable -Prefix 'r'
            $target = Join-Path $destination $restoreRelative.Replace('/', '\')
            if ($target.Length -ge $script:GssDriveBackupLegacySafePathLength) {
                $restoreRelative = Get-GssDriveBackupCompactRelativePath -PortablePath $portable -Prefix 'r' -OmitExtension
                $target = Join-Path $destination $restoreRelative.Replace('/', '\')
            }
            if ($target.Length -ge $script:GssDriveBackupLegacySafePathLength) {
                throw "Compacted verify-only restore destination still exceeds the safe Windows path budget before copy. Shorten LOCALAPPDATA: $target"
            }
        }
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
            restored_path = $restoreRelative
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

function Resolve-GssDriveBackupRestoredFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ReceiptPath,
        [Parameter(Mandatory)]
        [string]$PortablePath,
        [string]$ExpectedDestination
    )

    $portable = Assert-GssDriveBackupSafeRelativePath -Path $PortablePath
    $receipt = Read-GssDriveBackupJson -Path $ReceiptPath
    if ([int](Get-GssDriveBackupProperty $receipt @('schema_version')) -ne 1 -or
        [string](Get-GssDriveBackupProperty $receipt @('operation')) -cne 'verify_only_restore' -or
        [string](Get-GssDriveBackupProperty $receipt @('status')) -cne 'Verified' -or
        [bool](Get-GssDriveBackupProperty $receipt @('live_workbook_overwritten') $true)) {
        throw "Restore receipt is not verified, isolated restore evidence: $ReceiptPath"
    }

    $destination = [System.IO.Path]::GetFullPath([string](Get-GssDriveBackupProperty $receipt @('destination'))).TrimEnd('\', '/')
    if (-not [string]::IsNullOrWhiteSpace($ExpectedDestination)) {
        $expected = [System.IO.Path]::GetFullPath($ExpectedDestination).TrimEnd('\', '/')
        if (-not $destination.Equals($expected, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Restore receipt destination does not match the completed restore: $destination"
        }
    }

    $matchingRecords = @(
        Get-GssDriveBackupProperty $receipt @('files') @() |
            Where-Object {
                [string](Get-GssDriveBackupProperty $_ @('portable_path')) -ieq $portable
            }
    )
    if ($matchingRecords.Count -ne 1) {
        throw "Restore receipt must contain exactly one mapping for portable path '$portable'; found $($matchingRecords.Count)."
    }

    $record = $matchingRecords[0]
    $restoredRelative = [string](Get-GssDriveBackupProperty $record @('restored_path') $portable)
    if ([string]::IsNullOrWhiteSpace($restoredRelative)) {
        $restoredRelative = $portable
    }
    $restoredRelative = Assert-GssDriveBackupSafeRelativePath -Path $restoredRelative
    $path = [System.IO.Path]::GetFullPath((Join-Path $destination $restoredRelative.Replace('/', '\')))
    if (-not $path.StartsWith("$destination\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Restore receipt mapping escaped its isolated destination: $restoredRelative"
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Restore receipt mapping points to a missing file: $path"
    }
    $expectedHash = [string](Get-GssDriveBackupProperty $record @('sha256'))
    if ((Get-GssDriveBackupSha256 -Path $path) -ne $expectedHash) {
        throw "Restore receipt mapping hash does not match the restored file: $path"
    }

    return [pscustomobject]@{
        PortablePath = $portable
        RestoredPath = $restoredRelative
        Path = $path
        Destination = $destination
        Sha256 = $expectedHash
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
