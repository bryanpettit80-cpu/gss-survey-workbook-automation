[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$LiveRunLogPath,
    [string]$Folder,
    [switch]$OutputObject
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
$releaseValidator = Join-Path $scriptRoot 'Test-GSS-ReleaseIntegrity.ps1'
$releaseIntegrity = & $releaseValidator -RepoRoot $repoRoot

function Write-GssResumeAtomicJson {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][object]$InputObject)

    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $temporaryPath = Join-Path $directory ('.t-{0}' -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
    $replacementBackupPath = Join-Path $directory ('.b-{0}' -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
    try {
        [System.IO.File]::WriteAllText(
            $temporaryPath,
            ($InputObject | ConvertTo-Json -Depth 10),
            (New-Object System.Text.UTF8Encoding($false))
        )
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            for ($attempt = 1; $attempt -le 20; $attempt++) {
                try {
                    [System.IO.File]::Replace($temporaryPath, $Path, $replacementBackupPath, $true)
                    break
                }
                catch [System.IO.IOException] {
                    if ($attempt -eq 20) { throw }
                    if (Test-Path -LiteralPath $replacementBackupPath -PathType Leaf) {
                        Remove-Item -LiteralPath $replacementBackupPath -Force
                    }
                    Start-Sleep -Milliseconds 250
                }
            }
        }
        else {
            [System.IO.File]::Move($temporaryPath, $Path)
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

function Get-GssResumeHash {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required run evidence is missing: $Path"
    }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-GssResumeTextHash {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($Text)
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-GssResumeRunFingerprint {
    param([Parameter(Mandatory)][object]$Run)

    return Get-GssResumeTextHash (@(
        'gss-transaction-v1',
        ([string]$Run.RunId).Trim().ToLowerInvariant(),
        ([string]$Run.HostName).Trim().ToLowerInvariant(),
        ([string]$Run.CurrentWeekEnding).Trim(),
        ([string]$Run.StartingWorkbookSha256).Trim().ToLowerInvariant(),
        ([string]$Run.CurrentSourceSha256).Trim().ToLowerInvariant(),
        ([string]$Run.PriorYearSourceSha256).Trim().ToLowerInvariant(),
        ([string]$Run.StagedWorkbookSha256).Trim().ToLowerInvariant(),
        ([string]$Run.StagedPdfSha256).Trim().ToLowerInvariant(),
        ([string]$Run.ProgramRelease).Trim().ToLowerInvariant()
    ) -join "`n")
}

function Add-GssResumeReceiptValue {
    param([Parameter(Mandatory)][object]$Receipt, [Parameter(Mandatory)][string]$Name, [object]$Value)
    if ($Receipt.PSObject.Properties.Name -contains $Name) {
        $Receipt.$Name = $Value
    }
    else {
        $Receipt | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

function Get-GssResumeReleaseVersion {
    param([Parameter(Mandatory)][string]$ReleaseTag)

    $match = [regex]::Match(
        $ReleaseTag,
        '^v(?<major>0|[1-9][0-9]*)\.(?<minor>0|[1-9][0-9]*)\.(?<patch>0|[1-9][0-9]*)$',
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    )
    if (-not $match.Success) {
        throw "GSS release tag '$ReleaseTag' is not a canonical vMAJOR.MINOR.PATCH tag."
    }

    return [pscustomobject]@{
        Major = [long]::Parse($match.Groups['major'].Value, [System.Globalization.CultureInfo]::InvariantCulture)
        Minor = [long]::Parse($match.Groups['minor'].Value, [System.Globalization.CultureInfo]::InvariantCulture)
        Patch = [long]::Parse($match.Groups['patch'].Value, [System.Globalization.CultureInfo]::InvariantCulture)
    }
}

function Get-GssResumeReleaseMode {
    param(
        [Parameter(Mandatory)][string]$RunRelease,
        [Parameter(Mandatory)][string]$CurrentRelease
    )

    [void](Get-GssResumeReleaseVersion -ReleaseTag $RunRelease)
    [void](Get-GssResumeReleaseVersion -ReleaseTag $CurrentRelease)
    if ($RunRelease -ceq $CurrentRelease) {
        return 'CurrentRelease'
    }

    # Cross-release recovery is exceptional and must be declared here through a
    # code-reviewed release. Adjacency alone never grants compatibility.
    $packageRecoveryCompatibility = @(
        [pscustomobject]@{
            RunRelease = 'v1.1.8'
            CurrentRelease = 'v1.1.9'
        }
    )
    $declaredMatches = @($packageRecoveryCompatibility | Where-Object {
        [string]$_.RunRelease -ceq $RunRelease -and
        [string]$_.CurrentRelease -ceq $CurrentRelease
    })
    if ($declaredMatches.Count -eq 1) {
        return 'DeclaredPriorReleasePackageRecovery'
    }

    throw "Live run release '$RunRelease' does not have a code-reviewed package-recovery compatibility declaration for current release '$CurrentRelease'."
}

function Get-GssResumeAnnotatedReleaseCommit {
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][string]$ReleaseTag
    )

    [void](Get-GssResumeReleaseVersion -ReleaseTag $ReleaseTag)
    $git = Get-Command git.exe -ErrorAction Stop
    $tagReference = "refs/tags/$ReleaseTag"
    $tagType = @(& $git.Source -C $RepoRoot cat-file -t $tagReference 2>&1)
    if ($LASTEXITCODE -ne 0 -or $tagType.Count -ne 1 -or [string]$tagType[0].Trim() -cne 'tag') {
        throw "Prior-release recovery requires '$ReleaseTag' to exist locally as an annotated Git tag."
    }

    $peeledReference = "$tagReference^{commit}"
    $peeledCommit = @(& $git.Source -C $RepoRoot rev-parse --verify $peeledReference 2>&1)
    if ($LASTEXITCODE -ne 0 -or
        $peeledCommit.Count -ne 1 -or
        ([string]$peeledCommit[0].Trim()) -cnotmatch '^[a-f0-9]{40}$') {
        throw "Annotated release tag '$ReleaseTag' could not be peeled to one commit."
    }
    return [string]$peeledCommit[0].Trim()
}

function Test-GssResumeSamePath {
    param([string]$Left, [string]$Right)

    if ([string]::IsNullOrWhiteSpace($Left) -or [string]::IsNullOrWhiteSpace($Right)) {
        return $false
    }
    try {
        return [System.IO.Path]::GetFullPath($Left).Equals(
            [System.IO.Path]::GetFullPath($Right),
            [System.StringComparison]::OrdinalIgnoreCase
        )
    }
    catch {
        return $false
    }
}

function ConvertTo-GssResumeManifestPortablePath {
    param([Parameter(Mandatory)][string]$Path)

    $normalized = $Path.Replace('\', '/')
    $segments = @($normalized.Split('/'))
    if ($segments.Count -lt 2 -or
        $segments[0] -cne 'gss' -or
        @($segments | Where-Object { [string]::IsNullOrWhiteSpace($_) -or $_ -in @('.', '..') }).Count -gt 0) {
        throw "Committed snapshot contains an unsafe or noncanonical portable path: $Path"
    }
    return ($segments -join '/')
}

function ConvertTo-GssResumePortablePath {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$FolderPath
    )

    $folderRoot = [System.IO.Path]::GetFullPath($FolderPath).TrimEnd('\', '/')
    $folderPrefix = "$folderRoot\"
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if (-not $fullPath.StartsWith($folderPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Prior-release package input is outside the configured GSS folder: $fullPath"
    }
    $relativePath = $fullPath.Substring($folderPrefix.Length).Replace('\', '/')
    return ConvertTo-GssResumeManifestPortablePath -Path "gss/$relativePath"
}

function Get-GssResumePackageInputDescriptor {
    param(
        [Parameter(Mandatory)][string]$Kind,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$FolderPath
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required prior-release package input is missing: $Path"
    }
    $item = Get-Item -LiteralPath $Path
    return [pscustomobject]@{
        Kind = $Kind
        PortablePath = ConvertTo-GssResumePortablePath -Path $item.FullName -FolderPath $FolderPath
        FullPath = $item.FullName
        ByteSize = [long]$item.Length
        Sha256 = Get-GssResumeHash -Path $item.FullName
    }
}

function Get-GssResumeTrustedBackupManifest {
    param([Parameter(Mandatory)][object]$SnapshotEvidence)

    $manifestPath = [string]$SnapshotEvidence.BackupManifestPath
    $expectedHash = ([string]$SnapshotEvidence.BackupManifestSha256).ToLowerInvariant()
    if ($expectedHash -notmatch '^[a-f0-9]{64}$' -or
        (Get-GssResumeHash -Path $manifestPath) -cne $expectedHash) {
        throw 'The committed backup manifest changed after snapshot validation.'
    }
    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    }
    catch {
        throw 'The committed backup manifest is no longer valid JSON.'
    }
    if ([string]$manifest.status -cne 'Committed' -or
        [string]$manifest.run_id -cne [string]$SnapshotEvidence.RunId -or
        [string]$manifest.fingerprint -cne [string]$SnapshotEvidence.Fingerprint -or
        @($manifest.files).Count -eq 0) {
        throw 'The committed backup manifest no longer matches the validated snapshot identity.'
    }
    return $manifest
}

function Assert-GssResumePackageInputSet {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$LiveRunLogPath,
        [Parameter(Mandatory)][object]$Run,
        [Parameter(Mandatory)][object]$CommittedManifest
    )

    $folderRoot = [System.IO.Path]::GetFullPath($FolderPath).TrimEnd('\', '/')
    $detailFolder = Join-Path $folderRoot '03 Uploaded Survey Workbooks'
    if (-not (Test-Path -LiteralPath $detailFolder -PathType Container)) {
        throw 'Prior-release package input verification requires the guest-detail folder.'
    }

    $actualSources = @(
        [pscustomobject]@{ Kind = 'live_workbook'; Path = [string]$Run.TargetWorkbook },
        [pscustomobject]@{ Kind = 'comparison_pdf'; Path = [string]$Run.EmailComparisonPdf },
        [pscustomobject]@{ Kind = 'rolling_workbook'; Path = [string]$Run.CurrentSourceWorkbook },
        [pscustomobject]@{ Kind = 'prior_year_rolling_workbook'; Path = [string]$Run.PriorYearSourceWorkbook },
        [pscustomobject]@{ Kind = 'run_log'; Path = $LiveRunLogPath }
    )
    $actualSources += @(Get-ChildItem -LiteralPath $detailFolder -File -Filter '*.xlsx' -Recurse |
        Where-Object { $_.Name -notlike '~$*' -and $_.Length -gt 0 } |
        Sort-Object FullName |
        ForEach-Object { [pscustomobject]@{ Kind = 'detail_workbook'; Path = $_.FullName } })

    $ledgerPath = Join-Path $folderRoot '_automation_runs\state\gss_feedback_first_seen.json'
    if (Test-Path -LiteralPath $ledgerPath -PathType Leaf) {
        $actualSources += [pscustomobject]@{ Kind = 'feedback_ledger'; Path = $ledgerPath }
    }

    # Historical-recovery receipts and manifests are verification metadata, not
    # analysis content, and the workbook-transaction snapshot predates their
    # separate committed recovery evidence. Get-GssDetailInventory independently
    # validates their committed state, manifest hash, exact workbook bytes, row
    # counts, and response-set hashes. The recovered XLSX bytes themselves remain
    # part of this exact snapshot-bound detail-workbook set.

    $actualDescriptors = @($actualSources | ForEach-Object {
        Get-GssResumePackageInputDescriptor -Kind $_.Kind -Path $_.Path -FolderPath $folderRoot
    })
    $fixedKindByPath = @{}
    foreach ($descriptor in @($actualDescriptors | Where-Object Kind -in @(
        'live_workbook',
        'comparison_pdf',
        'rolling_workbook',
        'prior_year_rolling_workbook',
        'run_log'
    ))) {
        $key = ([string]$descriptor.PortablePath).ToLowerInvariant()
        if ($fixedKindByPath.ContainsKey($key)) {
            throw "Prior-release fixed package input was enumerated more than once: $($descriptor.PortablePath)"
        }
        $fixedKindByPath[$key] = [string]$descriptor.Kind
    }
    $liveRunLogPortablePath = ConvertTo-GssResumePortablePath -Path $LiveRunLogPath -FolderPath $folderRoot
    $ledgerPortablePath = 'gss/_automation_runs/state/gss_feedback_first_seen.json'
    $detailPrefix = 'gss/03 Uploaded Survey Workbooks/'

    $expectedDescriptors = @()
    foreach ($record in @($CommittedManifest.files)) {
        $rawPortablePath = ([string]$record.portable_path).Replace('\', '/')
        if (-not $rawPortablePath.StartsWith('gss/', [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        $portablePath = ConvertTo-GssResumeManifestPortablePath -Path $rawPortablePath
        $leafName = @($portablePath.Split('/'))[-1]
        $isDetailWorkbook = (
            $portablePath.StartsWith($detailPrefix, [System.StringComparison]::OrdinalIgnoreCase) -and
            $portablePath.EndsWith('.xlsx', [System.StringComparison]::OrdinalIgnoreCase) -and
            $leafName -notlike '~$*' -and
            [long]$record.byte_size -gt 0
        )
        $portablePathKey = $portablePath.ToLowerInvariant()
        $kind = if ($fixedKindByPath.ContainsKey($portablePathKey)) {
            $fixedKindByPath[$portablePathKey]
        }
        elseif ($isDetailWorkbook) {
            'detail_workbook'
        }
        elseif ($portablePath.Equals($ledgerPortablePath, [System.StringComparison]::OrdinalIgnoreCase)) {
            'feedback_ledger'
        }
        else {
            $null
        }
        if ($null -eq $kind) { continue }

        $sha256 = ([string]$record.sha256).ToLowerInvariant()
        try { $byteSize = [long]$record.byte_size }
        catch { throw "Committed package-input record has invalid byte-size evidence: $portablePath" }
        if ($byteSize -lt 0 -or $sha256 -notmatch '^[a-f0-9]{64}$') {
            throw "Committed package-input record has invalid size or hash evidence: $portablePath"
        }
        $expectedDescriptors += [pscustomobject]@{
            Kind = $kind
            PortablePath = $portablePath
            ByteSize = $byteSize
            Sha256 = $sha256
        }
    }

    $actualByPath = @{}
    foreach ($descriptor in $actualDescriptors) {
        $key = ([string]$descriptor.PortablePath).ToLowerInvariant()
        if ($actualByPath.ContainsKey($key)) {
            throw "Prior-release package input was enumerated more than once: $($descriptor.PortablePath)"
        }
        $actualByPath[$key] = $descriptor
    }
    $expectedByPath = @{}
    foreach ($descriptor in $expectedDescriptors) {
        $key = ([string]$descriptor.PortablePath).ToLowerInvariant()
        if ($expectedByPath.ContainsKey($key)) {
            throw "Committed snapshot contains duplicate package-input evidence: $($descriptor.PortablePath)"
        }
        $expectedByPath[$key] = $descriptor
    }

    foreach ($key in @($actualByPath.Keys | Sort-Object)) {
        if (-not $expectedByPath.ContainsKey($key)) {
            throw "Prior-release package input is absent from the committed snapshot: $($actualByPath[$key].PortablePath)"
        }
    }
    foreach ($key in @($expectedByPath.Keys | Sort-Object)) {
        if (-not $actualByPath.ContainsKey($key)) {
            throw "Committed package input is missing from the live source set: $($expectedByPath[$key].PortablePath)"
        }
        $actual = $actualByPath[$key]
        $expected = $expectedByPath[$key]
        if ([string]$actual.Kind -cne [string]$expected.Kind -or
            [long]$actual.ByteSize -ne [long]$expected.ByteSize -or
            [string]$actual.Sha256 -cne [string]$expected.Sha256) {
            throw "Prior-release package input changed after the committed snapshot: $($actual.PortablePath)"
        }
    }
    if ($actualByPath.Count -ne $expectedByPath.Count) {
        throw 'Prior-release package input exact-set verification failed.'
    }

    $sanitizedInputs = @($actualByPath.Keys | Sort-Object | ForEach-Object {
        $descriptor = $actualByPath[$_]
        [pscustomobject][ordered]@{
            PortablePath = [string]$descriptor.PortablePath
            ByteSize = [long]$descriptor.ByteSize
            Sha256 = [string]$descriptor.Sha256
        }
    })
    $canonicalEvidence = @($sanitizedInputs | ForEach-Object {
        "$(([string]$_.PortablePath).ToLowerInvariant()):$($_.ByteSize):$($_.Sha256)"
    }) -join "`n"
    return [pscustomobject]@{
        Inputs = $sanitizedInputs
        SourceSetSha256 = Get-GssResumeTextHash -Text $canonicalEvidence
        InputCount = $actualByPath.Count
        DetailWorkbookCount = @($actualDescriptors | Where-Object Kind -eq 'detail_workbook').Count
        FeedbackLedgerIncluded = @($actualDescriptors | Where-Object Kind -eq 'feedback_ledger').Count -eq 1
        LiveRunLogPortablePath = $liveRunLogPortablePath
        HistoricalRecoveryMetadataValidation = 'SeparateCommittedRecoveryOnlySnapshots'
    }
}

function Assert-GssResumeHistoricalRecoveryEvidence {
    param([Parameter(Mandatory)][string]$FolderPath)

    $runtimeRoot = Join-Path ([System.IO.Path]::GetFullPath($FolderPath)) '_automation_runs\historical-recovery'
    $receiptPaths = @()
    if (Test-Path -LiteralPath $runtimeRoot -PathType Container) {
        $receiptPaths = @(Get-ChildItem -LiteralPath $runtimeRoot -Directory |
            ForEach-Object {
                $candidate = Join-Path $_.FullName 'transaction-receipt.json'
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { $candidate }
            } |
            Sort-Object)
    }
    if ($receiptPaths.Count -eq 0) {
        return [pscustomobject]@{
            TransactionCount = 0
            EvidenceFileCount = 0
            EvidenceSetSha256 = Get-GssResumeTextHash -Text 'gss-historical-recovery-evidence/v1'
            Verification = 'CommittedRecoveryOnlySnapshots'
        }
    }

    $context = Get-GssDriveBackupRootContext
    $canonicalEvidence = @()
    foreach ($receiptPath in $receiptPaths) {
        $transactionDirectory = Split-Path -Parent $receiptPath
        $summaryPath = Join-Path $transactionDirectory 'drive-recovery-summary.json'
        $manifestPath = Join-Path $transactionDirectory 'recovery-manifest.json'
        foreach ($requiredPath in @($summaryPath, $manifestPath)) {
            if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
                throw "Historical recovery evidence is incomplete: $requiredPath"
            }
        }
        try {
            $summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
        }
        catch {
            throw "Historical recovery Drive summary is not valid JSON: $summaryPath"
        }
        $runId = [string]$summary.RunId
        $fingerprint = [string]$summary.RunFingerprint
        if ([string]$summary.schema_version -cne 'gss-recovery-drive-summary/v1' -or
            [string]$summary.SnapshotPurpose -cne 'RecoveryOnly' -or
            [string]::IsNullOrWhiteSpace($runId) -or
            $fingerprint -cnotmatch '^sha256:[a-f0-9]{64}$') {
            throw "Historical recovery Drive summary identity is invalid: $summaryPath"
        }

        $location = Find-GssDriveBackupSnapshot -RootPath $context.RootPath -RunId $runId
        if ($null -eq $location -or [bool]$location.IsPartial) {
            throw "Historical recovery requires an existing committed RecoveryOnly snapshot: $runId"
        }
        $validated = Test-GssCommittedBackupSnapshot -SnapshotPath ([string]$location.Path)
        foreach ($document in @($validated.Receipt, $validated.Manifest, $validated.PreparedManifest)) {
            if ([string]$document.run_id -cne $runId -or
                [string]$document.fingerprint -cne $fingerprint -or
                [string]$document.snapshot_purpose -cne 'RecoveryOnly') {
                throw "Historical RecoveryOnly snapshot identity does not match its live summary: $runId"
            }
        }
        if ([string]$validated.Receipt.status -cne 'Committed' -or
            [string]$validated.Receipt.verification_level -cne 'drivefs_hash_verified' -or
            [string]$validated.Manifest.status -cne 'Committed' -or
            [string]$validated.Manifest.drive.verification_level -cne 'drivefs_hash_verified' -or
            [string]$validated.PreparedManifest.status -cne 'Prepared' -or
            [string]$validated.PreparedManifest.drive.verification_level -cne 'drivefs_hash_verified') {
            throw "Historical recovery snapshot is not committed: $runId"
        }

        $roleBindings = @(
            [pscustomobject]@{ Role = 'recovery_run_summary'; PortablePath = 'recovery/evidence/drive-recovery-summary.json'; LivePath = $summaryPath },
            [pscustomobject]@{ Role = 'recovery_manifest'; PortablePath = 'recovery/evidence/recovery-manifest.json'; LivePath = $manifestPath },
            [pscustomobject]@{ Role = 'recovery_receipt'; PortablePath = 'recovery/evidence/transaction-receipt.json'; LivePath = $receiptPath }
        )
        foreach ($binding in $roleBindings) {
            $roleMatches = @($validated.Manifest.files | Where-Object {
                [string]$_.role -ceq [string]$binding.Role -and
                [string]$_.portable_path -ceq [string]$binding.PortablePath
            })
            if ($roleMatches.Count -ne 1) {
                throw "Historical RecoveryOnly snapshot does not contain exactly one '$($binding.Role)' record: $runId"
            }
            $record = $roleMatches[0]
            $item = Get-Item -LiteralPath $binding.LivePath
            $liveHash = Get-GssResumeHash -Path $binding.LivePath
            if ([long]$record.byte_size -ne [long]$item.Length -or
                [string]$record.sha256 -cne $liveHash) {
                throw "Historical recovery evidence changed after its committed RecoveryOnly snapshot: $($binding.LivePath)"
            }
            $canonicalEvidence += "$runId`:$($binding.Role):$($item.Length):$liveHash"
        }
    }

    return [pscustomobject]@{
        TransactionCount = $receiptPaths.Count
        EvidenceFileCount = $receiptPaths.Count * 3
        EvidenceSetSha256 = Get-GssResumeTextHash -Text (@($canonicalEvidence | Sort-Object) -join "`n")
        Verification = 'CommittedRecoveryOnlySnapshots'
    }
}

function Assert-GssResumePriorReleaseReceipt {
    param(
        [Parameter(Mandatory)][object]$Run,
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][string]$LiveRunLogPath,
        [Parameter(Mandatory)][string]$ExpectedReleaseCommit
    )

    $expectedFingerprint = ([string]$Run.RunFingerprint).ToLowerInvariant()
    if ([int]$Receipt.ReceiptSchemaVersion -ne 1 -or
        [string]$Receipt.RunId -cne [string]$Run.RunId -or
        [string]$Receipt.RunFingerprint -cne $expectedFingerprint -or
        [string]$Receipt.HostName -cne [string]$Run.HostName -or
        [string]$Receipt.HostName -cne [Environment]::MachineName -or
        [string]$Receipt.ProgramRelease -cne [string]$Run.ProgramRelease -or
        [string]$Receipt.ReleaseIntegrityStatus -cne 'Passed' -or
        [string]$Receipt.ReleaseCommit -cne $ExpectedReleaseCommit -or
        -not (Test-GssResumeSamePath -Left ([string]$Receipt.LiveRunLogPath) -Right $LiveRunLogPath)) {
        throw 'Package-only prior-release recovery receipt does not match the committed run, release-integrity evidence, log, and workstation.'
    }
    if ([string]$Receipt.TransactionStatus -cne 'PackageBlocked' -or
        [string]$Receipt.BackupStatus -cne 'Committed' -or
        -not ($Receipt.PSObject.Properties.Name -contains 'PackagePublished') -or
        [bool]$Receipt.PackagePublished) {
        throw 'Package-only prior-release recovery requires an unpublished PackageBlocked receipt with BackupStatus Committed.'
    }
    if ($null -eq $Receipt.BackupPrepare -or
        $null -eq $Receipt.BackupFinalize -or
        [string]$Receipt.BackupFinalize.Status -cne 'Committed' -or
        [string]$Receipt.BackupFinalize.BackupStatus -cne 'Committed' -or
        [string]$Receipt.BackupFinalize.RunId -cne [string]$Run.RunId -or
        [string]$Receipt.BackupFinalize.Fingerprint -cne $expectedFingerprint) {
        throw 'Package-only prior-release recovery requires the exact committed Drive finalization evidence in its transaction receipt.'
    }
}

function Get-GssResumeCommittedSnapshotEvidence {
    param(
        [Parameter(Mandatory)][object]$Run,
        [Parameter(Mandatory)][object]$Receipt
    )

    $expectedRunId = [string]$Run.RunId
    $expectedFingerprint = ([string]$Run.RunFingerprint).ToLowerInvariant()
    $context = Get-GssDriveBackupRootContext
    $location = Find-GssDriveBackupSnapshot -RootPath $context.RootPath -RunId $expectedRunId
    if ($null -eq $location -or [bool]$location.IsPartial) {
        throw "Package-only prior-release recovery requires an existing committed Drive snapshot for run '$expectedRunId'."
    }

    $validated = Test-GssCommittedBackupSnapshot -SnapshotPath ([string]$location.Path)
    foreach ($document in @($validated.Receipt, $validated.Manifest, $validated.PreparedManifest)) {
        if ([string]$document.run_id -cne $expectedRunId -or
            [string]$document.fingerprint -cne $expectedFingerprint -or
            [string]$document.snapshot_purpose -cne 'WorkbookTransaction') {
            throw 'Committed Drive snapshot identity or purpose does not match the prior-release workbook run.'
        }
    }
    if ([string]$validated.Receipt.status -cne 'Committed' -or
        [string]$validated.Receipt.verification_level -cne 'drivefs_hash_verified' -or
        [string]$validated.Manifest.status -cne 'Committed' -or
        [string]$validated.PreparedManifest.status -cne 'Prepared') {
        throw 'Prior-release Drive evidence is not a hash-verified committed workbook transaction snapshot.'
    }
    foreach ($manifest in @($validated.Manifest, $validated.PreparedManifest)) {
        if ([string]$manifest.host -cne [string]$Run.HostName -or
            [string]$manifest.host -cne [Environment]::MachineName -or
            [string]$manifest.release -cne [string]$Run.ProgramRelease -or
            [string]$manifest.report_week -cne [string]$Run.CurrentWeekEnding -or
            [string]$manifest.drive.verification_level -cne 'drivefs_hash_verified') {
            throw 'Committed Drive manifest release, report week, host, or verification level does not match the prior-release run.'
        }
    }

    $expectedPreparedHash = ([string]$Run.DrivePreparedManifestSha256).ToLowerInvariant()
    $expectedManifestHash = ([string]$Receipt.BackupFinalize.BackupManifestSha256).ToLowerInvariant()
    if ($expectedPreparedHash -notmatch '^[a-f0-9]{64}$' -or
        $expectedManifestHash -notmatch '^[a-f0-9]{64}$' -or
        [string]$validated.PreparedManifestSha256 -cne $expectedPreparedHash -or
        ([string]$Receipt.BackupPrepare.PreparedManifestSha256).ToLowerInvariant() -cne $expectedPreparedHash -or
        [string]$validated.ManifestSha256 -cne $expectedManifestHash) {
        throw 'Committed Drive manifest hashes do not match the live run and transaction receipt.'
    }
    if (-not (Test-GssResumeSamePath -Left ([string]$Receipt.BackupFinalize.SnapshotPath) -Right ([string]$location.Path)) -or
        -not (Test-GssResumeSamePath -Left ([string]$Receipt.BackupFinalize.BackupManifestPath) -Right ([string]$validated.ManifestPath)) -or
        -not (Test-GssResumeSamePath -Left ([string]$Receipt.BackupFinalize.CommitReceiptPath) -Right ([string]$validated.ReceiptPath))) {
        throw 'Committed Drive snapshot paths do not match the transaction receipt.'
    }

    return [pscustomobject]@{
        Status = 'Committed'
        BackupStatus = 'Committed'
        RunId = $expectedRunId
        Fingerprint = $expectedFingerprint
        SnapshotPath = [string]$location.Path
        BackupManifestPath = [string]$validated.ManifestPath
        BackupManifestSha256 = [string]$validated.ManifestSha256
        CommitReceiptPath = [string]$validated.ReceiptPath
        PreparedManifestPath = [string]$validated.PreparedManifestPath
        PreparedManifestSha256 = [string]$validated.PreparedManifestSha256
        VerificationLevel = 'drivefs_hash_verified'
        ValidatedFileCount = [int]$validated.ValidatedFileCount
        ValidatedPreparedFileCount = [int]$validated.ValidatedPreparedFileCount
        PackageOnlyRecovery = $true
        Idempotent = $true
    }
}

$LiveRunLogPath = (Resolve-Path -LiteralPath $LiveRunLogPath).Path
$run = Get-Content -LiteralPath $LiveRunLogPath -Raw | ConvertFrom-Json
if ([string]$run.Mode -ne 'ApplyToMainWorkbook' -or [string]$run.TransactionStatus -ne 'Committed') {
    throw 'ResumeFinalize requires a committed live workbook run log.'
}
if ([string]::IsNullOrWhiteSpace([string]$run.RunId) -or [string]::IsNullOrWhiteSpace([string]$run.RunFingerprint)) {
    throw 'Live run log is missing RunId or RunFingerprint.'
}
if ([string]$run.HostName -cne [Environment]::MachineName) {
    throw "Live run belongs to host '$($run.HostName)', not this workstation."
}
if ((Get-GssResumeRunFingerprint $run) -ne ([string]$run.RunFingerprint).ToLowerInvariant()) {
    throw 'Live run fingerprint no longer matches its immutable transaction evidence.'
}
$resumeReleaseMode = Get-GssResumeReleaseMode `
    -RunRelease ([string]$run.ProgramRelease) `
    -CurrentRelease ([string]$releaseIntegrity.ReleaseTag)

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $Folder = [string]$run.Folder
}
$Folder = (Resolve-Path -LiteralPath $Folder).Path
$allowedLogRoot = [System.IO.Path]::GetFullPath((Join-Path $Folder '_automation_runs\logs')).TrimEnd('\') + '\'
if (-not $LiveRunLogPath.StartsWith($allowedLogRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'Live run log is outside the configured GSS automation log directory.'
}
$DriveBackupScript = Join-Path $scriptRoot 'Invoke-GSS-DriveBackup.ps1'
if (-not (Test-Path -LiteralPath $DriveBackupScript -PathType Leaf)) {
    throw "Drive backup coordinator is unavailable: $DriveBackupScript"
}
$driveBackupLibrary = Join-Path $scriptRoot 'Gss-DriveBackup.ps1'
if ($resumeReleaseMode -eq 'DeclaredPriorReleasePackageRecovery' -and
    -not (Test-Path -LiteralPath $driveBackupLibrary -PathType Leaf)) {
    throw "Drive backup verification library is unavailable: $driveBackupLibrary"
}

foreach ($check in @(
    [pscustomobject]@{ Name = 'live workbook'; Path = [string]$run.TargetWorkbook; Expected = [string]$run.PromotedWorkbookSha256 },
    [pscustomobject]@{ Name = 'staged workbook'; Path = [string]$run.StagedWorkbook; Expected = [string]$run.StagedWorkbookSha256 },
    [pscustomobject]@{ Name = 'staged PDF'; Path = [string]$run.StagedPdf; Expected = [string]$run.StagedPdfSha256 },
    [pscustomobject]@{ Name = 'promoted PDF'; Path = [string]$run.EmailComparisonPdf; Expected = [string]$run.PromotedPdfSha256 },
    [pscustomobject]@{ Name = 'current source'; Path = [string]$run.CurrentSourceWorkbook; Expected = [string]$run.CurrentSourceSha256 },
    [pscustomobject]@{ Name = 'prior-year source'; Path = [string]$run.PriorYearSourceWorkbook; Expected = [string]$run.PriorYearSourceSha256 }
)) {
    if ((Get-GssResumeHash $check.Path) -ne $check.Expected.ToLowerInvariant()) {
        throw "ResumeFinalize blocked because the $($check.Name) no longer matches the committed run."
    }
}

$analyzer = Join-Path $scriptRoot 'Analyze-GSS-Run.ps1'
$stateDirectory = Join-Path $Folder '_automation_runs\state'
$activeRunPath = Join-Path $stateDirectory 'active-transaction.json'
$receiptPath = Join-Path $stateDirectory "transaction-$($run.RunId).json"
$receiptExists = Test-Path -LiteralPath $receiptPath -PathType Leaf
$receipt = if ($receiptExists) {
    Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
}
else {
    [pscustomobject]@{
        ReceiptSchemaVersion = 1
        RunId = [string]$run.RunId
        HostName = [Environment]::MachineName
        ProgramRelease = [string]$run.ProgramRelease
        LiveRunLogPath = $LiveRunLogPath
        AutomaticSending = 'Disabled'
    }
}
$priorReleaseSnapshotEvidence = $null
$priorReleasePackageInputEvidence = $null
if ($resumeReleaseMode -eq 'DeclaredPriorReleasePackageRecovery') {
    if (-not $receiptExists) {
        throw 'Package-only prior-release recovery requires an existing transaction receipt.'
    }
    $priorReleaseCommit = Get-GssResumeAnnotatedReleaseCommit `
        -RepoRoot $repoRoot `
        -ReleaseTag ([string]$run.ProgramRelease)
    Assert-GssResumePriorReleaseReceipt `
        -Run $run `
        -Receipt $receipt `
        -LiveRunLogPath $LiveRunLogPath `
        -ExpectedReleaseCommit $priorReleaseCommit
    . $driveBackupLibrary
    $priorReleaseSnapshotEvidence = Get-GssResumeCommittedSnapshotEvidence -Run $run -Receipt $receipt
    $trustedCommittedManifest = Get-GssResumeTrustedBackupManifest -SnapshotEvidence $priorReleaseSnapshotEvidence
    $priorReleasePackageInputEvidence = Assert-GssResumePackageInputSet `
        -FolderPath $Folder `
        -LiveRunLogPath $LiveRunLogPath `
        -Run $run `
        -CommittedManifest $trustedCommittedManifest
    $priorReleaseHistoricalRecoveryEvidence = Assert-GssResumeHistoricalRecoveryEvidence -FolderPath $Folder
    $priorReleasePackageInputEvidence | Add-Member `
        -NotePropertyName HistoricalRecoveryEvidence `
        -NotePropertyValue $priorReleaseHistoricalRecoveryEvidence
    $priorReleaseSnapshotEvidence | Add-Member `
        -NotePropertyName PackageInputEvidence `
        -NotePropertyValue $priorReleasePackageInputEvidence
}

if (Test-Path -LiteralPath $activeRunPath -PathType Leaf) {
    $activePreflight = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
    if ([string]$activePreflight.RunId -ne [string]$run.RunId) {
        throw "A different unresolved GSS transaction is active: $($activePreflight.RunId)"
    }
    if ([string]$activePreflight.TransactionStatus -in @('RollbackAttempting', 'RollbackBlocked')) {
        throw "Transaction '$($run.RunId)' requires rollback review and cannot be resumed as a Drive finalization."
    }
}

$mutex = New-Object System.Threading.Mutex($false, 'Global\GSSSurveyWorkbookAutomationTransaction')
$ownsMutex = $false
$ownsActiveMarker = $false
$driveCommitted = ($null -ne $priorReleaseSnapshotEvidence)
$driveBlocked = $false
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

    if (Test-Path -LiteralPath $activeRunPath -PathType Leaf) {
        $active = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
        if ([string]$active.RunId -ne [string]$run.RunId) {
            throw "A different unresolved GSS transaction is active: $($active.RunId)"
        }
    }
    Add-GssResumeReceiptValue $receipt 'ReceiptPath' $receiptPath
    if ($resumeReleaseMode -eq 'DeclaredPriorReleasePackageRecovery') {
        Add-GssResumeReceiptValue $receipt 'ResumeReleaseMode' 'DeclaredV1.1.8ToV1.1.9PackageOnly'
        Add-GssResumeReceiptValue $receipt 'RecoveryProgramRelease' ([string]$releaseIntegrity.ReleaseTag)
        Add-GssResumeReceiptValue $receipt 'TransactionStatus' 'RetryingPackage'
    }
    else {
        Add-GssResumeReceiptValue $receipt 'TransactionStatus' 'RetryingFinalize'
    }
    Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
    Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
    $ownsActiveMarker = $true
    Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt

    $finalize = if ($resumeReleaseMode -eq 'DeclaredPriorReleasePackageRecovery') {
        $priorReleaseSnapshotEvidence
    }
    else {
        & $DriveBackupScript -Operation RetryFinalize -RunSummaryPath $LiveRunLogPath -OutputObject
    }
    if ([string]$finalize.Status -eq 'Blocked') {
        $driveBlocked = $true
        throw 'Drive retry is blocked by a non-retryable backup-chain or identity conflict. Manual review is required.'
    }
    if ([string]$finalize.Status -ne 'Committed' -or
        [string]$finalize.RunId -ne [string]$run.RunId -or
        [string]$finalize.Fingerprint -ne [string]$run.RunFingerprint) {
        throw "Drive retry did not commit the exact run (status: '$($finalize.Status)')."
    }
    $driveCommitted = $true
    Add-GssResumeReceiptValue $receipt 'BackupStatus' 'Committed'
    if ($resumeReleaseMode -eq 'DeclaredPriorReleasePackageRecovery') {
        Add-GssResumeReceiptValue $receipt 'PackageRecoveryBackupEvidence' $finalize
    }
    else {
        Add-GssResumeReceiptValue $receipt 'BackupFinalize' $finalize
    }
    Add-GssResumeReceiptValue $receipt 'TransactionStatus' 'BackupCommitted'
    Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
    Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
    Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt

    $review = & $analyzer -Folder $Folder -LogPath $LiveRunLogPath -OutputObject
    if ($review.WorkbookStatus -eq 'Blocked') {
        throw 'The committed live workbook no longer passes final review; package publication remains blocked.'
    }
    if ($resumeReleaseMode -eq 'DeclaredPriorReleasePackageRecovery') {
        $trustedCommittedManifest = Get-GssResumeTrustedBackupManifest -SnapshotEvidence $priorReleaseSnapshotEvidence
        $recheckedPackageInputEvidence = Assert-GssResumePackageInputSet `
            -FolderPath $Folder `
            -LiveRunLogPath $LiveRunLogPath `
            -Run $run `
            -CommittedManifest $trustedCommittedManifest
        $recheckedHistoricalRecoveryEvidence = Assert-GssResumeHistoricalRecoveryEvidence -FolderPath $Folder
        if ([string]$recheckedPackageInputEvidence.SourceSetSha256 -cne [string]$priorReleasePackageInputEvidence.SourceSetSha256) {
            throw 'Prior-release package inputs changed during final analysis.'
        }
        if ([string]$recheckedHistoricalRecoveryEvidence.EvidenceSetSha256 -cne [string]$priorReleaseHistoricalRecoveryEvidence.EvidenceSetSha256) {
            throw 'Historical recovery evidence changed during final analysis.'
        }
    }
    $published = if ($resumeReleaseMode -eq 'DeclaredPriorReleasePackageRecovery') {
        & $analyzer `
            -Folder $Folder `
            -LogPath $LiveRunLogPath `
            -OutputObject `
            -PublishEmailPackage `
            -ExpectedPackageInputEvidence $priorReleasePackageInputEvidence
    }
    else {
        & $analyzer -Folder $Folder -LogPath $LiveRunLogPath -OutputObject -PublishEmailPackage
    }
    if ($published.EmailReadiness -ne 'Ready' -or -not $published.EmailPackage) {
        Add-GssResumeReceiptValue $receipt 'TransactionStatus' 'PackageBlocked'
        Add-GssResumeReceiptValue $receipt 'PackagePublished' $false
        Add-GssResumeReceiptValue $receipt 'Error' (@($published.Qa.EmailBlockers) -join '; ')
        Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
        Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
        Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt
        throw 'Drive is committed, but email package publication remains blocked.'
    }

    Add-GssResumeReceiptValue $receipt 'TransactionStatus' 'Committed'
    Add-GssResumeReceiptValue $receipt 'PackagePublished' $true
    Add-GssResumeReceiptValue $receipt 'PackagePath' ([string]$published.EmailPackage.PackagePath)
    Add-GssResumeReceiptValue $receipt 'Error' $null
    Add-GssResumeReceiptValue $receipt 'CompletedUtc' ([datetime]::UtcNow.ToString('o'))
    Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
    Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
    Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt

    Write-Information $(if ($resumeReleaseMode -eq 'DeclaredPriorReleasePackageRecovery') {
        'GSS committed-snapshot recovery and package publication completed.'
    }
    else {
        'GSS Drive finalization and package publication completed.'
    }) -InformationAction Continue
    Write-Information "  Run: $($run.RunId)" -InformationAction Continue
    Write-Information "  Package: $($published.EmailPackage.PackagePath)" -InformationAction Continue
    Write-Information '  Automatic sending: disabled' -InformationAction Continue
    if ($OutputObject) { return $receipt }
}
catch {
    Add-GssResumeReceiptValue $receipt 'TransactionStatus' $(if ($driveCommitted) { 'PackageBlocked' } elseif ($driveBlocked) { 'BackupBlocked' } else { 'PendingFinalize' })
    Add-GssResumeReceiptValue $receipt 'BackupStatus' $(if ($driveCommitted) { 'Committed' } elseif ($driveBlocked) { 'Blocked' } else { 'PendingFinalize' })
    Add-GssResumeReceiptValue $receipt 'Error' $_.Exception.Message
    Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
    try {
        if ($ownsActiveMarker) {
            Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
        }
        Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt
    }
    catch { Write-Verbose "The resume transaction receipt could not be updated: $($_.Exception.Message)" }
    throw
}
finally {
    if ($ownsActiveMarker -and
        $receipt.TransactionStatus -in @('Committed', 'PackageBlocked') -and
        (Test-Path -LiteralPath $activeRunPath -PathType Leaf)) {
        try {
            $active = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
            if ([string]$active.RunId -eq [string]$run.RunId) {
                Remove-Item -LiteralPath $activeRunPath -Force
            }
        }
        catch { Write-Verbose "The resumed active transaction marker could not be removed: $($_.Exception.Message)" }
    }
    if ($ownsMutex) {
        try { $mutex.ReleaseMutex() }
        catch { Write-Verbose "The GSS transaction mutex could not be released cleanly: $($_.Exception.Message)" }
    }
    $mutex.Dispose()
}
