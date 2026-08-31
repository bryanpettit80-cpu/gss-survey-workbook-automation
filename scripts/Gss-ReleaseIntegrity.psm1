Set-StrictMode -Version Latest

$script:ReleaseManifestRelativePath = 'release/release-manifest.json'
$script:ReleaseExtensions = @(
    '.cmd',
    '.json',
    '.md',
    '.ps1',
    '.psd1',
    '.psm1',
    '.py',
    '.lock',
    '.txt',
    '.yml',
    '.yaml'
)
$script:ExecutableExtensions = @(
    '.bat',
    '.cmd',
    '.com',
    '.dll',
    '.exe',
    '.js',
    '.jse',
    '.msi',
    '.msp',
    '.ps1',
    '.psd1',
    '.psm1',
    '.py',
    '.scr',
    '.vbs',
    '.wsf',
    '.wsh'
)

function ConvertTo-GssReleasePath {
    param([Parameter(Mandatory)][string]$Path)

    return $Path.Replace('\', '/').TrimStart('/')
}

function Test-GssReleaseTrackedPath {
    param([Parameter(Mandatory)][string]$Path)

    $portablePath = ConvertTo-GssReleasePath -Path $Path
    if ($portablePath.Equals($script:ReleaseManifestRelativePath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    $extension = [System.IO.Path]::GetExtension($portablePath)
    return $script:ReleaseExtensions -contains $extension.ToLowerInvariant()
}

function Get-GssGitOutput {
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][string[]]$Arguments
    )

    $output = @(& git -C $RepoRoot @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "Git command failed: git -C `"$RepoRoot`" $($Arguments -join ' ')`n$($output -join [Environment]::NewLine)"
    }
    return @($output | ForEach-Object { [string]$_ })
}

function Get-GssReleaseTrackedFile {
    param([Parameter(Mandatory)][string]$RepoRoot)

    $paths = @(Get-GssGitOutput -RepoRoot $RepoRoot -Arguments @('ls-files'))
    return @(
        $paths |
            ForEach-Object { ConvertTo-GssReleasePath -Path $_ } |
            Where-Object { Test-GssReleaseTrackedPath -Path $_ } |
            Sort-Object -Unique
    )
}

function Get-GssReleaseRole {
    param([Parameter(Mandatory)][string]$Path)

    if ($Path -like 'scripts/Test-*' -or $Path -like 'tests/*') { return 'validation' }
    if ($Path -like 'scripts/*') { return 'runtime' }
    if ($Path -like 'requirements*.lock') { return 'runtime-dependency' }
    if ($Path -like 'templates/*') { return 'operator-template' }
    if ($Path -like 'config/*') { return 'policy' }
    if ($Path -like '.github/*') { return 'ci' }
    if ($Path -like 'release/*') { return 'release-control' }
    return 'documentation'
}

function Get-GssReleaseFileEvidence {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Release file is missing: $Path"
    }
    $text = [System.IO.File]::ReadAllText($Path)
    $normalizedText = $text.Replace("`r`n", "`n").Replace("`r", "`n")
    $encoding = New-Object System.Text.UTF8Encoding($false)
    $bytes = $encoding.GetBytes($normalizedText)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha256.ComputeHash($bytes)
    }
    finally {
        $sha256.Dispose()
    }
    $hash = -join ($hashBytes | ForEach-Object { $_.ToString('x2') })
    return [pscustomobject]@{
        HashMode = 'utf8_lf'
        CanonicalSizeBytes = [long]$bytes.Length
        Sha256 = $hash
    }
}

function Write-GssReleaseAtomicJson {
    param(
        [Parameter(Mandatory)][object]$InputObject,
        [Parameter(Mandatory)][string]$Path,
        [int]$Depth = 12
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $temporaryPath = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        $InputObject | ConvertTo-Json -Depth $Depth | Set-Content -LiteralPath $temporaryPath -Encoding UTF8
        Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-GssReleaseManifest {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version,
        [string]$Tag = "v$Version",
        [string]$OutputPath
    )

    $repo = (Resolve-Path -LiteralPath $RepoRoot).Path
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path $repo ($script:ReleaseManifestRelativePath.Replace('/', '\'))
    }

    $files = @(
        foreach ($relativePath in @(Get-GssReleaseTrackedFile -RepoRoot $repo)) {
            $absolutePath = Join-Path $repo ($relativePath.Replace('/', '\'))
            if (-not (Test-Path -LiteralPath $absolutePath -PathType Leaf)) {
                throw "Tracked release file is missing: $relativePath"
            }
            $evidence = Get-GssReleaseFileEvidence -Path $absolutePath
            [ordered]@{
                path = $relativePath
                role = Get-GssReleaseRole -Path $relativePath
                hash_mode = $evidence.HashMode
                canonical_size_bytes = $evidence.CanonicalSizeBytes
                sha256 = $evidence.Sha256
            }
        }
    )

    $manifest = [ordered]@{
        schema_version = 1
        release_version = $Version
        release_tag = $Tag
        commit_binding = 'exact_release_tag'
        generated_at_utc = [datetime]::UtcNow.ToString('o')
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
        archive_name = "gss-survey-workbook-automation-$Tag.zip"
        files = $files
    }

    if ($PSCmdlet.ShouldProcess($OutputPath, "Write release manifest for $Tag")) {
        Write-GssReleaseAtomicJson -InputObject $manifest -Path $OutputPath
    }
    else {
        return
    }
    return [pscustomobject]@{
        ManifestPath = $OutputPath
        Version = $Version
        Tag = $Tag
        FileCount = $files.Count
        ManifestSha256 = (Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    }
}

function Test-GssReleaseManifestFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Manifest,
        [Parameter(Mandatory)][string]$RepoRoot
    )

    $failures = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($entry in @($Manifest.files)) {
        $portablePath = ConvertTo-GssReleasePath -Path ([string]$entry.path)
        if (-not (Test-GssReleaseTrackedPath -Path $portablePath)) {
            $failures.Add("Manifest contains an unsupported or self-referential path: $portablePath")
            continue
        }
        if ($seen.ContainsKey($portablePath.ToLowerInvariant())) {
            $failures.Add("Manifest contains a duplicate path: $portablePath")
            continue
        }
        $seen[$portablePath.ToLowerInvariant()] = $true

        $absolutePath = Join-Path $RepoRoot ($portablePath.Replace('/', '\'))
        if (-not (Test-Path -LiteralPath $absolutePath -PathType Leaf)) {
            $failures.Add("Release file is missing: $portablePath")
            continue
        }

        if ([string]$entry.hash_mode -ne 'utf8_lf') {
            $failures.Add("Release file has an unsupported hash mode: $portablePath")
            continue
        }
        $evidence = Get-GssReleaseFileEvidence -Path $absolutePath
        if ([long]$entry.canonical_size_bytes -ne $evidence.CanonicalSizeBytes) {
            $failures.Add("Release file canonical size changed: $portablePath")
        }
        if (-not $evidence.Sha256.Equals(([string]$entry.sha256).ToLowerInvariant(), [System.StringComparison]::Ordinal)) {
            $failures.Add("Release file hash changed: $portablePath")
        }
    }

    return @($failures)
}

function Test-GssExcelValidationReceipt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Receipt,
        [Parameter(Mandatory)][string]$ExpectedHead,
        [Parameter(Mandatory)][string]$ExpectedTag,
        [Parameter(Mandatory)][string]$DataRoot
    )

    $failures = New-Object System.Collections.Generic.List[string]
    $portableWorkbookPath = ([string]$Receipt.WorkbookPath).Replace('\', '/')
    $portableSourceLogPath = ([string]$Receipt.SourceRunLogPath).Replace('\', '/')
    if ([int]$Receipt.ReceiptSchemaVersion -ne 1) { $failures.Add('Receipt schema is not supported.') }
    if ([string]$Receipt.Status -ne 'Passed') { $failures.Add('Receipt status is not Passed.') }
    if ([string]$Receipt.GitHead -ne $ExpectedHead) { $failures.Add('Receipt Git HEAD does not match the release.') }
    if ([string]$Receipt.ReleaseTag -ne $ExpectedTag) { $failures.Add('Receipt tag does not match the release.') }
    if ([string]::IsNullOrWhiteSpace([string]$Receipt.ExcelVersion)) { $failures.Add('Receipt has no desktop Excel version.') }
    if ([string]$Receipt.WorkbookSha256 -notmatch '^[a-fA-F0-9]{64}$') { $failures.Add('Receipt workbook hash is invalid.') }
    if ([string]$Receipt.SourceRunFingerprint -notmatch '^[a-fA-F0-9]{64}$') { $failures.Add('Receipt run fingerprint is invalid.') }
    if ([string]::IsNullOrWhiteSpace([string]$Receipt.SourceRunHostName)) { $failures.Add('Receipt has no certifying workstation.') }
    if ($portableWorkbookPath -notlike '_automation_runs/test-output/*') { $failures.Add('Receipt workbook is not an isolated copy-test artifact.') }
    if ($portableSourceLogPath -notlike '_automation_runs/logs/*') { $failures.Add('Receipt source log is not in the GSS automation log directory.') }
    if ([int]$Receipt.FormulaErrors -ne 0) { $failures.Add('Receipt reports formula errors.') }
    if ([int]$Receipt.ConstantErrors -ne 0) { $failures.Add('Receipt reports constant errors.') }

    try {
        $resolvedDataRoot = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $DataRoot).Path).TrimEnd('\', '/')
        foreach ($portablePath in @($portableWorkbookPath, $portableSourceLogPath)) {
            if ([string]::IsNullOrWhiteSpace($portablePath) -or
                [System.IO.Path]::IsPathRooted($portablePath) -or
                @($portablePath -split '/' | Where-Object { $_ -eq '..' }).Count -gt 0) {
                throw "Release evidence path is not a safe portable relative path: $portablePath"
            }
        }
        $workbookPath = [System.IO.Path]::GetFullPath((Join-Path $resolvedDataRoot $portableWorkbookPath.Replace('/', '\')))
        $sourceLogPath = [System.IO.Path]::GetFullPath((Join-Path $resolvedDataRoot $portableSourceLogPath.Replace('/', '\')))
        foreach ($resolvedPath in @($workbookPath, $sourceLogPath)) {
            if (-not $resolvedPath.StartsWith("$resolvedDataRoot\", [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Release evidence escaped the GSS data root: $resolvedPath"
            }
            if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
                throw "Release evidence file is missing: $resolvedPath"
            }
        }

        $actualWorkbookHash = (Get-FileHash -LiteralPath $workbookPath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actualWorkbookHash -ne ([string]$Receipt.WorkbookSha256).ToLowerInvariant()) {
            $failures.Add('Receipt workbook hash does not match the copy-test artifact.')
        }

        $sourceRun = Get-Content -LiteralPath $sourceLogPath -Raw | ConvertFrom-Json
        if ([string]$sourceRun.Mode -ne 'CopyTestOnly' -or
            [string]$sourceRun.TransactionStatus -ne 'Prepared') {
            $failures.Add('Receipt source log is not a Prepared copy-only run.')
        }
        if ([string]$sourceRun.ProgramRelease -ne $ExpectedTag) {
            $failures.Add('Receipt source run does not identify the exact release tag.')
        }
        if ([string]::IsNullOrWhiteSpace([string]$sourceRun.HostName)) {
            $failures.Add('Receipt source run has no audit workstation.')
        }
        if ([string]$sourceRun.HostName -cne [string]$Receipt.SourceRunHostName) {
            $failures.Add('Receipt certifying workstation does not match its source run.')
        }
        if ([string]$sourceRun.RunFingerprint -ne [string]$Receipt.SourceRunFingerprint) {
            $failures.Add('Receipt run fingerprint does not match its source log.')
        }
        if ([string]$sourceRun.StagedWorkbookRelativePath -ne $portableWorkbookPath -or
            [string]$sourceRun.StagedWorkbookSha256 -ne [string]$Receipt.WorkbookSha256) {
            $failures.Add('Receipt workbook evidence does not match its source run.')
        }
        if ((Get-GssReleaseRunFingerprint -Run $sourceRun) -ne ([string]$sourceRun.RunFingerprint).ToLowerInvariant()) {
            $failures.Add('Receipt source run fingerprint cannot be reproduced from its immutable evidence.')
        }
    }
    catch {
        $failures.Add("Receipt evidence could not be resolved: $($_.Exception.Message)")
    }
    return @($failures)
}

function Get-GssReleaseRunFingerprint {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Run)

    $parts = @(
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
    )
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes(($parts -join "`n"))
        return ([System.BitConverter]::ToString($sha256.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Test-GssReleaseIntegrity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [string]$ManifestPath,
        [switch]$SkipTagCheck,
        [switch]$AllowDirtyTree
    )

    $repo = (Resolve-Path -LiteralPath $RepoRoot).Path
    if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
        $ManifestPath = Join-Path $repo ($script:ReleaseManifestRelativePath.Replace('/', '\'))
    }
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
        throw "Release manifest not found: $ManifestPath"
    }

    $manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
    if ([int]$manifest.schema_version -ne 1) {
        throw "Unsupported release-manifest schema: $($manifest.schema_version)"
    }
    if ([string]::IsNullOrWhiteSpace([string]$manifest.release_tag)) {
        throw 'Release manifest does not identify a release tag.'
    }

    $failures = New-Object System.Collections.Generic.List[string]
    foreach ($failure in @(Test-GssReleaseManifestFile -Manifest $manifest -RepoRoot $repo)) {
        $failures.Add($failure)
    }

    $manifestPaths = @(
        @($manifest.files) |
            ForEach-Object { (ConvertTo-GssReleasePath -Path ([string]$_.path)).ToLowerInvariant() } |
            Sort-Object -Unique
    )
    $trackedPaths = @(
        Get-GssReleaseTrackedFile -RepoRoot $repo |
            ForEach-Object { $_.ToLowerInvariant() } |
            Sort-Object -Unique
    )
    foreach ($path in @($trackedPaths | Where-Object { $_ -notin $manifestPaths })) {
        $failures.Add("Tracked release file is absent from the manifest: $path")
    }
    foreach ($path in @($manifestPaths | Where-Object { $_ -notin $trackedPaths })) {
        $failures.Add("Manifest path is not tracked by Git: $path")
    }

    $untrackedExecutables = @(
        Get-GssGitOutput -RepoRoot $repo -Arguments @('ls-files', '--others', '--exclude-standard') |
            Where-Object {
                $extension = [System.IO.Path]::GetExtension($_).ToLowerInvariant()
                $script:ExecutableExtensions -contains $extension
            }
    )
    foreach ($path in $untrackedExecutables) {
        $failures.Add("Untracked executable file is not allowed: $path")
    }

    $statusLines = @(Get-GssGitOutput -RepoRoot $repo -Arguments @('status', '--porcelain', '--untracked-files=all'))
    if (-not $AllowDirtyTree -and $statusLines.Count -gt 0) {
        $failures.Add("Git working tree is not clean ($($statusLines.Count) status line(s)).")
    }

    $head = @(Get-GssGitOutput -RepoRoot $repo -Arguments @('rev-parse', 'HEAD'))[0].Trim()
    $tagCommit = $null
    $excelReceiptPath = $null
    $excelReceiptSha256 = $null
    if (-not $SkipTagCheck) {
        try {
            $tagCommit = @(Get-GssGitOutput -RepoRoot $repo -Arguments @('rev-list', '-n', '1', "refs/tags/$($manifest.release_tag)"))[0].Trim()
        }
        catch {
            $failures.Add("Required release tag is missing: $($manifest.release_tag)")
        }
        if ($tagCommit -and -not $head.Equals($tagCommit, [System.StringComparison]::OrdinalIgnoreCase)) {
            $failures.Add("HEAD $head is not the exact commit tagged $($manifest.release_tag) ($tagCommit).")
        }

        $receiptRelativePath = [string]$manifest.runtime_contract.excel_validation_receipt_relative_path
        if (-not [bool]$manifest.runtime_contract.excel_validation_receipt_required -or
            $receiptRelativePath -ne '_automation_runs/state/release/local-excel-validation-receipt.json') {
            $failures.Add('Release manifest does not require the exact local Excel validation receipt path.')
        }
        else {
            $dataRoot = Split-Path -Parent $repo
            $excelReceiptPath = Join-Path $dataRoot $receiptRelativePath.Replace('/', '\')
            if (-not (Test-Path -LiteralPath $excelReceiptPath -PathType Leaf)) {
                $failures.Add("Required local Excel validation receipt is missing: $excelReceiptPath")
            }
            else {
                try {
                    $excelReceipt = Get-Content -LiteralPath $excelReceiptPath -Raw | ConvertFrom-Json
                    $receiptFailures = @(
                        Test-GssExcelValidationReceipt `
                            -Receipt $excelReceipt `
                            -ExpectedHead $head `
                            -ExpectedTag ([string]$manifest.release_tag) `
                            -DataRoot $dataRoot
                    )
                    if ($receiptFailures.Count -gt 0) {
                        foreach ($receiptFailure in $receiptFailures) {
                            $failures.Add("Local Excel validation receipt failed: $receiptFailure")
                        }
                    }
                    else {
                        $excelReceiptSha256 = (Get-FileHash -LiteralPath $excelReceiptPath -Algorithm SHA256).Hash.ToLowerInvariant()
                    }
                }
                catch {
                    $failures.Add("Local Excel validation receipt could not be verified: $($_.Exception.Message)")
                }
            }
        }
    }

    if ($failures.Count -gt 0) {
        throw "GSS release integrity check failed:`n- $($failures -join "`n- ")"
    }

    return [pscustomobject]@{
        Status = 'Passed'
        ReleaseVersion = [string]$manifest.release_version
        ReleaseTag = [string]$manifest.release_tag
        HeadCommit = $head
        TagCommit = $tagCommit
        ManifestPath = $ManifestPath
        ManifestSha256 = (Get-FileHash -LiteralPath $ManifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
        VerifiedFileCount = @($manifest.files).Count
        CleanTreeRequired = -not [bool]$AllowDirtyTree
        TagCheckSkipped = [bool]$SkipTagCheck
        ExcelValidationReceiptPath = $excelReceiptPath
        ExcelValidationReceiptSha256 = $excelReceiptSha256
    }
}

Export-ModuleMember -Function @(
    'ConvertTo-GssReleasePath',
    'Get-GssReleaseFileEvidence',
    'Get-GssReleaseRunFingerprint',
    'Get-GssReleaseTrackedFile',
    'New-GssReleaseManifest',
    'Test-GssExcelValidationReceipt',
    'Test-GssReleaseIntegrity',
    'Test-GssReleaseManifestFile',
    'Test-GssReleaseTrackedPath'
)
