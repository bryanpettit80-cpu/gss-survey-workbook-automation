[CmdletBinding()]
param(
    [string]$OutputDirectory,

    [string]$PythonPath,

    [string]$CycleLedgerPath,

    [switch]$ComputeSourceHash,

    [AllowEmptyString()]
    [string]$InputJson
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
$gssRoot = Split-Path -Parent $repoRoot
$modelScript = Join-Path $scriptRoot 'gss_shadow_model.py'
$managedPythonPath = Join-Path (
    Join-Path $gssRoot '_automation_runs\runtime\shadow-model-py312'
) 'Scripts\python.exe'
$defaultCycleLedgerPath = Join-Path (
    Join-Path $gssRoot '_automation_runs\state'
) 'shadow-model-cycle-ledger.json'
$process = $null
$rowJson = $null
$fullOutputDirectory = $null
try {
    if (-not (Test-Path -LiteralPath $modelScript -PathType Leaf)) {
        throw "GSS shadow model script not found: $modelScript"
    }

    $rowJson = if ([string]::IsNullOrWhiteSpace($InputJson)) {
        [Console]::In.ReadToEnd()
    }
    else {
        $InputJson
    }
    if ([string]::IsNullOrWhiteSpace($rowJson)) {
        throw 'A gss-model-input/v1 JSON document is required on standard input.'
    }

    $fullRepoRoot = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $repoPrefix = $fullRepoRoot + [System.IO.Path]::DirectorySeparatorChar
    if ([string]::IsNullOrWhiteSpace($CycleLedgerPath)) {
        $CycleLedgerPath = $defaultCycleLedgerPath
    }
    $CycleLedgerPath = [System.IO.Path]::GetFullPath($CycleLedgerPath)
    if ($CycleLedgerPath.Equals(
            $fullRepoRoot,
            [System.StringComparison]::OrdinalIgnoreCase
        ) -or $CycleLedgerPath.StartsWith(
            $repoPrefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
        throw 'The shadow-cycle ledger must be stored outside the program repository.'
    }
    if (-not $ComputeSourceHash) {
        if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
            throw 'OutputDirectory is required unless ComputeSourceHash is selected.'
        }
        $fullOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
        if ($fullOutputDirectory.Equals(
                $fullRepoRoot,
                [System.StringComparison]::OrdinalIgnoreCase
            ) -or $fullOutputDirectory.StartsWith(
                $repoPrefix,
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
            throw 'Model artifacts must be written outside the program repository.'
        }
        if ($fullOutputDirectory.Contains('"')) {
            throw 'Model output paths cannot contain a double quote.'
        }
        New-Item -ItemType Directory -Path $fullOutputDirectory -Force | Out-Null
    }

    if ($modelScript.Contains('"') -or $CycleLedgerPath.Contains('"')) {
        throw 'Model script and state paths cannot contain a double quote.'
    }

    $pythonExecutable = $PythonPath
    if ([string]::IsNullOrWhiteSpace($pythonExecutable)) {
        if (Test-Path -LiteralPath $managedPythonPath -PathType Leaf) {
            $pythonExecutable = $managedPythonPath
        }
        else {
            throw (
                'The managed GSS Python runtime is required and was not found at {0}. ' +
                'Install requirements-shadow-model.lock into that runtime before modeling.'
            ) -f $managedPythonPath
        }
    }
    if ($pythonExecutable.Contains('"')) {
        throw 'PythonPath cannot contain a double quote.'
    }

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $pythonExecutable
    $startInfo.Arguments = if ($ComputeSourceHash) {
        '"{0}" --compute-source-sha256' -f $modelScript
    }
    else {
        '"{0}" --output-dir "{1}" --cycle-ledger "{2}"' -f (
            $modelScript,
            $fullOutputDirectory,
            $CycleLedgerPath
        )
    }
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.EnvironmentVariables['PYTHONHASHSEED'] = '0'

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    if (-not $process.Start()) {
        throw 'Failed to start the managed Python modeling process.'
    }

    $process.StandardInput.Write($rowJson)
    $process.StandardInput.Close()
    $standardOutput = $process.StandardOutput.ReadToEnd()
    $standardError = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    if ($process.ExitCode -ne 0) {
        $safeDetail = if ([string]::IsNullOrWhiteSpace($standardError)) {
            'No runtime detail was returned.'
        }
        else {
            $standardError.Trim()
        }
        throw "GSS shadow modeling failed with exit code $($process.ExitCode). $safeDetail"
    }

    if ([string]::IsNullOrWhiteSpace($standardOutput)) {
        throw 'GSS shadow modeling returned no aggregate status.'
    }

    if ($ComputeSourceHash) {
        $sourceSha256 = $standardOutput.Trim().ToLowerInvariant()
        if ($sourceSha256 -notmatch '^[0-9a-f]{64}$') {
            throw 'GSS shadow modeling returned an invalid canonical source SHA-256.'
        }
        [pscustomobject][ordered]@{
            Status = 'HashComputed'
            SourceSha256 = $sourceSha256
            SummaryPath = $null
            EstimatesPath = $null
            DiagnosticsPath = $null
            InputManifestPath = $null
            ModelCardPath = $null
            CycleLedgerPath = $CycleLedgerPath
            TechnicalError = $null
        }
    }
    else {
        $aggregateStatus = $standardOutput.Trim() | ConvertFrom-Json
        [pscustomobject][ordered]@{
            Status = [string]$aggregateStatus.status
            SourceSha256 = $null
            SummaryPath = Join-Path $fullOutputDirectory 'model_summary.json'
            EstimatesPath = Join-Path $fullOutputDirectory 'model_estimates.csv'
            DiagnosticsPath = Join-Path $fullOutputDirectory 'model_diagnostics.json'
            InputManifestPath = Join-Path $fullOutputDirectory 'input_manifest.json'
            ModelCardPath = Join-Path $fullOutputDirectory 'model_card.md'
            CycleLedgerPath = $CycleLedgerPath
            TechnicalError = $null
        }
    }
}
catch {
    [pscustomobject][ordered]@{
        Status = 'TechnicalError'
        SourceSha256 = $null
        SummaryPath = $null
        EstimatesPath = $null
        DiagnosticsPath = $null
        InputManifestPath = $null
        ModelCardPath = $null
        CycleLedgerPath = $CycleLedgerPath
        TechnicalError = $_.Exception.Message
    }
}
finally {
    if ($null -ne $process) {
        try {
            if (-not $process.HasExited) {
                $process.Kill()
            }
        }
        catch {
            Write-Verbose (
                'Best-effort GSS modeling process cleanup failed after the primary result was captured.'
            )
        }
        $process.Dispose()
    }
    $rowJson = $null
}
