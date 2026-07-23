[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$BaselinePath,
    [switch]$UpdateBaseline
)

$ErrorActionPreference = 'Stop'
$requiredVersion = [version]'1.25.0'

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path
if ([string]::IsNullOrWhiteSpace($BaselinePath)) {
    $BaselinePath = Join-Path $RepoRoot 'config\psscriptanalyzer-baseline.json'
}

$availableModule = Get-Module -ListAvailable PSScriptAnalyzer |
    Where-Object { $_.Version -eq $requiredVersion } |
    Select-Object -First 1
if (-not $availableModule) {
    throw "PSScriptAnalyzer $requiredVersion is required. Install-Module PSScriptAnalyzer -RequiredVersion $requiredVersion -Scope CurrentUser"
}
Import-Module PSScriptAnalyzer -RequiredVersion $requiredVersion -Force

function ConvertTo-AnalyzerPortablePath {
    param([Parameter(Mandatory)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if (-not $fullPath.StartsWith("$RepoRoot\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Analyzer returned a path outside the repository: $Path"
    }
    return $fullPath.Substring($RepoRoot.Length + 1).Replace('\', '/')
}

function Get-AnalyzerFingerprint {
    param([Parameter(Mandatory)][object]$Result)

    $portablePath = ConvertTo-AnalyzerPortablePath -Path ([string]$Result.ScriptPath)
    $normalizedMessage = [regex]::Replace(([string]$Result.Message).Trim(), '\s+', ' ')
    return "$portablePath|$($Result.RuleName)|$normalizedMessage"
}

$results = @(
    Invoke-ScriptAnalyzer -Path $RepoRoot -Recurse -Severity @('Error', 'Warning') |
        Sort-Object ScriptPath, RuleName, Line, Column
)
$currentRawEntries = @(
    $results |
        ForEach-Object {
            [pscustomobject][ordered]@{
                fingerprint = Get-AnalyzerFingerprint -Result $_
                path = ConvertTo-AnalyzerPortablePath -Path ([string]$_.ScriptPath)
                rule = [string]$_.RuleName
                severity = [string]$_.Severity
                message = [regex]::Replace(([string]$_.Message).Trim(), '\s+', ' ')
            }
        }
)
$currentEntries = @(
    $currentRawEntries |
        Group-Object fingerprint |
        ForEach-Object {
            $first = $_.Group | Select-Object -First 1
            [pscustomobject][ordered]@{
                fingerprint = [string]$first.fingerprint
                path = [string]$first.path
                rule = [string]$first.rule
                severity = [string]$first.severity
                message = [string]$first.message
                count = [int]$_.Count
            }
        } |
        Sort-Object fingerprint
)

if ($UpdateBaseline) {
    $baseline = [ordered]@{
        schema_version = 1
        psscriptanalyzer_version = $requiredVersion.ToString()
        generated_at_utc = [datetime]::UtcNow.ToString('o')
        policy = 'Existing warnings are reviewed debt. CI rejects any new warning fingerprint.'
        entries = $currentEntries
    }
    $parent = Split-Path -Parent $BaselinePath
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $temporaryPath = "$BaselinePath.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        $baseline | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $temporaryPath -Encoding UTF8
        Move-Item -LiteralPath $temporaryPath -Destination $BaselinePath -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Output "PSScriptAnalyzer baseline updated: $($currentEntries.Count) reviewed warning fingerprint(s)."
    return
}

if (-not (Test-Path -LiteralPath $BaselinePath -PathType Leaf)) {
    throw "PSScriptAnalyzer baseline not found: $BaselinePath"
}
$baseline = Get-Content -LiteralPath $BaselinePath -Raw | ConvertFrom-Json
if ([int]$baseline.schema_version -ne 1) {
    throw "Unsupported PSScriptAnalyzer baseline schema: $($baseline.schema_version)"
}
if ([version]$baseline.psscriptanalyzer_version -ne $requiredVersion) {
    throw "PSScriptAnalyzer baseline expects $($baseline.psscriptanalyzer_version), not $requiredVersion."
}

$approved = @{}
foreach ($entry in @($baseline.entries)) {
    $approved[[string]$entry.fingerprint] = [int]$entry.count
}
$newResults = @(
    $currentEntries |
        Where-Object {
            -not $approved.ContainsKey([string]$_.fingerprint) -or
            [int]$_.count -gt [int]$approved[[string]$_.fingerprint]
        }
)
if ($newResults.Count -gt 0) {
    $newResults |
        Select-Object path, rule, severity, count, @{ Name = 'approved_count'; Expression = {
            if ($approved.ContainsKey([string]$_.fingerprint)) { [int]$approved[[string]$_.fingerprint] } else { 0 }
        } }, message |
        Format-Table -Wrap -AutoSize
    throw "PSScriptAnalyzer found $($newResults.Count) new or increased warning fingerprint(s). Fix them or explicitly review and refresh the baseline."
}

$current = @{}
foreach ($entry in $currentEntries) {
    $current[[string]$entry.fingerprint] = [int]$entry.count
}
$resolvedCount = 0
foreach ($entry in @($baseline.entries)) {
    $currentCount = if ($current.ContainsKey([string]$entry.fingerprint)) {
        [int]$current[[string]$entry.fingerprint]
    }
    else { 0 }
    $resolvedCount += [math]::Max(0, ([int]$entry.count - $currentCount))
}

Write-Output "PSScriptAnalyzer passed with no new warnings. Reviewed fingerprints: $($approved.Count); reviewed warning instances: $(@($baseline.entries | ForEach-Object { [int]$_.count } | Measure-Object -Sum).Sum); currently resolved instances: $resolvedCount."
