[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string[]]$TestFile = @()
)

$ErrorActionPreference = 'Stop'
$requiredVersion = [version]'5.8.0'

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

$availableModule = Get-Module -ListAvailable Pester |
    Where-Object { $_.Version -eq $requiredVersion } |
    Select-Object -First 1
if (-not $availableModule) {
    throw "Pester $requiredVersion is required. Install-Module Pester -RequiredVersion $requiredVersion -Scope CurrentUser"
}
Import-Module Pester -RequiredVersion $requiredVersion -Force

$configuration = [PesterConfiguration]::Default
$testRoot = Join-Path $RepoRoot 'tests'
if ($TestFile.Count -gt 0) {
    $configuration.Run.Path = @(
        foreach ($name in $TestFile) {
            if ([IO.Path]::GetFileName($name) -cne $name -or -not $name.EndsWith('.Tests.ps1', [StringComparison]::Ordinal)) {
                throw "TestFile must name a Pester test directly inside tests/: $name"
            }
            $path = Join-Path $testRoot $name
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Pester test file is missing: $path" }
            $path
        }
    )
}
else { $configuration.Run.Path = $testRoot }
$configuration.Run.PassThru = $true
$configuration.Output.Verbosity = 'Detailed'
$configuration.TestResult.Enabled = $true
$configuration.TestResult.OutputPath = Join-Path ([IO.Path]::GetTempPath()) ('gss-pester-' + [guid]::NewGuid().ToString('N') + '.xml')
$configuration.TestResult.OutputFormat = 'NUnitXml'

$result = Invoke-Pester -Configuration $configuration
if ($result.FailedCount -gt 0 -or $result.Result -ne 'Passed' -or $result.PassedCount -lt 1) {
    throw "Pester failed: $($result.FailedCount) failed test(s)."
}
if ($TestFile.Count -gt 0 -and $result.SkippedCount -gt 0) {
    throw "The selected Pester scope contains $($result.SkippedCount) skipped test(s); this is not complete portable coverage."
}

Write-Output "Pester passed: $($result.PassedCount) test(s)."
Write-Output "Pester result: $($configuration.TestResult.OutputPath.Value)"
