[CmdletBinding()]
param(
    [string]$RepoRoot
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
$configuration.Run.Path = Join-Path $RepoRoot 'tests'
$configuration.Run.PassThru = $true
$configuration.Output.Verbosity = 'Detailed'
$configuration.TestResult.Enabled = $true
$configuration.TestResult.OutputPath = Join-Path $env:TEMP 'gss-pester-results.xml'
$configuration.TestResult.OutputFormat = 'NUnitXml'

$result = Invoke-Pester -Configuration $configuration
if ($result.FailedCount -gt 0 -or $result.Result -ne 'Passed') {
    throw "Pester failed: $($result.FailedCount) failed test(s)."
}

Write-Output "Pester passed: $($result.PassedCount) test(s)."
