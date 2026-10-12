[CmdletBinding()]
param([string]$RepoRoot)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'7.2') {
    throw 'The portable development suite requires PowerShell 7.2 or later. Windows reporting continues to use Windows PowerShell 5.1 and desktop Excel.'
}
if ([string]::IsNullOrWhiteSpace($RepoRoot)) { $RepoRoot = Split-Path -Parent $PSScriptRoot }
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

& (Join-Path $RepoRoot 'tools/Initialize-GSS-Development.ps1') -CheckOnly
Write-Output 'Portable scope: all PowerShell syntax, reviewed static analysis, synthetic reporting/source-choice/workbook-discovery tests, and analytics.'
Write-Output 'Windows transaction, backup, package, historical-recovery, and desktop Excel integration suites are outside this scope and remain required separately.'
Write-Output 'On non-Windows platforms, analytics also excludes its single Windows Dropbox-root workbook-path assertion; all pure analytics assertions run.'

& (Join-Path $RepoRoot 'tools/Test-GSS-PowerShellSyntax.ps1')
& (Join-Path $RepoRoot 'tools/Test-GSS-PSScriptAnalyzer.ps1') -RepoRoot $RepoRoot
& (Join-Path $RepoRoot 'tools/Test-GSS-Pester.ps1') -RepoRoot $RepoRoot -TestFile @(
    'Gss.Reporting.Tests.ps1',
    'Gss.SourceChoices.Tests.ps1',
    'Gss.Paths.Tests.ps1'
)
& (Join-Path $RepoRoot 'tests/Test-GSS-Analytics.ps1')

Write-Output 'Portable synthetic development suite passed. This does not verify desktop Excel, cloud workspace connection, or live reporting.'
