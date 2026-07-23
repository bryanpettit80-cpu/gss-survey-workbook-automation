[CmdletBinding()]
param(
    [string]$RepoRoot,
    [string]$ManifestPath,
    [switch]$SkipTagCheck,
    [switch]$AllowDirtyTree
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
}

Import-Module (Join-Path (Split-Path -Parent $PSCommandPath) 'Gss-ReleaseIntegrity.psm1') -Force

$arguments = @{
    RepoRoot = $RepoRoot
    SkipTagCheck = $SkipTagCheck
    AllowDirtyTree = $AllowDirtyTree
}
if (-not [string]::IsNullOrWhiteSpace($ManifestPath)) {
    $arguments.ManifestPath = $ManifestPath
}

Test-GssReleaseIntegrity @arguments
