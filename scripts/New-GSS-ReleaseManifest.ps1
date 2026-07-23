[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version,
    [string]$Tag,
    [string]$RepoRoot,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
}

Import-Module (Join-Path (Split-Path -Parent $PSCommandPath) 'Gss-ReleaseIntegrity.psm1') -Force

$arguments = @{
    RepoRoot = $RepoRoot
    Version = $Version
}
if (-not [string]::IsNullOrWhiteSpace($Tag)) {
    $arguments.Tag = $Tag
}
if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $arguments.OutputPath = $OutputPath
}

New-GssReleaseManifest @arguments
