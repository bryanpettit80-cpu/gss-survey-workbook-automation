[CmdletBinding(SupportsShouldProcess = $true)]
param([switch]$CheckOnly)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') {
    throw 'GSS development tools require Windows PowerShell 5.1 or PowerShell 7.2 or later.'
}
if ([IO.Path]::DirectorySeparatorChar -ne '\' -and $PSVersionTable.PSVersion -lt [version]'7.2') {
    throw 'Linux development requires PowerShell 7.2 or later. Install PowerShell before running this script.'
}

$requirements = @(
    [pscustomobject]@{ Name = 'Pester'; Version = [version]'5.8.0' },
    [pscustomobject]@{ Name = 'PSScriptAnalyzer'; Version = [version]'1.25.0' }
)
foreach ($requirement in $requirements) {
    $available = @(Get-Module -ListAvailable -Name $requirement.Name | Where-Object Version -eq $requirement.Version)
    if ($available.Count -eq 0) {
        if ($CheckOnly) {
            throw "$($requirement.Name) $($requirement.Version) is missing. Run tools/Initialize-GSS-Development.ps1 during environment setup."
        }
        if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
            throw 'Install-Module is unavailable. Install PowerShellGet before initializing development modules.'
        }
        if ($PSCmdlet.ShouldProcess("Current-user PowerShell modules: $($requirement.Name) $($requirement.Version)", 'Install pinned development dependency from PSGallery')) {
            $install = @{
                Name = $requirement.Name
                RequiredVersion = $requirement.Version.ToString()
                Repository = 'PSGallery'
                Scope = 'CurrentUser'
                Force = $true
                AllowClobber = $true
                ErrorAction = 'Stop'
            }
            if ($requirement.Name -eq 'Pester') { $install.SkipPublisherCheck = $true }
            Install-Module @install
        }
        elseif ($WhatIfPreference) { continue }
    }
    $available = @(Get-Module -ListAvailable -Name $requirement.Name | Where-Object Version -eq $requirement.Version)
    if ($available.Count -eq 0) { throw "Required development module remains unavailable: $($requirement.Name) $($requirement.Version)." }
    Import-Module -Name $requirement.Name -RequiredVersion $requirement.Version -Force -ErrorAction Stop
    Write-Output "$($requirement.Name) $($requirement.Version): available."
}

Write-Output 'Development dependencies checked. No GSS runtime settings, schedules, reports, or mail were changed.'
