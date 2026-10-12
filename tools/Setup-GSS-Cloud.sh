#!/usr/bin/env bash
# Development-only bootstrap for a disposable Ubuntu x64 Cloud environment.
set -euo pipefail

repo_root="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
readonly powershell_version='7.6.6'
readonly package_sha256='9585f38ab5a026c3fc0995486e26e12050777960fef47a22dca98b577c5d27a7'

if [[ "$(uname -s)" != 'Linux' || "$(uname -m)" != 'x86_64' ]]; then
    echo 'This bootstrap supports Ubuntu x64 development environments only.' >&2
    exit 1
fi
if [[ ! -r /etc/os-release ]]; then
    echo 'Cannot identify the Linux distribution.' >&2
    exit 1
fi
# os-release is provided by the operating system.
source /etc/os-release
if [[ "${ID:-}" != 'ubuntu' || ! "${VERSION_ID:-}" =~ ^(22\.04|24\.04|26\.04)$ ]]; then
    echo 'Use a supported Ubuntu LTS environment: 22.04, 24.04, or 26.04.' >&2
    exit 1
fi

installed_version=''
if command -v pwsh >/dev/null 2>&1; then
    installed_version="$(pwsh -NoLogo -NoProfile -Command '$PSVersionTable.PSVersion.ToString()')"
    if dpkg --compare-versions "$installed_version" gt "$powershell_version"; then
        echo "PowerShell $installed_version is newer than the reviewed pin $powershell_version; choose a clean environment or review a pin update." >&2
        exit 1
    fi
fi
if [[ "$installed_version" != "$powershell_version" ]]; then
    elevate=()
    if [[ "$EUID" -ne 0 ]]; then
        if ! command -v sudo >/dev/null 2>&1 || ! sudo -n true; then
            echo 'Installing the pinned PowerShell package needs root or passwordless sudo inside the Cloud environment.' >&2
            exit 1
        fi
        elevate=(sudo -n)
    fi
    "${elevate[@]}" apt-get update
    "${elevate[@]}" apt-get install -y --no-install-recommends ca-certificates curl
    package_dir="$(mktemp -d -t gss-powershell.XXXXXXXX)"
    package_file="$package_dir/powershell_${powershell_version}-1.deb_amd64.deb"
    # Only remove the named package and its own empty temporary directory.
    trap 'rm -f -- "$package_file"; rmdir -- "$package_dir"' EXIT
    curl --fail --show-error --silent --location --retry 3 \
        "https://github.com/PowerShell/PowerShell/releases/download/v${powershell_version}/powershell_${powershell_version}-1.deb_amd64.deb" \
        --output "$package_file"
    printf '%s  %s\n' "$package_sha256" "$package_file" | sha256sum --check --strict -
    "${elevate[@]}" apt-get install -y --no-install-recommends "$package_file"
fi

actual_version="$(pwsh -NoLogo -NoProfile -Command '$PSVersionTable.PSVersion.ToString()')"
if [[ "$actual_version" != "$powershell_version" ]]; then
    echo "Expected PowerShell $powershell_version; found $actual_version." >&2
    exit 1
fi
pwsh -NoLogo -NoProfile -File "$repo_root/tools/Initialize-GSS-Development.ps1"
echo "GSS development dependencies are ready with PowerShell $actual_version."
echo "Run: pwsh -NoProfile -File $repo_root/tools/Test-GSS-Cloud.ps1"
