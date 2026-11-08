#!/usr/bin/env bash
# Exercise the bootstrap with synthetic host/package commands; never install.
set -euo pipefail
repo_root="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
fixture="$(mktemp -d -t gss-bootstrap-tests.XXXXXXXX)"
trap 'rm -f -- "$fixture"/*; rmdir -- "$fixture"' EXIT
printf 'ID=ubuntu\nVERSION_ID=24.04\n' > "$fixture/os-release"
# Redirect only the OS identity file in a disposable copy. Runtime logic and
# package/version/checksum values remain the production script's own values.
sed "s#/etc/os-release#$fixture/os-release#g" \
    "$repo_root/tools/Setup-GSS-Cloud.sh" > "$fixture/setup.sh"
cat > "$fixture/commands.sh" <<'MOCKS'
set -euo pipefail
uname() { case "$1" in -s) echo Linux ;; -m) echo x86_64 ;; *) return 2 ;; esac; }
pwsh() {
    if [[ " $* " == *' -Command '* ]]; then
        cat "$GSS_FIXTURE/version"
    elif [[ " $* " == *' -File '* ]]; then
        echo initialized >> "$GSS_FIXTURE/events"
    else
        return 2
    fi
}
dpkg() {
    [[ "$1" == --compare-versions && "$3" == gt ]]
    [[ "$2" != "$4" && "$(printf '%s\n' "$2" "$4" | sort -V | tail -n 1)" == "$2" ]]
}
sudo() {
    [[ "$1" == -n ]]
    shift
    if [[ "$1" == true ]]; then return 0; fi
    "$@"
}
apt-get() {
    printf 'apt %s\n' "$*" >> "$GSS_FIXTURE/events"
    if [[ "$*" == *'.deb' ]]; then
        # The checksum verifier must complete before the install command.
        grep -qx verified "$GSS_FIXTURE/events"
        [[ " $* " == *' --no-remove '* ]]
        printf '%s\n' "${GSS_RESULT_VERSION:-7.6.6}" > "$GSS_FIXTURE/version"
    fi
}
curl() {
    printf 'download %s\n' "$*" >> "$GSS_FIXTURE/events"
    [[ "$*" == *'https://github.com/PowerShell/PowerShell/releases/download/v7.6.6/powershell_7.6.6-1.deb_amd64.deb'* ]]
    while [[ "$1" != --output ]]; do shift; done
    printf synthetic-package > "$2"
}
sha256sum() {
    [[ "$*" == '--check --strict -' ]]
    local evidence
    IFS= read -r evidence
    [[ "$evidence" == '9585f38ab5a026c3fc0995486e26e12050777960fef47a22dca98b577c5d27a7  '* ]]
    if [[ "${GSS_HASH_FAILURE:-false}" == true ]]; then return 1; fi
    echo verified >> "$GSS_FIXTURE/events"
}
script="$1"
shift
source "$script" "$@"
MOCKS

export GSS_FIXTURE="$fixture"
run_case() {
    local expected="$1" version="$2"
    shift 2
    printf '%s\n' "$version" > "$fixture/version"
    : > "$fixture/events"
    local status=0
    bash "$fixture/commands.sh" "$fixture/setup.sh" "$@" > "$fixture/output" 2>&1 || status=$?
    if [[ "$status" != "$expected" ]]; then
        cat "$fixture/output" >&2
        echo "Expected exit $expected; found $status for version $version ($*)." >&2
        exit 1
    fi
}
assert_no_package_install() {
    if grep -q '^apt .*\.deb$' "$fixture/events"; then
        echo 'Unexpected PowerShell package installation.' >&2
        exit 1
    fi
}
assert_no_initialization() {
    if grep -qx initialized "$fixture/events"; then
        echo 'Development initialization ran after a failed prerequisite.' >&2
        exit 1
    fi
}

run_case 0 7.6.6
[[ "$(cat "$fixture/events")" == initialized ]]
echo 'PASS: matching pin is reused without package operations'

run_case 1 7.6.9
[[ ! -s "$fixture/events" ]]
grep -q 'newer than the reviewed pin' "$fixture/output"
echo 'PASS: newer version is refused without explicit opt-in'

run_case 0 7.6.9 --replace-newer-powershell
grep -q '^apt .*--allow-downgrades .*powershell_7.6.6-1.deb_amd64.deb$' "$fixture/events"
grep -qx initialized "$fixture/events"
echo 'PASS: opted-in replacement verifies the pinned package before installation'

run_case 0 7.6.5
grep -q '^apt .*powershell_7.6.6-1.deb_amd64.deb$' "$fixture/events"
if grep -q -- --allow-downgrades "$fixture/events"; then exit 1; fi
echo 'PASS: older version upgrades without downgrade permission'

export GSS_HASH_FAILURE=true
run_case 1 7.6.9 --replace-newer-powershell
assert_no_package_install
assert_no_initialization
unset GSS_HASH_FAILURE
echo 'PASS: checksum failure prevents package installation and initialization'

export GSS_RESULT_VERSION=7.6.9
run_case 1 7.6.9 --replace-newer-powershell
assert_no_initialization
grep -q 'Expected PowerShell 7.6.6; found 7.6.9' "$fixture/output"
unset GSS_RESULT_VERSION
echo 'PASS: incorrect post-install runtime prevents initialization'

run_case 2 7.6.9 --unknown-option
[[ ! -s "$fixture/events" ]]
echo 'PASS: unknown options fail before host changes'
echo 'All 7 synthetic bootstrap scenarios passed.'
