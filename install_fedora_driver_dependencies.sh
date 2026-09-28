#!/usr/bin/env bash
set -euo pipefail

fail() {
    echo "ERROR: $*" >&2
    exit 1
}

[[ "$(uname -s)" == "Linux" ]] || fail "This installer only runs on Linux."
command -v dnf >/dev/null 2>&1 || fail "dnf was not found. This dependency installer is for Fedora."

kernel_release="$(uname -r)"
kernel_major="${kernel_release%%.*}"

if [[ "$kernel_major" != "6" ]]; then
    echo "WARNING: Moxa's bundled UPort driver supports Linux kernel 6.x; the running kernel is $kernel_release." >&2
fi

if [[ ${EUID:-$(id -u)} -eq 0 ]]; then
    dnf_command=(dnf)
else
    command -v sudo >/dev/null 2>&1 || fail "sudo is required when this script is not run as root."
    dnf_command=(sudo dnf)
fi

echo "Installing Fedora driver build dependencies for kernel $kernel_release..."
"${dnf_command[@]}" install -y \
    gcc \
    make \
    setserial \
    tar \
    coreutils \
    curl \
    elfutils-libelf-devel \
    openssl-devel

if ! "${dnf_command[@]}" install -y "kernel-devel-uname-r == ${kernel_release}"; then
    kernel_arch="$(uname -m)"
    kernel_nvr="${kernel_release%.${kernel_arch}}"
    kernel_version="${kernel_nvr%%-*}"
    kernel_package_release="${kernel_nvr#*-}"

    [[ "$kernel_nvr" != "$kernel_release" && "$kernel_package_release" != "$kernel_nvr" ]] || \
        fail "Could not derive a Fedora kernel-devel package name from $kernel_release."

    archive_url="https://kojipkgs.fedoraproject.org/packages/kernel/${kernel_version}/${kernel_package_release}/${kernel_arch}/kernel-devel-${kernel_release}.rpm"
    archive_dir="$(mktemp -d)"
    trap 'rm -rf -- "$archive_dir"' EXIT
    archive_rpm="${archive_dir}/kernel-devel-${kernel_release}.rpm"

    echo "The exact package is no longer in the active Fedora repositories."
    echo "Trying Fedora's official Koji archive instead; the system kernel will not be upgraded."
    if ! curl --fail --location --proto '=https' --output "$archive_rpm" "$archive_url"; then
        cat >&2 <<EOF
ERROR: Fedora's archive does not contain kernel-devel for $kernel_release.
Do not upgrade to kernel 7 solely for this Moxa driver; Moxa v6.0 supports
kernel 6.x only. Keep the installed 6.x kernel and obtain its exact
kernel-devel RPM, or obtain a kernel 7-compatible driver from Moxa.
Archive URL checked:
  $archive_url
EOF
        exit 1
    fi

    command -v rpmkeys >/dev/null 2>&1 || fail "rpmkeys is required to verify the archived Fedora package."
    rpmkeys --checksig "$archive_rpm" | grep -q "signatures OK" || \
        fail "Fedora signature verification failed for the archived kernel-devel package."
    "${dnf_command[@]}" install -y "$archive_rpm" || \
        fail "The archived kernel-devel package was verified but could not be installed."
fi

for required_command in gcc make setserial tar sha512sum; do
    command -v "$required_command" >/dev/null 2>&1 || fail "$required_command is still unavailable after package installation."
done

kernel_build_file="/lib/modules/${kernel_release}/build/Makefile"
[[ -f "$kernel_build_file" ]] || fail "The matching kernel build files were not created at $kernel_build_file."

echo "Fedora driver dependencies are ready for kernel $kernel_release."
