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
if ! "${dnf_command[@]}" install -y \
    gcc \
    make \
    setserial \
    tar \
    coreutils \
    curl \
    elfutils-libelf-devel \
    openssl-devel \
    "kernel-devel-uname-r == ${kernel_release}"; then
    cat >&2 <<EOF
ERROR: Fedora could not install the dependencies for the running kernel.
If the exact kernel-devel package is no longer available, run:
  sudo dnf upgrade --refresh
  sudo reboot
Then run this script again after rebooting into the updated kernel.
EOF
    exit 1
fi

for required_command in gcc make setserial tar sha512sum; do
    command -v "$required_command" >/dev/null 2>&1 || fail "$required_command is still unavailable after package installation."
done

kernel_build_file="/lib/modules/${kernel_release}/build/Makefile"
[[ -f "$kernel_build_file" ]] || fail "The matching kernel build files were not created at $kernel_build_file."

echo "Fedora driver dependencies are ready for kernel $kernel_release."
