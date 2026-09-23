#!/usr/bin/env bash
set -euo pipefail

target_user="${SUDO_USER:-${USER:-}}"

if [[ "${EUID}" -ne 0 ]]; then
    if ! command -v sudo >/dev/null 2>&1; then
        echo "Run this installer as root." >&2
        exit 1
    fi
    exec sudo "$0" "$@"
fi

if ! getent group dialout >/dev/null 2>&1; then
    groupadd --system dialout
fi

rules_path="/etc/udev/rules.d/70-serial-loopback-tester.rules"
install -d -m 0755 /etc/udev/rules.d
printf '%s\n' \
    'SUBSYSTEM=="tty", KERNEL=="ttyUSB[0-9]*", GROUP="dialout", MODE="0660", TAG+="uaccess"' \
    'SUBSYSTEM=="tty", KERNEL=="ttyACM[0-9]*", GROUP="dialout", MODE="0660", TAG+="uaccess"' \
    'SUBSYSTEM=="tty", KERNEL=="ttyr[0-9a-fA-F]*", GROUP="dialout", MODE="0660", TAG+="uaccess"' \
    > "$rules_path"
chmod 0644 "$rules_path"

if [[ -n "$target_user" && "$target_user" != "root" ]] && id "$target_user" >/dev/null 2>&1; then
    usermod -aG dialout "$target_user"
    echo "Added $target_user to the dialout group."
fi

udevadm control --reload-rules
udevadm trigger --subsystem-match=tty

echo "Installed serial access rules at $rules_path."
echo "Unplug/replug USB serial adapters, then log out and back in."
echo "Until then, use ./run_as_root.sh for immediate access."
