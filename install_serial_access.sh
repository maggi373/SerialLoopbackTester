#!/usr/bin/env bash
set -euo pipefail

target_user="${SUDO_USER:-${USER:-}}"
script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"

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
    'SUBSYSTEM=="tty", KERNEL=="ttyS[0-9]*", ATTRS{vendor}=="0x18f7", ATTRS{device}=="0x000a", GROUP="dialout", MODE="0660", TAG+="uaccess"' \
    'SUBSYSTEM=="tty", KERNEL=="ttyS[0-9]*", ATTRS{vendor}=="0x18f7", ATTRS{device}=="0x000b", GROUP="dialout", MODE="0660", TAG+="uaccess"' \
    'SUBSYSTEM=="usb", ATTR{idVendor}=="110a", ATTR{idProduct}=="1150", GROUP="dialout", MODE="0660", TAG+="uaccess"' \
    'SUBSYSTEM=="usb", ATTR{idVendor}=="110a", ATTR{idProduct}=="1151", GROUP="dialout", MODE="0660", TAG+="uaccess"' \
    > "$rules_path"
chmod 0644 "$rules_path"

if [[ -n "$target_user" && "$target_user" != "root" ]] && id "$target_user" >/dev/null 2>&1; then
    usermod -aG dialout "$target_user"
    echo "Added $target_user to the dialout group."

    source_helper="${script_dir}/program_fastcom_clock.sh"
    source_programmer="${script_dir}/program_fastcom_clock.py"
    if [[ ! -f "$source_helper" || ! -f "$source_programmer" ]]; then
        echo "Missing restricted Fastcom clock helper files beside the installer." >&2
        exit 1
    fi
    helper_path="/usr/local/libexec/serial-loopback-program-fastcom-clock"
    programmer_path="/usr/local/libexec/serial-loopback-program-fastcom-clock.py"
    install -d -o root -g root -m 0755 /usr/local/libexec
    install -o root -g root -m 0755 "$source_helper" "$helper_path"
    install -o root -g root -m 0755 "$source_programmer" "$programmer_path"

    command -v visudo >/dev/null 2>&1 || {
        echo "visudo is required to install the restricted Fastcom permission." >&2
        exit 1
    }
    install -d -o root -g root -m 0750 /etc/sudoers.d
    target_uid="$(id -u "$target_user")"
    sudoers_path="/etc/sudoers.d/serial-loopback-tester-fastcom-${target_uid}"
    sudoers_temp="$(mktemp)"
    trap 'rm -f -- "${sudoers_temp:-}"' EXIT
    printf '%s ALL=(root) NOPASSWD: %s *\n' "$target_user" "$helper_path" > "$sudoers_temp"
    chmod 0440 "$sudoers_temp"
    visudo -cf "$sudoers_temp" >/dev/null
    install -o root -g root -m 0440 "$sudoers_temp" "$sudoers_path"
    echo "Installed restricted Fastcom PCI-335 clock-programming permission for $target_user."
fi

udevadm control --reload-rules
udevadm trigger --subsystem-match=tty
udevadm trigger --subsystem-match=usb --attr-match=idVendor=110a

echo "Installed serial access rules at $rules_path."
echo "Unplug/replug USB serial adapters, then log out and back in."
echo "Until then, use ./run_as_root.sh for immediate access."
