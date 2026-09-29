#!/usr/bin/env bash
set -euo pipefail

usage() {
    echo "Usage: $0 get /dev/ttyS<N> | set /dev/ttyS<N> {921600|1843200}" >&2
    exit 2
}

[[ "${EUID}" -eq 0 ]] || {
    echo "This restricted Fastcom helper must run as root through sudo." >&2
    exit 1
}

action="${1:-}"
requested_device="${2:-}"
[[ "$action" == "get" || "$action" == "set" ]] || usage
[[ "$requested_device" =~ ^/dev/ttyS[0-9]+$ ]] || {
    echo "Only /dev/ttyS<N> devices are accepted." >&2
    exit 2
}

resolved_device="$(readlink -f -- "$requested_device")"
[[ "$resolved_device" =~ ^/dev/ttyS[0-9]+$ && -c "$resolved_device" ]] || {
    echo "$requested_device is not a local ttyS character device." >&2
    exit 2
}

tty_name="${resolved_device##*/}"
sys_cursor="$(readlink -f -- "/sys/class/tty/${tty_name}/device")"
fastcom_card_found=0
while [[ -n "$sys_cursor" && "$sys_cursor" != "/" ]]; do
    if [[ -r "${sys_cursor}/vendor" && -r "${sys_cursor}/device" ]]; then
        pci_vendor="$(tr '[:upper:]' '[:lower:]' < "${sys_cursor}/vendor")"
        pci_device="$(tr '[:upper:]' '[:lower:]' < "${sys_cursor}/device")"
        if [[ "$pci_vendor" == "0x18f7" && ( "$pci_device" == "0x000a" || "$pci_device" == "0x000b" ) ]]; then
            fastcom_card_found=1
            break
        fi
    fi
    sys_cursor="${sys_cursor%/*}"
    [[ -n "$sys_cursor" ]] || sys_cursor="/"
done

[[ "$fastcom_card_found" -eq 1 ]] || {
    echo "$resolved_device is not a Fastcom 232/4 or 232/8 PCI-335 port (18f7:000a/000b)." >&2
    exit 2
}

setserial_path="$(command -v setserial || true)"
[[ -n "$setserial_path" ]] || {
    echo "setserial is not installed (Fedora: sudo dnf install setserial)." >&2
    exit 1
}

if [[ "$action" == "get" ]]; then
    [[ "$#" -eq 2 ]] || usage
    exec "$setserial_path" -a "$resolved_device"
fi

[[ "$#" -eq 3 ]] || usage
baud_base="$3"
[[ "$baud_base" == "921600" || "$baud_base" == "1843200" ]] || {
    echo "Only baud_base 921600 (Fastcom hardware default) or 1843200 (8250_exar default) is allowed." >&2
    exit 2
}

"$setserial_path" "$resolved_device" baud_base "$baud_base"
exec "$setserial_path" -a "$resolved_device"
