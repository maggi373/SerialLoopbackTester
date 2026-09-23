#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
app_path=""

for candidate in "${script_dir}"/SerialLoopbackTester-v*-linux-*; do
    if [[ -f "$candidate" && -x "$candidate" ]]; then
        app_path="$candidate"
        break
    fi
done

if [[ -z "$app_path" ]]; then
    echo "Serial Loopback Tester executable was not found beside this script." >&2
    exit 1
fi

if [[ "${EUID}" -eq 0 ]]; then
    exec "$app_path" "$@"
fi

if ! command -v sudo >/dev/null 2>&1; then
    echo "sudo is required. Run this script as root or install sudo." >&2
    exit 1
fi

# Preserve the desktop-session variables required by Tk/X11/Wayland while the
# application itself runs as root and can open restricted serial devices.
exec sudo --preserve-env=DISPLAY,XAUTHORITY,WAYLAND_DISPLAY,XDG_RUNTIME_DIR,DBUS_SESSION_BUS_ADDRESS \
    "$app_path" "$@"
