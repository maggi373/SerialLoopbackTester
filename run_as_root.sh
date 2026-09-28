#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
app_command=()

for candidate in "${script_dir}"/SerialLoopbackTester-v*-linux-*; do
    if [[ -f "$candidate" && -x "$candidate" ]]; then
        app_command=("$candidate")
        break
    fi
done

source_venv="${SERIAL_LOOPBACK_TESTER_VENV:-${script_dir}/.venv-production}"
if [[ "${#app_command[@]}" -eq 0 && -x "${source_venv}/bin/python" && -f "${script_dir}/serial_tester_gui.py" ]]; then
    app_command=("${source_venv}/bin/python" "${script_dir}/serial_tester_gui.py")
fi

if [[ "${#app_command[@]}" -eq 0 ]]; then
    echo "No packaged executable or production repository environment was found beside this script." >&2
    echo "For a repository checkout, run: bash ./run_production.sh --root" >&2
    exit 1
fi

if [[ "${EUID}" -eq 0 ]]; then
    if [[ -n "${SUDO_USER:-}" && "${SUDO_USER}" != "root" ]]; then
        original_uid="$(id -u "$SUDO_USER")"
        original_gid="$(id -g "$SUDO_USER")"
        original_home="$(getent passwd "$SUDO_USER" | cut -d: -f6)"
        export SERIAL_LOOPBACK_TESTER_CONFIG_HOME="${original_home}/.config"
        export SERIAL_LOOPBACK_TESTER_CONFIG_UID="$original_uid"
        export SERIAL_LOOPBACK_TESTER_CONFIG_GID="$original_gid"
    fi
    exec "${app_command[@]}" "$@"
fi

if ! command -v sudo >/dev/null 2>&1; then
    echo "sudo is required. Run this script as root or install sudo." >&2
    exit 1
fi

# Preserve the desktop-session variables required by Tk/X11/Wayland while the
# application itself runs as root and can open restricted serial devices. The
# explicit config variables keep settings in the desktop user's config folder
# and let the application retain that user's ownership on every atomic save.
config_home="${XDG_CONFIG_HOME:-${HOME}/.config}"
if [[ "$config_home" != /* ]]; then
    config_home="${HOME}/.config"
fi
config_uid="$(id -u)"
config_gid="$(id -g)"
exec sudo --preserve-env=DISPLAY,XAUTHORITY,WAYLAND_DISPLAY,XDG_RUNTIME_DIR,DBUS_SESSION_BUS_ADDRESS \
    env SERIAL_LOOPBACK_TESTER_CONFIG_HOME="$config_home" \
        SERIAL_LOOPBACK_TESTER_CONFIG_UID="$config_uid" \
        SERIAL_LOOPBACK_TESTER_CONFIG_GID="$config_gid" \
        "${app_command[@]}" "$@"
