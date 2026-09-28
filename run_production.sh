#!/usr/bin/env bash
set -euo pipefail

if [[ "$(uname -s)" != "Linux" ]]; then
    echo "This production runner is for Linux. On Windows, run: python serial_tester_gui.py" >&2
    exit 1
fi

repo_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
cd "$repo_dir"

run_as_root=0
if [[ "${1:-}" == "--root" ]]; then
    run_as_root=1
    shift
fi

python_bin="${PYTHON:-python3}"
if ! command -v "$python_bin" >/dev/null 2>&1; then
    echo "Python 3 was not found. On Fedora, run: sudo dnf install python3 python3-pip python3-tkinter" >&2
    exit 1
fi

if ! "$python_bin" -c 'import sys; raise SystemExit(sys.version_info < (3, 10))'; then
    echo "Python 3.10 or newer is required." >&2
    exit 1
fi

if ! "$python_bin" -c 'import tkinter' >/dev/null 2>&1; then
    echo "Python Tk support is missing." >&2
    echo "Fedora: sudo dnf install python3-tkinter" >&2
    echo "Debian/Ubuntu: sudo apt install python3-tk python3-venv" >&2
    exit 1
fi

venv_dir="${SERIAL_LOOPBACK_TESTER_VENV:-${repo_dir}/.venv-production}"
venv_python="${venv_dir}/bin/python"
if [[ ! -x "$venv_python" ]]; then
    echo "Creating production virtual environment at $venv_dir..."
    if ! "$python_bin" -m venv "$venv_dir"; then
        echo "Could not create the virtual environment. Install python3-venv (Debian/Ubuntu) or python3 (Fedora)." >&2
        exit 1
    fi
fi

requirements_hash="$("$venv_python" -c 'import hashlib, pathlib; print(hashlib.sha256(pathlib.Path("requirements.txt").read_bytes()).hexdigest())')"
requirements_stamp="${venv_dir}/.requirements.sha256"
installed_hash=""
if [[ -f "$requirements_stamp" ]]; then
    installed_hash="$(<"$requirements_stamp")"
fi

if [[ "$installed_hash" != "$requirements_hash" ]] || ! "$venv_python" -c 'import serial' >/dev/null 2>&1; then
    echo "Installing production dependencies..."
    "$venv_python" -m pip install --disable-pip-version-check -r requirements.txt
    printf '%s\n' "$requirements_hash" > "$requirements_stamp"
fi

if [[ "$run_as_root" -eq 1 ]]; then
    exec bash "${repo_dir}/run_as_root.sh" "$@"
fi

exec "$venv_python" "${repo_dir}/serial_tester_gui.py" "$@"
