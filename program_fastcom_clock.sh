#!/usr/bin/env bash
set -euo pipefail

[[ "$#" -eq 1 ]] || {
    echo "Usage: $0 /dev/ttyS<N>" >&2
    exit 2
}
[[ "${EUID}" -eq 0 ]] || {
    echo "This restricted Fastcom clock helper must run as root through sudo." >&2
    exit 1
}

python_helper="/usr/local/libexec/serial-loopback-program-fastcom-clock.py"
[[ -f "$python_helper" ]] || {
    echo "Missing installed Fastcom clock programmer: $python_helper" >&2
    exit 1
}
python_path="$(command -v python3 || true)"
[[ -n "$python_path" ]] || {
    echo "python3 is required by the Fastcom clock programmer." >&2
    exit 1
}

exec "$python_path" "$python_helper" "$1"
