#!/usr/bin/env bash
set -euo pipefail

if [[ "$(uname -s)" != "Linux" ]]; then
    echo "This helper only runs on Linux." >&2
    exit 1
fi

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
python_bin="${PYTHON:-python3}"

if ! command -v "$python_bin" >/dev/null 2>&1; then
    echo "Python 3 is required." >&2
    exit 1
fi

exec "$python_bin" "${script_dir}/moxa_uport_mode.py" "$@"
