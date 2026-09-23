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

exec "$app_path" "$@"
