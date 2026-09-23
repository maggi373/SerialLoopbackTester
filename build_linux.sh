#!/usr/bin/env bash
set -euo pipefail

if [[ "$(uname -s)" != "Linux" ]]; then
    echo "This build script must be run on Linux." >&2
    exit 1
fi

repo_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
cd "$repo_dir"

python_bin="${PYTHON:-python3}"
app_version="1.3.0"
architecture="$(uname -m)"
portable_name="SerialLoopbackTester-v${app_version}-linux-${architecture}"
archive_path="${repo_dir}/dist/${portable_name}.tar.gz"
moxa_driver_dir="${repo_dir}/drivers/moxa"
moxa_cache_dir="${repo_dir}/build/moxa-drivers"
uport_archive="moxa-uport-1100-series-linux-kernel-6.x-driver-v6.0.tgz"
uport_url="https://cdn-cms-frontdoor-dfc8ebanh6bkb3hs.a02.azurefd.net/getmedia/c7a1d4ee-ff6f-46fe-b707-e6e2c6fcc152/${uport_archive}"
uport_sha512="B812A68A776EFC345FB34ACCFA57C928268278BBCCEA674D7A6ED24FDA674FD7A6064FDE23818CA34A760E97229C4F832E78B5C9F9AC273D4932583CF3E27DFE"
real_tty_archive="moxa-real-tty-linux-kernel-6.x-driver-v6.2.tar"
real_tty_url="https://cdn-cms-frontdoor-dfc8ebanh6bkb3hs.a02.azurefd.net/getmedia/c1b228ac-3fc5-4612-bb49-99cbcc140ebe/${real_tty_archive}"
real_tty_sha512="237F77E2F40BCA7D51F88942ACD49C96AA5328E9C34A61261E950042E08C2E1406221321BB315258BF2AB1FC8304E19D3BC495949BDB960E1C68964C1C9F330B"

download_and_verify() {
    local url="$1"
    local expected_sha512="$2"
    local destination="$3"
    local actual_sha512=""

    if [[ -f "$destination" ]]; then
        actual_sha512="$(sha512sum "$destination" | awk '{print toupper($1)}')"
    fi
    if [[ "$actual_sha512" != "$expected_sha512" ]]; then
        local partial_path="${destination}.part"
        rm -f -- "$partial_path"
        echo "Downloading $(basename "$destination") from Moxa..."
        if command -v curl >/dev/null 2>&1; then
            curl --fail --location --retry 3 --output "$partial_path" "$url"
        elif command -v wget >/dev/null 2>&1; then
            wget --tries=3 --output-document="$partial_path" "$url"
        else
            echo "curl or wget is required to bundle the Moxa driver archives." >&2
            exit 1
        fi
        actual_sha512="$(sha512sum "$partial_path" | awk '{print toupper($1)}')"
        if [[ "$actual_sha512" != "$expected_sha512" ]]; then
            rm -f -- "$partial_path"
            echo "SHA-512 verification failed for $(basename "$destination")." >&2
            exit 1
        fi
        mv -- "$partial_path" "$destination"
    fi
}

if ! "$python_bin" -c "import tkinter" >/dev/null 2>&1; then
    echo "Python Tk support is missing. Install your distribution's python3-tk/Tk package." >&2
    exit 1
fi

echo "Installing Python dependencies..."
"$python_bin" -m pip install -r requirements.txt -r requirements-build.txt

echo "Running tests..."
"$python_bin" -m unittest discover -s tests

echo "Building Linux application with PyInstaller..."
"$python_bin" -m PyInstaller \
    --noconfirm \
    --clean \
    --noupx \
    --onedir \
    --windowed \
    --hidden-import serial.urlhandler.protocol_socket \
    --name "$portable_name" \
    serial_tester_gui.py

executable_path="${repo_dir}/dist/${portable_name}/${portable_name}"
if [[ ! -x "$executable_path" ]]; then
    echo "Build failed: executable was not created at $executable_path" >&2
    exit 1
fi

cp README.md "${repo_dir}/dist/${portable_name}/README.md"

bundle_driver_dir="${repo_dir}/dist/${portable_name}/drivers/moxa"
mkdir -p "$bundle_driver_dir"
cp "${moxa_driver_dir}/README.md" "${moxa_driver_dir}/install_uport_1150i.sh" \
    "${moxa_driver_dir}/install_nport_real_tty.sh" "$bundle_driver_dir/"
chmod +x "$bundle_driver_dir/install_uport_1150i.sh" "$bundle_driver_dir/install_nport_real_tty.sh"

if [[ "${SKIP_MOXA_DRIVERS:-0}" == "1" ]]; then
    echo "Skipping bundled Moxa archives (installers will download and verify them when run)."
else
    command -v sha512sum >/dev/null 2>&1 || {
        echo "sha512sum is required to bundle the Moxa driver archives." >&2
        exit 1
    }
    mkdir -p "$moxa_cache_dir"
    download_and_verify "$uport_url" "$uport_sha512" "${moxa_cache_dir}/${uport_archive}"
    download_and_verify "$real_tty_url" "$real_tty_sha512" "${moxa_cache_dir}/${real_tty_archive}"
    cp "${moxa_cache_dir}/${uport_archive}" "${moxa_cache_dir}/${real_tty_archive}" "$bundle_driver_dir/"
fi

echo "Creating portable tar.gz archive..."
tar -C "${repo_dir}/dist" -czf "$archive_path" "$portable_name"

echo "Portable folder ready: ${repo_dir}/dist/${portable_name}"
echo "Portable archive ready: $archive_path"
