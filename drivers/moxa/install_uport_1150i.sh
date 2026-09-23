#!/usr/bin/env bash
set -euo pipefail

driver_archive="moxa-uport-1100-series-linux-kernel-6.x-driver-v6.0.tgz"
driver_url="https://cdn-cms-frontdoor-dfc8ebanh6bkb3hs.a02.azurefd.net/getmedia/c7a1d4ee-ff6f-46fe-b707-e6e2c6fcc152/${driver_archive}"
driver_sha512="B812A68A776EFC345FB34ACCFA57C928268278BBCCEA674D7A6ED24FDA674FD7A6064FDE23818CA34A760E97229C4F832E78B5C9F9AC273D4932583CF3E27DFE"

usage() {
  cat <<'EOF'
Install Moxa's UPort 1100 Linux 6.x driver and select the default interface.

Usage:
  sudo ./install_uport_1150i.sh [--mode MODE] [--device /dev/ttyUSBn]

Modes:
  rs485-2w   RS-485 two-wire (default)
  rs485-4w   RS-485 four-wire
  rs422      RS-422
  rs232      RS-232

--device is optional. When supplied, setserial also applies the selected mode
to that connected device immediately. Without it, the compiled default applies
when a supported UPort is attached.
EOF
}

fail() {
  echo "ERROR: $*" >&2
  exit 1
}

require_command() {
  command -v "$1" >/dev/null 2>&1 || fail "Required command '$1' was not found."
}

download() {
  local url="$1"
  local destination="$2"
  if command -v curl >/dev/null 2>&1; then
    curl --fail --location --retry 3 --output "$destination" "$url"
  elif command -v wget >/dev/null 2>&1; then
    wget --tries=3 --output-document="$destination" "$url"
  else
    fail "curl or wget is required because the bundled driver archive was not found."
  fi
}

mode="rs485-2w"
device=""
while [[ $# -gt 0 ]]; do
  case "$1" in
    --mode)
      [[ $# -ge 2 ]] || fail "--mode requires a value."
      mode="$2"
      shift 2
      ;;
    --device)
      [[ $# -ge 2 ]] || fail "--device requires a /dev path."
      device="$2"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      fail "Unknown argument '$1'. Run with --help for usage."
      ;;
  esac
done

case "$mode" in
  rs485-2w)
    compile_mode=0
    setserial_mode=1
    ;;
  rs232)
    compile_mode=1
    setserial_mode=0
    ;;
  rs422)
    compile_mode=2
    setserial_mode=2
    ;;
  rs485-4w)
    compile_mode=3
    setserial_mode=3
    ;;
  *)
    fail "Unsupported mode '$mode'. Use rs485-2w, rs485-4w, rs422, or rs232."
    ;;
esac

[[ ${EUID:-$(id -u)} -eq 0 ]] || fail "Run this installer as root (for example, with sudo)."
[[ "$(uname -s)" == "Linux" ]] || fail "This driver installer only runs on Linux."
kernel_release="$(uname -r)"
kernel_major="${kernel_release%%.*}"
[[ "$kernel_major" == "6" ]] || fail "This bundled Moxa driver targets Linux kernel 6.x; running kernel is $kernel_release."
[[ -d "/lib/modules/${kernel_release}/build" ]] || fail "Kernel headers for $kernel_release are missing. Install the matching headers package first."

require_command gcc
require_command make
require_command sha512sum
require_command tar
if [[ -n "$device" ]]; then
  [[ "$device" == /dev/* ]] || fail "--device must be an explicit path below /dev."
  require_command setserial
fi

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
work_dir="$(mktemp -d)"
trap 'rm -rf -- "$work_dir"' EXIT

archive_path="${script_dir}/${driver_archive}"
if [[ ! -f "$archive_path" ]]; then
  archive_path="${work_dir}/${driver_archive}"
  echo "Bundled UPort archive not found; downloading Moxa v6.0 from the official URL..."
  download "$driver_url" "$archive_path"
fi

actual_sha512="$(sha512sum "$archive_path" | awk '{print toupper($1)}')"
[[ "$actual_sha512" == "$driver_sha512" ]] || fail "SHA-512 verification failed for $archive_path. The driver was not installed."

tar -xzf "$archive_path" -C "$work_dir"
source_dir="${work_dir}/mxu11x0"
[[ -x "${source_dir}/mxinstall" ]] || fail "The verified archive did not contain mxu11x0/mxinstall."

echo "Installing Moxa UPort driver for kernel $kernel_release with default mode $mode..."
if ! (cd "$source_dir" && ./mxinstall install "kflags=-DDEFAULT_UART_MODE=${compile_mode}"); then
  fail "Moxa's installer failed. Review the compiler output above; Secure Boot may also reject an unsigned module."
fi

if [[ -n "$device" ]]; then
  [[ -e "$device" ]] || fail "Driver installed, but $device does not exist. Reconnect the UPort and rerun setserial manually."
  setserial "$device" port "$setserial_mode"
  echo "Applied $mode to $device."
fi

echo "UPort driver installation completed. Use /dev/serial/by-id/... when available so device naming remains stable."
