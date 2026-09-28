#!/usr/bin/env bash
set -euo pipefail

driver_archive="moxa-real-tty-linux-kernel-6.x-driver-v6.2.tar"
driver_url="https://cdn-cms-frontdoor-dfc8ebanh6bkb3hs.a02.azurefd.net/getmedia/c1b228ac-3fc5-4612-bb49-99cbcc140ebe/${driver_archive}"
driver_sha512="237F77E2F40BCA7D51F88942ACD49C96AA5328E9C34A61261E950042E08C2E1406221321BB315258BF2AB1FC8304E19D3BC495949BDB960E1C68964C1C9F330B"

usage() {
  cat <<'EOF'
Install Moxa's NPort Real TTY Linux 6.x driver and optionally map an NPort 5410.

Usage:
  sudo ./install_nport_real_tty.sh [--nport-ip ADDRESS]
       [--ports 4] [--data-port PORT --command-port PORT]

Before mapping, configure every required NPort channel for Real COM mode in
the NPort web interface. If --nport-ip is omitted, only the driver is installed.
The Moxa installer is interactive and may ask you to press Enter.
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

nport_ip=""
ports=4
data_port=""
command_port=""
while [[ $# -gt 0 ]]; do
  case "$1" in
    --nport-ip)
      [[ $# -ge 2 ]] || fail "--nport-ip requires a value."
      nport_ip="$2"
      shift 2
      ;;
    --ports)
      [[ $# -ge 2 ]] || fail "--ports requires a value."
      ports="$2"
      shift 2
      ;;
    --data-port)
      [[ $# -ge 2 ]] || fail "--data-port requires a value."
      data_port="$2"
      shift 2
      ;;
    --command-port)
      [[ $# -ge 2 ]] || fail "--command-port requires a value."
      command_port="$2"
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

[[ "$ports" =~ ^[1-9][0-9]*$ ]] || fail "--ports must be a positive integer."
[[ -z "$data_port" || -n "$command_port" ]] && [[ -z "$command_port" || -n "$data_port" ]] || fail "--data-port and --command-port must be supplied together."
if [[ -n "$data_port" ]]; then
  [[ "$data_port" =~ ^[0-9]+$ ]] && (( data_port >= 1 && data_port <= 65535 )) || fail "--data-port must be from 1 to 65535."
  [[ "$command_port" =~ ^[0-9]+$ ]] && (( command_port >= 1 && command_port <= 65535 )) || fail "--command-port must be from 1 to 65535."
fi
[[ -z "$nport_ip" || "$nport_ip" != -* ]] || fail "--nport-ip is invalid."

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

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
work_dir="$(mktemp -d)"
trap 'rm -rf -- "$work_dir"' EXIT

archive_path="${script_dir}/${driver_archive}"
if [[ ! -f "$archive_path" ]]; then
  archive_path="${work_dir}/${driver_archive}"
  echo "Bundled Real TTY archive not found; downloading Moxa v6.2 from the official URL..."
  download "$driver_url" "$archive_path"
fi

actual_sha512="$(sha512sum "$archive_path" | awk '{print toupper($1)}')"
[[ "$actual_sha512" == "$driver_sha512" ]] || fail "SHA-512 verification failed for $archive_path. The driver was not installed."

tar -xf "$archive_path" -C "$work_dir"
source_dir="${work_dir}/moxa"
[[ -x "${source_dir}/mxinst" ]] || fail "The verified archive did not contain moxa/mxinst."

echo "Starting Moxa's interactive Real TTY installer for kernel $kernel_release..."
if ! (cd "$source_dir" && ./mxinst); then
  fail "Moxa's installer failed. Review the output above; missing libraries or Secure Boot module signing may need attention."
fi

if [[ -n "$nport_ip" ]]; then
  mapper="/usr/lib/npreal2/driver/mxaddsvr"
  [[ -x "$mapper" ]] || fail "Driver installed, but $mapper was not created."
  echo "Mapping $ports Real TTY ports for NPort $nport_ip..."
  if [[ -n "$data_port" ]]; then
    "$mapper" "$nport_ip" "$ports" "$data_port" "$command_port"
  else
    "$mapper" "$nport_ip" "$ports"
  fi
  echo "Mapping completed. The first NPort normally appears as /dev/ttyr00 through /dev/ttyr03."
else
  echo "Real TTY driver installation completed. Map the NPort later with /usr/lib/npreal2/driver/mxaddsvr."
fi
