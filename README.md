# Serial Loopback Tester
Version: `1.3.1`
Made by: `maggi373`

Python GUI tool for:
- 40 RS232 loopback tests (same port send/receive)
- Per-port RS232 roles: normal loopback testing, passive RS485 reply, or a stateful PARO/Digiquartz sensor simulator
- PARO ports support addressed commands, independent device IDs, measurement ramps, configuration reads/writes, errors, held/continuous readings, and output masks based on the Arduino `parosim` implementation
- 8 RS485 tests (RS485 sends -> an RS232 port in **RS485 Reply** role echoes the received bytes -> RS485 verifies the response)
- RS485 reply routing is port-agnostic; any RS232 channel assigned the dedicated reply role can answer, without transmitting its own loopback packets
- Message mismatches are ignored for the first 2 seconds after communication starts on every channel; port-open and serial I/O errors remain visible
- Customizable RS232 and RS485 counts from Settings (defaults: 40 and 8) (MAX: 256 and 128)
- Combined Overview page for all ports/channels with color status bars and row outlines
- Overview supports compact 2-column row mode, 2-column card mode, and optional preset filtering
- Live worker updates are batched, hidden tabs are not redrawn, and rendering pauses while the window is moving/resizing
- Health summary block in Overview (alarm, counts, and recent failures)
- Health page with live alarm state and current FAIL/ERROR list
- Fault Review page for PASS -> FAIL/ERROR transitions
- 5 named presets (toolbar buttons) with per-preset channel name selection on the Presets page
- Presets apply enable/disable states to channels based on selected channel names
- The last applied preset and Overview preset filter are restored on the next launch
- Editable serial-port mapping and custom port names (`COM...` on Windows, `/dev/...` on Linux)
- Raw TCP serial endpoints (`socket://host:port`) for devices such as Moxa NPort
- Serial-port dropdowns with manual typing allowed
- JSON settings file (`serial_tester_settings.json`)
- Default test interval: 25 ms
- Default baudrate: 19200
- Default payload: 8 bytes
- Global serial settings control for applying baudrate, test interval, and packet size to all channels
- Fullscreen support (`Fullscreen` button, `F11` toggle, `Esc` exit)
- Auto-start tests 2 seconds after launch (default ON)
- Optional startup setting: launch in fullscreen by default
- Optional startup setting (default ON): delay communications by 2 seconds

![solder.py](https://files.thorfusion.com/images/serial.jpg)

## Requirements
- Python 3.10+
- `pyserial`

Install:
```powershell
pip install -r requirements.txt
```

## Run
```powershell
python serial_tester_gui.py
```

### Linux

Install Python, Tk, and virtual-environment support (Debian/Ubuntu example):

```bash
sudo apt install python3 python3-pip python3-tk python3-venv
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python serial_tester_gui.py
```

The user running the application needs read/write access to the serial devices. On Debian/Ubuntu this commonly means membership in `dialout`:

```bash
sudo usermod -aG dialout "$USER"
```

Log out and back in after changing group membership. Other distributions may use a different group; check the owner/group reported by `ls -l /dev/ttyUSB0` or the relevant device.

On the first Linux launch only, the application detects all serial devices reported by pyserial, assigns them to RS232 configuration rows, expands the RS232 row count when more than 40 are present, and saves the result. Auto-added ports start disabled so no unknown RS-485 device is sent loopback traffic; assign each port's role and enable the required ports or apply a preset. Later launches preserve the saved mapping and do not auto-add newly connected devices.

#### Fedora

Install the runtime dependencies and run from source:

```bash
sudo dnf install python3 python3-pip python3-tkinter
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python serial_tester_gui.py
```

For a downloaded Linux release, extract and run the packaged executable instead:

```bash
tar -xzf SerialLoopbackTester-v1.3.1-linux-x86_64.tar.gz
cd SerialLoopbackTester-v1.3.1-linux-x86_64
./start.sh
```

For immediate unrestricted serial-port access, use the included root startup script. It preserves the graphical-session variables needed by Tk:

```bash
./start_as_root.sh
```

Both startup scripts pass any command-line arguments through to the packaged executable. `run_as_root.sh` is also included as a compatibility alias for direct use.

This runs the full application as root, so its settings may be stored under root's config directory. To make access persistent while running the application as your normal account, install the included udev rules and group membership instead:

```bash
./install_serial_access.sh
```

Then unplug/replug USB serial devices and log out and back in. The installer covers `/dev/ttyUSB*`, `/dev/ttyACM*`, and Moxa Real TTY `/dev/ttyr*` devices.

Check the group assigned to the adapter and add your account to it. Fedora normally uses `dialout` for USB serial devices:

```bash
ls -l /dev/ttyUSB0
sudo usermod -aG dialout "$USER"
getent group dialout
```

The `usermod` command above adds the currently logged-in Fedora user to `dialout`; do not replace `$USER` with `root`. Log out of the desktop completely and back in, then verify with `id -nG` before starting the application. Using `/dev/serial/by-id/...` in the application is recommended because `/dev/ttyUSBn` numbers can change after reconnecting devices.

To compile the bundled Moxa kernel modules on Fedora, install the compiler and development files matching the currently running kernel:

```bash
sudo dnf install gcc make kernel-devel-$(uname -r) elfutils-libelf-devel openssl-devel setserial
test -e /lib/modules/$(uname -r)/build
```

Then run the appropriate installer from the extracted application folder:

```bash
# UPort 1150I, RS-485 two-wire
sudo bash ./drivers/moxa/install_uport_1150i.sh --mode rs485-2w

# Optional NPort 5410 local TTY mapping
sudo bash ./drivers/moxa/install_nport_real_tty.sh --nport-ip 192.168.1.50
```

The bundled Moxa modules are officially for Linux kernel 6.x, and the wrappers stop on other kernel major versions. Fedora systems running kernel 7.x can still use the application and NPort direct TCP mode, but should not force-install these modules. Use a supported Fedora kernel 6.x installation or obtain a kernel 7-compatible driver from Moxa. On Secure Boot systems, the newly compiled module may also need to be signed and enrolled before it can load.

## Build Installer (Windows)
Install build tools:
```powershell
pip install -r requirements.txt -r requirements-build.txt
```

Build portable folder and standard ZIP only:
```powershell
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1 -SkipInno
```

Build EXE + Setup installer (requires Inno Setup 6):
```powershell
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1
```
The build script checks `ISCC.exe` in `PATH`, `%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe`, and standard Program Files locations.
Installer includes an optional checkbox to start the app with Windows (Startup folder shortcut for the installing user).
The default installation is per-user and does not require administrator elevation.

Outputs:
- Portable folder: `dist\SerialLoopbackTester-v1.3.1-portable\`
- Inspectable portable ZIP: `dist\SerialLoopbackTester-v1.3.1-portable.zip`
- Installer: `dist\installer\SerialLoopbackTester-v1.3.1-installer.exe`

The portable ZIP is a standard archive that can be opened with 7-Zip or Windows Explorer. It uses PyInstaller folder mode, so the application files are visible instead of being wrapped in a self-extracting one-file executable.

## Build Portable Package (Linux)

Run the build on the Linux architecture you want to support; PyInstaller packages are platform-specific:

```bash
chmod +x ./build_linux.sh
./build_linux.sh
```

This runs the tests and creates both a portable folder and `tar.gz` archive under `dist/`, named with the machine architecture (for example, `linux-x86_64`). By default it also downloads Moxa's official UPort 1100 v6.0 and Real TTY v6.2 source archives, verifies their SHA-512 checksums, and includes them under `drivers/moxa`. Set `SKIP_MOXA_DRIVERS=1` to omit the archives; the included installers can download and verify them later.

## Automated GitHub releases

The workflow in `.github/workflows/release.yml` builds and publishes release files when a GitHub release is published or a version tag is pushed. Both `1.3.1` and `v1.3.1` tag styles are accepted, but the numeric version must match `APP_VERSION`; for this release:

```bash
git tag v1.3.1
git push origin v1.3.1
```

The GitHub release receives:

- Windows portable ZIP
- Windows Inno Setup installer
- Linux x86_64 portable `tar.gz`
- Moxa UPort 1100 and Real TTY source-driver archives
- `SHA256SUMS.txt` covering every uploaded file

Running the workflow manually builds downloadable workflow artifacts but intentionally does not create an untagged GitHub release.

To repair an existing release, run **Build release** manually and enter its tag in the `release_tag` field. The workflow checks out that tag and uploads/replaces all release assets.

## Moxa on Linux

### NPort 5410

The NPort 5410 is a four-port RS-232-only device server; it cannot provide a physical RS-485 interface. Use the UPort 1150I below for the RS-485 side. There are two supported ways to use the NPort 5410 on the RS-232 side:

- **Direct TCP:** configure each NPort channel as a TCP Server, then enter a URL such as `socket://192.168.1.50:4001` directly in an RS232 port field. Use `4002`, `4003`, and `4004` for the other channels if those are the ports configured on the NPort. Raw TCP cannot change the NPort's serial settings, so configure baud rate, parity, data bits, and stop bits in the NPort web interface.
- **Real TTY:** configure the channels as Real COM Mode, then install and map Moxa's driver:

  ```bash
  sudo bash ./drivers/moxa/install_nport_real_tty.sh --nport-ip 192.168.1.50
  ```

  The first mapped four-port NPort normally appears as `/dev/ttyr00` through `/dev/ttyr03`; enter those paths in the application.

Direct TCP is the simpler option and needs no kernel module. The Real TTY option is useful when other software specifically requires a local TTY device.

### UPort 1150I RS-485

Linux can identify this USB adapter with an in-kernel driver, but selecting the 1150I electrical interface requires Moxa's driver for this use. The bundled wrapper compiles Moxa's Linux 6.x driver with RS-485 two-wire as the default:

```bash
sudo bash ./drivers/moxa/install_uport_1150i.sh --mode rs485-2w
```

If `setserial` is installed and the connected adapter is `/dev/ttyUSB0`, the mode can also be applied immediately:

```bash
sudo bash ./drivers/moxa/install_uport_1150i.sh --mode rs485-2w --device /dev/ttyUSB0
```

Both driver installers check for root access, Linux 6.x, matching kernel headers, GCC, Make, archive integrity, and expected archive contents before running Moxa's installer. Secure Boot may require signing and enrolling these out-of-tree modules according to the Linux distribution's instructions. See `drivers/moxa/README.md` for mode and mapping details.

## Usage
1. Open the **Settings** tab.
2. Configure RS232 ports and names.
   - Set **Role** to **Loopback Test** for normal testing.
   - Set **Role** to **RS485 Reply** for a passive port that sends nothing by itself and echoes only bytes it receives.
   - Set **Role** to **PARO Simulator** to make that port behave like a PARO sensor.
   - Set a **PARO Device ID** from `00` to `99` for each simulated sensor. The PARO/Arduino default serial format is `9600 8N1`.
3. Configure each RS485 port and name. Assign at least one connected RS232 channel the **RS485 Reply** role for the return path.
4. Click **Save Settings**.
5. Start tests using **Start RS232**, **Start RS485**, or **Start All**.
6. Use **Overview** tab to see all entries at once:
   - Use **Compact View (2 Columns)** to switch to the old compact row layout for 1920x1080 screens
   - Use **Hide Non-Preset Ports** to show only the channel names selected by the last applied preset
   - Top health strip shows alarm state, fault log count, and live totals
   - Green bar = good message match
   - Purple bar = communication recovered after prior fault on that channel
   - Yellow bar = standby/running without pass yet
   - Red bar = wrong message or serial-port error (the row text distinguishes them)
7. Use **Health** tab to watch global pass/fail totals, total errors, run time, 1-hour fail count, alarm color, and **Faults Logged** count next to the alarm box.
   - Green = good communication only after at least 1 hour runtime and 0 errors in the last 1 hour
   - Purple = good communication, but faults are logged (recovered state)
   - Red = active alarm (current FAIL/ERROR issue)
8. Use **Fault Review** tab to review channels that were PASS and then changed to FAIL/ERROR.
9. In **Overview**, use per-row **Start** and **Stop** buttons to control individual RS232/RS485 channels.
10. In **Settings > Application**, enable/disable **Auto-start tests 2 seconds after launch** (default ON).
11. In **Settings > Application**, enable **Start application in fullscreen** if desired.
12. In **Settings > Application**, keep **Delay communications startup by 2 seconds** enabled (default ON) to wait before first TX when starting tests manually.
13. In **Settings > Application**, use **Enable All Ports** or **Disable All Ports** for a global channel state change.
14. In **Settings > Application**, use **Global Serial Settings** to apply baudrate, interval, and packet size to every RS232/RS485 channel.
15. In **Settings > Application**, set **RS232 count** and **RS485 count**, then click **Apply Counts** to resize channels (defaults are 40 and 8).
16. To disable a port completely, either uncheck **Enabled** or leave the port field blank. Blank ports are skipped.
17. In **Edit Selected RS232 Port**, use **Apply To All RS232 (Keep Name/Port)** to copy serial settings to all RS232 rows while preserving each row's Name and Port.
18. In **Edit Selected RS485 Port**, use **Apply To All RS485 (Keep Name/Port)** to copy serial settings to all RS485 channels while preserving each channel's Name and Port.
19. In **Presets**, set up each preset name and pick channel names for that preset.
   - Each preset card header also shows the custom name for quick identification.
   - Select one or more RS232 names, choose a role, and click **Assign** to make the preset apply that role.
   - Click **Keep Current** to make the preset enable those selected RS232 names without changing their existing roles.
20. Use the 5 preset buttons to the right of **Fullscreen** to apply presets:
   - Selected name-matching channels are set to **Enabled**.
   - Non-selected channels are set to **Disabled**.
   - Running channels that become disabled are stopped automatically.
   - When that protocol group is already running, newly selected channels start automatically and selected running channels continue.
   - Applying a preset while tests are idle only changes enablement; it does not start communications.
21. Use **Refresh Port List** to reload dropdown values from system ports; manual values are still allowed (including blank/duplicate/custom values, Linux `/dev/serial/by-id/...` paths, and raw TCP URLs such as `socket://192.168.1.50:4001`).
    - On Linux, automatic assignment happens only when no settings file exists. Refreshing later updates dropdown choices but does not change configured rows.

## Settings file
- Windows settings path (script + EXE): `%USERPROFILE%\Documents\SerialLoopbackTester\serial_tester_settings.json`.
- On Windows, the app resolves this using the system **My Documents** known-folder API (works with localized folder names).
- If no settings file exists there yet, the app does a one-time copy from legacy `%APPDATA%\SerialLoopbackTester\serial_tester_settings.json` (if present).
- Linux settings path: `$XDG_CONFIG_HOME/SerialLoopbackTester/serial_tester_settings.json`, falling back to `~/.config/SerialLoopbackTester/serial_tester_settings.json`.
- On Linux, an older settings file under `~/Documents/SerialLoopbackTester/` is migrated automatically.
- The file is auto-created with default 40 RS232 entries and 8 RS485 entries on first run.
- Settings are validated and normalized when loaded.
- Fullscreen default lives in `ui.start_fullscreen`.
- Auto-start-after-launch default lives in `ui.auto_start_after_launch_2s`.
- 2-second startup delay default lives in `ui.delay_comm_start_2s`.
- Overview compact mode default lives in `ui.overview_compact_view`.
- Overview preset filtering default lives in `ui.overview_hide_non_preset_ports`.
- The last applied preset lives in `ui.active_preset_idx`.
- Global serial control values live in `ui.global_baudrate`, `ui.global_interval_ms`, and `ui.global_packet_size_bytes`.
- Channel counts live in `ui.rs232_count` and `ui.rs485_pair_count`.
- Each RS232 entry stores its assignment in `mode` (`loopback`, `rs485_reply`, or `paro`) and its simulated address in `paro_device_id`.
- Preset button labels, selected channel names, and optional per-name RS232 role assignments live in `ui.presets`.
- Port names are not forced to be unique/valid, so you can stage configs on systems with fewer serial ports.
- Presets match channel names, so multiple names can use the same serial port; duplicate channel names are enabled/disabled together.
