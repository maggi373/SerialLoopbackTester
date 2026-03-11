# Serial Loopback Tester
Version: `1.0.0`  
Made by: `maggi373`

Python GUI tool for:
- 40 RS232 loopback tests (same port send/receive)
- 8 RS485 pair tests (sender -> receiver -> echo back -> sender verify)
- Combined Overview page for all ports/pairs with color status bars
- Overview supports compact 2-column row mode and 2-column card mode
- Health summary block in Overview (alarm, counts, and recent failures)
- Health page with live alarm state and current FAIL/ERROR list
- Fault Review page for PASS -> FAIL/ERROR transitions
- 5 named presets (toolbar buttons) with per-preset COM port selection on the Presets page
- Presets apply enable/disable states to channels based on selected COM ports
- Editable COM port mapping and custom port names
- COM dropdowns (with manual typing allowed)
- JSON settings file (`serial_tester_settings.json`)
- Default test interval: 100 ms
- Fullscreen support (`Fullscreen` button, `F11` toggle, `Esc` exit)
- Auto-start tests 2 seconds after launch (default ON)
- Optional startup setting: launch in fullscreen by default
- Optional startup setting (default ON): delay communications by 2 seconds

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

## Build Installer (Windows)
Install build tools:
```powershell
pip install -r requirements.txt -r requirements-build.txt
```

Build portable EXE only:
```powershell
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1 -SkipInno
```

Build EXE + Setup installer (requires Inno Setup 6):
```powershell
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1
```
The build script checks `ISCC.exe` in `PATH`, `%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe`, and standard Program Files locations.
Installer includes an optional checkbox to start the app with Windows (Startup folder shortcut for the installing user).

Outputs:
- Portable EXE: `dist\SerialLoopbackTester-v1.0.0-portable.exe`
- Installer: `dist\installer\SerialLoopbackTester-v1.0.0-installer.exe`

## Usage
1. Open the **Settings** tab.
2. Configure RS232 ports and names.
3. Configure RS485 sender/echo port pairs and names.
4. Click **Save Settings**.
5. Start tests using **Start RS232**, **Start RS485**, or **Start All**.
6. Use **Overview** tab to see all entries at once:
   - Use **Compact View (2 Columns)** to switch to the old compact row layout for 1920x1080 screens
   - Top health strip shows alarm state, fault log count, and live totals
   - Green bar = good message match
   - Purple bar = communication recovered after prior fault on that channel
   - Yellow bar = standby/running without pass yet
   - Red bar = wrong message/error
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
14. To disable a port completely, either uncheck **Enabled** or leave the port field blank. Blank ports are skipped.
15. In **Edit Selected RS232 Port**, use **Apply To All RS232 (Keep Name/Port)** to copy serial settings to all RS232 rows while preserving each row's Name and Port.
16. In **Edit Selected RS485 Pair**, use **Apply To All RS485 (Keep Name/Ports)** to copy serial settings to all RS485 pairs while preserving each pair's Name, Sender Port, and Echo Port.
17. In **Presets**, set up each preset name and pick COM ports for that preset.
   - Each preset card header also shows the custom name for quick identification.
18. Use the 5 preset buttons to the right of **Fullscreen** to apply presets:
   - Selected COM-matching channels are set to **Enabled**.
   - Non-selected channels are set to **Disabled**.
   - Running channels that become disabled are stopped automatically.
19. Use **Refresh COM List** to reload dropdown values from system ports; manual values are still allowed (including blank/duplicate/custom values).

## Settings file
- Script mode path: `serial_tester_settings.json` next to `serial_tester_gui.py`.
- Installed EXE path: `%APPDATA%\SerialLoopbackTester\serial_tester_settings.json`.
- The file is auto-created with 40 RS232 entries and 8 RS485 pair entries on first run.
- Settings are validated and normalized when loaded.
- Fullscreen default lives in `ui.start_fullscreen`.
- Auto-start-after-launch default lives in `ui.auto_start_after_launch_2s`.
- 2-second startup delay default lives in `ui.delay_comm_start_2s`.
- Overview compact mode default lives in `ui.overview_compact_view`.
- Preset names and COM selections live in `ui.presets`.
- Port names are not forced to be unique/valid, so you can stage configs on systems with fewer COM ports.
