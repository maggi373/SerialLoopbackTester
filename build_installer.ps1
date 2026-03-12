param(
    [switch]$SkipInno
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$appVersion = "1.1.0"
$portableBaseName = "SerialLoopbackTester-v$appVersion-portable"
$installerBaseName = "SerialLoopbackTester-v$appVersion-installer"

Write-Host "Installing Python dependencies..."
python -m pip install --upgrade pip
python -m pip install -r requirements.txt -r requirements-build.txt

Write-Host "Building EXE with PyInstaller..."
python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name $portableBaseName `
    serial_tester_gui.py

$oneFileExePath = Join-Path $PSScriptRoot ("dist\\{0}.exe" -f $portableBaseName)
$oneDirExePath = Join-Path $PSScriptRoot ("dist\\{0}\\{0}.exe" -f $portableBaseName)

if (Test-Path $oneFileExePath) {
    $exePath = $oneFileExePath
} elseif (Test-Path $oneDirExePath) {
    Copy-Item -Path $oneDirExePath -Destination $oneFileExePath -Force
    $exePath = $oneFileExePath
} else {
    throw "Build failed: EXE was not created in dist\\ (expected $oneFileExePath or $oneDirExePath)"
}

Write-Host "Portable EXE ready: $exePath"

if ($SkipInno) {
    Write-Host "Skipping installer packaging because -SkipInno was supplied."
    exit 0
}

$iscc = (Get-Command ISCC.exe -ErrorAction SilentlyContinue).Source
if (-not $iscc) {
    $localAppDataInno = $null
    if ($env:LOCALAPPDATA) {
        $localAppDataInno = Join-Path $env:LOCALAPPDATA "Programs\\Inno Setup 6\\ISCC.exe"
    }
    $candidates = @(
        $localAppDataInno,
        "C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe",
        "C:\\Program Files\\Inno Setup 6\\ISCC.exe"
    ) | Where-Object { $_ }
    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            $iscc = $candidate
            break
        }
    }
}

if (-not $iscc) {
    Write-Warning "Inno Setup was not found. Install Inno Setup 6 to build a Setup installer."
    Write-Host "EXE build is complete and usable."
    exit 0
}

Write-Host "Building Setup installer with Inno Setup..."
& $iscc "/DMyAppVersion=$appVersion" "/DMyAppExeBaseName=$portableBaseName" "/DMyOutputBaseFilename=$installerBaseName" "installer\\serial_loopback_tester.iss"

Write-Host "Installer build complete. Check dist\\installer."
