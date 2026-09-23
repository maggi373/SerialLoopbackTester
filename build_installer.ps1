param(
    [switch]$SkipInno
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$appVersion = "1.3.0"
$portableBaseName = "SerialLoopbackTester-v$appVersion-portable"
$portableZipName = "$portableBaseName.zip"
$installerBaseName = "SerialLoopbackTester-v$appVersion-installer"
$legacyPortableExePath = Join-Path $PSScriptRoot ("dist\\{0}.exe" -f $portableBaseName)

if (Test-Path $legacyPortableExePath) {
    Remove-Item -LiteralPath $legacyPortableExePath -Force
}

Write-Host "Installing Python dependencies..."
python -m pip install --upgrade pip
python -m pip install -r requirements.txt -r requirements-build.txt

Write-Host "Running tests..."
python -m unittest discover -s tests

Write-Host "Building EXE with PyInstaller..."
python -m PyInstaller `
    --noconfirm `
    --clean `
    --noupx `
    --onedir `
    --windowed `
    --hidden-import serial.urlhandler.protocol_socket `
    --name $portableBaseName `
    serial_tester_gui.py

$portableDirPath = Join-Path $PSScriptRoot ("dist\\{0}" -f $portableBaseName)
$exePath = Join-Path $portableDirPath ("{0}.exe" -f $portableBaseName)
$portableZipPath = Join-Path $PSScriptRoot ("dist\\{0}" -f $portableZipName)

if (-not (Test-Path $exePath)) {
    throw "Build failed: folder-mode EXE was not created at $exePath"
}

Write-Host "Creating inspectable portable ZIP..."
$zipCreated = $false
for ($attempt = 1; $attempt -le 10; $attempt++) {
    try {
        Compress-Archive `
            -Path (Join-Path $portableDirPath "*") `
            -DestinationPath $portableZipPath `
            -CompressionLevel Optimal `
            -Force `
            -ErrorAction Stop
        $zipCreated = $true
        break
    } catch {
        if ($attempt -eq 10) {
            throw
        }
        Write-Warning "Portable files are temporarily locked; retrying ZIP creation ($attempt/10)..."
        Start-Sleep -Milliseconds 750
    }
}

if (-not $zipCreated) {
    throw "Build failed: portable ZIP was not created at $portableZipPath"
}

Write-Host "Portable folder ready: $portableDirPath"
Write-Host "Portable ZIP ready: $portableZipPath"

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
    Write-Host "Portable folder and ZIP builds are complete and usable."
    exit 0
}

Write-Host "Building Setup installer with Inno Setup..."
& $iscc "/DMyAppVersion=$appVersion" "/DMyAppExeBaseName=$portableBaseName" "/DMyOutputBaseFilename=$installerBaseName" "installer\\serial_loopback_tester.iss"

Write-Host "Installer build complete. Check dist\\installer."
