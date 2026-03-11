#ifndef MyAppName
#define MyAppName "Serial Loopback Tester"
#endif
#ifndef MyAppVersion
#define MyAppVersion "1.0.0"
#endif
#ifndef MyAppPublisher
#define MyAppPublisher "PoldenTEK"
#endif
#ifndef MyAppExeBaseName
#define MyAppExeBaseName "SerialLoopbackTester-portable"
#endif
#define MyAppExeName "{#MyAppExeBaseName}.exe"
#ifndef MyOutputBaseFilename
#define MyOutputBaseFilename "SerialLoopbackTester-installer"
#endif

[Setup]
AppId={{F85F50B8-B490-49AF-84D8-198A8D8D478A}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Serial Loopback Tester
DefaultGroupName=Serial Loopback Tester
OutputDir=..\dist\installer
OutputBaseFilename={#MyOutputBaseFilename}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
#ifexist "..\dist\{#MyAppExeName}"
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
#else
Source: "..\dist\{#MyAppExeBaseName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
#endif
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Serial Loopback Tester"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Serial Loopback Tester"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Serial Loopback Tester"; Flags: nowait postinstall skipifsilent
