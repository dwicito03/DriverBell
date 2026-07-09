; ================================================
; DriverBell Installer (clean, onedir build)
; ================================================

#define MyAppName "DriverBell"
#define MyAppVersion "1.7.2"
#define MyAppPublisher "dwicito.com"
#define MyAppURL "https://dwicito.com"
#define MyAppExeName "DriverBell.exe"
#define MyAppReleaseDate "2026-07-13"

[Setup]
AppId={{A88B83C4-9DAF-4EE8-BA1F-DBE821D1E9A2}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\DriverBell
DefaultGroupName=dwicito.com
DisableProgramGroupPage=no
OutputDir=installerOutput
OutputBaseFilename=DriverBellSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=assets\icons\driverbell_brand.ico

; Pastikan file lisensi ini ada di folder yang sama
LicenseFile=license_en.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Pakai build onedir: salin seluruh isi dist\DriverBell ke folder instalasi
Source: "dist\DriverBell\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "CHANGELOG.txt"; DestDir: "{app}"; Flags: ignoreversion
; Padding terkontrol untuk menyesuaikan ukuran installer (tidak diinstal ke user)
Source: "installer\sizepad.bin"; DestDir: "{tmp}"; Flags: dontcopy nocompression

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\CHANGELOG.txt"; Description: "View What's New"; Flags: shellexec postinstall skipifsilent unchecked

[Code]
function InitializeSetup(): Boolean;
begin
  { Tidak ada lagi pengecekan VLC; selalu lanjut instalasi }
  Result := True;
end;


