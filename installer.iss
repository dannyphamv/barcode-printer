[Setup]
AppName=Legacy Barcode Printer
AppVersion=1.0
AppVerName=Legacy Barcode Printer 1.0
AppPublisher=Danny Pham
AppPublisherURL=https://github.com/dannyphamv/barcode-printer
AppCopyright=Copyright (C) 2026 Danny Pham
PrivilegesRequired=lowest
CloseApplications=yes
CloseApplicationsFilter=LegacyBarcodePrinter.exe
RestartApplications=yes
LicenseFile=LICENSE.txt
AppId={{d0e7e2e6-c6be-434a-927c-a7a105c5f12a}}
AppComments=A Windows application using Tkinter for printing Code128 barcodes to any connected printer

; Install paths
DefaultDirName={autopf}\Legacy Barcode Printer
DefaultGroupName=Legacy Barcode Printer
DisableProgramGroupPage=yes

; Output
OutputDir=installer_output
OutputBaseFilename=LegacyBarcodePrinter_Setup
SetupIconFile=favicon.ico

; Compression
Compression=lzma
SolidCompression=yes

; Minimum Windows version (Windows 10)
MinVersion=10.0

; Uninstall
UninstallDisplayIcon={app}\LegacyBarcodePrinter.exe
UninstallDisplayName=Legacy Barcode Printer

; Prevents multiple instances of the installer running
AppMutex=LegacyBarcodePrinterSetupMutex

; Version info shown in installer EXE properties
VersionInfoVersion=1.0
VersionInfoCompany=Danny Pham
VersionInfoDescription=Legacy Barcode Printer Installer
VersionInfoProductName=Legacy Barcode Printer
VersionInfoProductVersion=1.0

; 64-bit install
ArchitecturesInstallIn64BitMode=x64compatible

; Modern wizard UI
WizardStyle=dynamic

[Code]
function InitializeSetup(): Boolean;
begin
  if not IsWin64 then
  begin
    MsgBox('Legacy Barcode Printer requires a 64-bit version of Windows.', mbError, MB_OK);
    Result := False;
  end else
    Result := True;
end;

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"
Name: "startupicon"; Description: "Start automatically with Windows"; GroupDescription: "Startup:"; Flags: unchecked

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "LegacyBarcodePrinter"; ValueData: """{app}\LegacyBarcodePrinter.exe"""; Tasks: startupicon; Flags: uninsdeletevalue

[Files]
Source: "dist\LegacyBarcodePrinter\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Legacy Barcode Printer"; Filename: "{app}\LegacyBarcodePrinter.exe"
Name: "{group}\Uninstall Legacy Barcode Printer"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Legacy Barcode Printer"; Filename: "{app}\LegacyBarcodePrinter.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\LegacyBarcodePrinter.exe"; Description: "Launch Legacy Barcode Printer"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\logs"
Type: filesandordirs; Name: "{app}\cache"
Type: filesandordirs; Name: "{userappdata}\LegacyBarcodePrinter"