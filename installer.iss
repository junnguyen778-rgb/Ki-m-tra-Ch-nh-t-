; Inno Setup script - dong goi ViSpellCheck.exe thanh ViSpellCheck-Setup.exe
; Yeu cau: da build dist\ViSpellCheck.exe (xem build_exe.bat)
#define AppName "ViSpellCheck Pro"
#define AppExe  "ViSpellCheck.exe"
#define AppVer  "2.0.0"

[Setup]
AppId={{8F3B2C1E-5A47-4D9B-9C0E-6B7A1D2E4F10}
AppName={#AppName}
AppVersion={#AppVer}
DefaultDirName={autopf}\ViSpellCheck
DefaultGroupName={#AppName}
OutputDir=dist
OutputBaseFilename=ViSpellCheck-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern
UninstallDisplayIcon={app}\{#AppExe}
ArchitecturesInstallIn64BitMode=x64compatible

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Shortcuts:"

[Files]
Source: "dist\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
