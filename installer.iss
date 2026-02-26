; Inno Setup Script for RLA Playlist Creator
; Build with: iscc installer.iss

#define MyAppName "RLA Playlist Creator"
#define MyAppVersion "2.0"
#define MyAppPublisher "RLA"
#define MyAppExeName "RLA_PlaylistCreator.exe"

[Setup]
AppId={{B7A3F2D1-4E5C-4A8B-9D6F-1C2E3A4B5D6E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\RLA_PlaylistCreator
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=dist
OutputBaseFilename=RLA_PlaylistCreator_Setup
SetupIconFile=RLA-PlaylistCreator.ico
UninstallDisplayIcon={app}\RLA_PlaylistCreator.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=force

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\RLA_PlaylistCreator\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\{#MyAppExeName}"; Flags: nowait postinstall skipifdontexist silent
