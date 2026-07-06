[Setup]
AppName=Disp Application
AppVersion=1.0
DefaultDirName={autopf}\Disp
DefaultGroupName=Disp
OutputDir=installer
OutputBaseFilename=Disp_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=logo2.ico
UninstallDisplayIcon={app}\Disp.exe
PrivilegesRequired=lowest
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1

[Files]
Source: "dist\Disp.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "logo2.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "app.db"; DestDir: "{app}"; Flags: ignoreversion
Source: "people.db"; DestDir: "{app}"; Flags: ignoreversion
Source: "app_2025_*.db"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Disp"; Filename: "{app}\Disp.exe"; IconFilename: "{app}\logo2.ico"
Name: "{group}\{cm:UninstallProgram,Disp}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Disp"; Filename: "{app}\Disp.exe"; IconFilename: "{app}\logo2.ico"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\Disp"; Filename: "{app}\Disp.exe"; IconFilename: "{app}\logo2.ico"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\Disp.exe"; Description: "{cm:LaunchProgram,Disp}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\app_2025_*.db"

