; PWQInstaller.iss - Inno Setup script for PWQ standalone installer
[Setup]
AppName=PWQ
AppVersion=1.0.4
DefaultDirName={pf}\PWQ
DefaultGroupName=PWQ
OutputDir=.
OutputBaseFilename=PWQInstaller
Compression=lzma
SolidCompression=yes
SetupIconFile=C:\Users\sanjayad\MyApps\PWQ\PWQ_logo.ico

[Files]
Source: "C:\\Users\\sanjayad\\MyApps\\PWQ\\bin\\Release\\net6.0-windows\\win-x64\\publish\\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\PWQ"; Filename: "{app}\PWQ.exe"
Name: "{userdesktop}\PWQ"; Filename: "{app}\PWQ.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"

[Run]
Filename: "{app}\PWQ.exe"; Description: "Launch PWQ"; Flags: nowait postinstall skipifsilent
