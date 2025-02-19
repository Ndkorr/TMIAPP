[Setup]
AppName=CF Viewer
AppVersion=1.0
DefaultDirName={pf}\CF Viewer
DefaultGroupName=CF Viewer
OutputDir=Output
OutputBaseFilename=CFViewer_setup
Compression=lzma
SolidCompression=yes

[Files]
; Copy your executable and any additional files
Source: "C:\Users\HR-IT-MATTHEW-PC\output\Custom File Viewer\Custom File Viewer.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\My Application"; Filename: "{app}\MyApp.exe"
Name: "{commondesktop}\My Application"; Filename: "{app}\MyApp.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"
