; ApneScan installer (CI version) — GitHub Actions par apne aap banta hai.
; (Tesseract alag se install hota hai; ye sirf app install karta hai.)

[Setup]
AppName=ApneScan
AppVersion=171
AppPublisher=ApneSoft
AppPublisherURL=https://apnescan.apnesoft.com
DefaultDirName={autopf}\ApneScan
DefaultGroupName=ApneScan
OutputDir=Output
OutputBaseFilename=ApneScan-Setup
Compression=lzma2
SolidCompression=yes
SetupIconFile=apnescan.ico
DisableProgramGroupPage=yes

[Tasks]
Name: "desktopicon"; Description: "Desktop par icon banao"; GroupDescription: "Shortcuts:"

[Files]
Source: "dist\ApneScan.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "apnescan.ico"; DestDir: "{app}"

[Icons]
Name: "{autoprograms}\ApneScan"; Filename: "{app}\ApneScan.exe"
Name: "{autodesktop}\ApneScan"; Filename: "{app}\ApneScan.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ApneScan.exe"; Description: "ApneScan chalao"; Flags: nowait postinstall skipifsilent
