; ApneScan installer (CI version) — GitHub Actions par apne aap banta hai.
; (Tesseract alag se install hota hai; ye sirf app install karta hai.)

[Setup]
; Fixed AppId => har naya per-user version PURANI jagah par hi upgrade hota
; hai (do-do copy nahi bante), aur future updates saaf replace hote hain.
AppId={{8F3A1C2E-5B7D-4E9A-9C1F-2A6B8D4E7F03}
AppName=ApneScan
AppVersion=326
AppPublisher=ApneSoft
AppPublisherURL=https://apnescan.apnesoft.com
; (v268) PER-USER install (LocalAppData) — Administrator/UAC ki zaroorat NAHI.
; Isse auto-update ke waqt UAC prompt bilkul nahi aata (Chrome/VS Code jaisa),
; app khud ko chupchaap update kar leti hai.
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\ApneScan
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
; (v280) ONEDIR — poora dist\ApneScan folder (ApneScan.exe + _internal\ me
; python312.dll aur baaki DLL). Runtime par kuch extract nahi hota, isliye
; 'Failed to load python312.dll' error kabhi nahi aata.
Source: "dist\ApneScan\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
Source: "apnescan.ico"; DestDir: "{app}"
; (v205) Tesseract OCR BUNDLED — auto-rotate/OCR ke liye alag install nahi
Source: "tesseract_dist\*"; DestDir: "{app}\tesseract"; Flags: recursesubdirs ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\ApneScan"; Filename: "{app}\ApneScan.exe"
Name: "{autodesktop}\ApneScan"; Filename: "{app}\ApneScan.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ApneScan.exe"; Description: "ApneScan chalao"; Flags: nowait postinstall skipifsilent
