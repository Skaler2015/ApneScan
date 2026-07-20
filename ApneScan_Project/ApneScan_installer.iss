; ===================================================================
;  ApneScan Installer script (Inno Setup)  —  ab Tesseract OCR bhi
;  isi installer ke saath APNE-AAP (silently) install ho jayega.
;
;  Kaise use karein:
;   1. Inno Setup install karo (free): https://jrsoftware.org/isdl.php
;   2. Pehle apni .exe bana lo (build wale PC par):
;      py -3.12-32 -m PyInstaller --noconfirm --onefile --windowed ^
;         --name "ApneScan" --icon apnescan.ico --add-data "apnescan.ico;." ^
;         --collect-all win32com --collect-all win32 ^
;         --hidden-import pythoncom --hidden-import pywintypes scanner_app_v14.py
;   3. In 3 file ko is .iss ke SAATH-WALE folder me rakho:
;         dist\ApneScan.exe
;         apnescan.ico
;         tesseract-ocr-w64-setup-5_5_0_20241111.exe   <-- OCR installer
;   4. Is .iss ko Inno Setup me kholo aur "Compile" dabao.
;   5. Ek "ApneScan-Setup.exe" banega. Use koi bhi double-click karke
;      install kare -> ApneScan + Tesseract dono lag jayenge.
;
;  NOTE: Tesseract silent install ke liye admin rights chahiye (Program
;  Files me jata hai). Inno Setup default me admin maangta hai, to theek hai.
; ===================================================================

#define MyAppName "ApneScan"
#define MyAppVersion "1.0"
#define MyAppPublisher "Noble Care Hospital"
#define MyAppExeName "ApneScan.exe"
#define TesseractSetup "tesseract-ocr-w64-setup-5_5_0_20241111.exe"

[Setup]
AppId={{9F3B2A10-APNE-SCAN-0001-NOBLECARE0001}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=ApneScan-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=apnescan.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
; Tesseract Program Files me jaata hai -> admin chahiye
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Desktop par icon banayein"; GroupDescription: "Shortcuts:"

[Files]
; Ye teen file compile karne se pehle isi folder me honi chahiye:
Source: "dist\ApneScan.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "apnescan.ico";      DestDir: "{app}"; Flags: ignoreversion
; Tesseract OCR installer — temp me nikaalo, silent chalao, phir hata do:
Source: "{#TesseractSetup}"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
Name: "{group}\{#MyAppName}";            Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\apnescan.ico"
Name: "{group}\Uninstall {#MyAppName}";  Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";      Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\apnescan.ico"; Tasks: desktopicon

[Run]
; 1) Pehle Tesseract silently install karo — SIRF tab jab pehle se na ho.
Filename: "{tmp}\{#TesseractSetup}"; Parameters: "/S"; \
  StatusMsg: "OCR engine (Tesseract) install ho raha hai... thoda ruko"; \
  Check: TesseractNotInstalled; Flags: waituntilterminated

; 2) Phir ApneScan chalane ka option (install ke ant me).
Filename: "{app}\{#MyAppExeName}"; Description: "ApneScan abhi chalayein"; \
  Flags: nowait postinstall skipifsilent

[Code]
{ Agar Tesseract pehle se install hai to dobara mat chalao. }
function TesseractNotInstalled(): Boolean;
begin
  Result := not (
    FileExists(ExpandConstant('{commonpf}\Tesseract-OCR\tesseract.exe')) or
    FileExists(ExpandConstant('{commonpf32}\Tesseract-OCR\tesseract.exe'))
  );
end;
