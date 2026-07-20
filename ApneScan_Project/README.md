# ApneScan — Document Scanner (Noble Care Hospital)

ApneScan ek Windows document-scanning software hai (Python + PyQt5) jo HP ScanJet
Pro N4000 snw1 (network/eSCL) ke liye banaya gaya hai — hospital ke ECHS/RGHS
claims ke liye aur public release ke liye.

---

## Folder me kya hai

| File | Kya hai |
|------|---------|
| `apnescan.py` | **MAIN code** (yahi build karna hai). TWAIN file-transfer (continuous feed) ab isi me hai — Settings me "TWAIN continuous feed (experimental)" se on hota hai. |
| `apnescan.ico` | App ka icon (build me `--icon` ke liye). |
| `build.bat` | **Double-click** karke .exe bana sakte ho (aasaan tareeka). |
| `requirements.txt` | Zaroori Python libraries. |
| `ApneScan_installer.iss` | Inno Setup installer script (ApneScan + Tesseract dono install karta hai). |
| `ApneScan_Stats.gs` | Google Apps Script — worldwide stats ka free server. |
| `apnescan_website.html` | Product website (host karne ke liye). |
| `apnescan_logo.png/svg`, `apnescan_icon.png` | Branding. |

---

## Zaroorat (ek baar setup)

1. **32-bit Python 3.12** install karo (HP scanner drivers 32-bit hain):
   https://www.python.org/downloads/  → install me "Add to PATH" ✓
2. Libraries install karo:
   ```
   py -3.12-32 -m pip install -r requirements.txt
   ```
3. **Tesseract OCR** (document ke naam padhne ke liye, optional):
   https://github.com/UB-Mannheim/tesseract/wiki
4. **Scanner** wahi WiFi/LAN par ho. Scanner IP: Settings → Scanner IP (jaise 192.168.1.8).

---

## Chalaana (bina .exe banaye, test ke liye)

```
py -3.12-32 apnescan.py
```

## .exe banaana (deploy ke liye)

Aasaan: **`build.bat` par double-click** karo.

Ya command se:
```
py -3.12-32 -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "ApneScan" --icon apnescan.ico ^
  --collect-all win32com --collect-all win32 ^
  --hidden-import pythoncom --hidden-import pywintypes apnescan.py
```
`.exe` `dist\ApneScan.exe` me banegi.

> Tip: har build ke liye naya `--name` do (jaise ApneScanV24) taaki purani .exe
> chalu hone par "Access denied" na aaye.

---

## Installer banaana (ApneScan + Tesseract dono ek saath)

1. Inno Setup install karo (free): https://jrsoftware.org/isdl.php
2. In 3 file ko `ApneScan_installer.iss` ke saath rakho: `dist\ApneScan.exe`,
   `apnescan.ico`, aur `tesseract-ocr-w64-setup-5_5_0_20241111.exe`.
3. `.iss` ko Inno Setup me kholo → Compile → `ApneScan-Setup.exe` ban jayegi.

---

## GitHub se automatic .exe build (naya!)

Ab har `git push` par GitHub Actions ka **Windows server** khud .exe bana deta hai:
- Har push: https://github.com/Skaler2015/ApneScan/actions → run kholo → neeche
  **Artifacts** me `ApneScan-exe` (ZIP me ApneScan.exe).
- **Release nikalna ho** (public download link ke liye): ek tag push karo, jaise
  `git tag v17 && git push origin v17` → exe apne aap
  https://github.com/Skaler2015/ApneScan/releases par aa jayegi.
- App me **Help → Update check karo** isi Releases page se naya version batata hai
  (startup par bhi chupchaap check hota hai).

---

## Worldwide Stats (already live)

`ApneScan_Stats.gs` Google Apps Script par deploy hai; app me URL already daala hua
hai. Sidebar me total/aaj/online dikhta hai. Sirf GINTI jaati hai — koi document ya
patient data kabhi nahi.

> **Zaroori (July 2026 fix):** "Aaj (today): 0" wala bug .gs code me fix hua hai.
> Naya code lagane ke liye: Google Sheet → Extensions → Apps Script → poora code
> is file ke naye code se REPLACE karo → Save → **Deploy → Manage deployments →
> (pencil) Edit → Version: New version → Deploy.** URL wahi rahega, app me kuch
> nahi badalna.

---

## Main features (short)

- Network duplex scan (eSCL, bina kisi aur software ke)
- Fast mode (200dpi + B&W), page-size Auto/A4/Letter/Legal/A5 (Auto me kuch nahi katta)
- Blank page hatao, black backing → white (ECG jaise pages)
- Document ka naam OCR se (learning: F2 se naam sikhao, agli baar khud lag jayega)
- Save PDF (all/selected/password), Save Images, Print (All/Selected/ID 2-up)
- NAPS2-jaisa preview (zoom, rotate, rename, delete)
- Profiles (per-profile page size), keyboard shortcuts (Scan=Enter, Save selected=Space)
- Hindi/English, light/dark theme, bilingual help everywhere
- Excel register, backup, barcode claim autofill, merge/split PDF, search old PDFs
  (naam se AUR PDF ke andar ke text se)
- **Share**: saved PDF ko WhatsApp/Email se bhejo (File → Share/Bhejo)
- **PDF compress tool**: kisi bhi PDF ko 200KB/500KB/1MB/2MB tak chhota karo
  (Tools → PDF chhota karo) — portal upload limits ke liye
- **Phone-photo se PDF**: phone ki document-photo import karo — shadow hat kar
  scan-jaisi saaf (Tools → Phone-photo se PDF)
- **ID cards alag karo**: ek page par 2-3 cards scan karke unhe alag-alag pages
  me kaato (Tools menu)
- **Update check**: Help → Update check karo (GitHub Releases se)
- **TWAIN continuous feed (experimental)**: Settings me on karo to ADF bina ruke
  chalta hai (support na ho to khud purane tareeke par aa jata hai)

---

## Aage kaam kaise continue karein

- Code kisi bhi editor (VS Code) me kholo — `apnescan.py` single file hai.
- Naye chat me kaam continue karna ho to `apnescan.py` upload kar dena; poora context
  usi file me hai.
- Har change ke baad `build.bat` se nayi .exe bana lena.

Backlog (jo abhi banana baaki hai): installer ko bhi CI me banana, website ko
GitHub Pages par host karna, code ko chhote modules me baantna, image-processing
functions ke automatic tests.
