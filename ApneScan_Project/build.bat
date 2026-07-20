@echo off
REM ==== ApneScan .exe banane ka one-click script ====
REM Ise apnescan.py ke SAATH-WALE folder me rakho, phir double-click karo.
cd /d "%~dp0"
echo Building ApneScan.exe ... (1-2 minute)
py -3.12-32 -m PyInstaller --noconfirm --clean --onefile --windowed --name "ApneScan" --icon apnescan.ico --collect-all win32com --collect-all win32 --hidden-import pythoncom --hidden-import pywintypes apnescan.py
echo.
if exist "dist\ApneScan.exe" (
  echo DONE!  dist\ApneScan.exe ban gayi.
) else (
  echo Build FAIL - upar ka error dekho. Python 3.12-32 + libraries install hain?
)
pause
