# -*- coding: utf-8 -*-
"""
ApneScan - v18  [31 naye public features: bundle auto-split, auto-orient, auto-colour,
                 scanner discovery, book/business-card/photo-restore modes, PDF tools
                 (editor, sign/stamp, Word/Excel/JPG, watermark, unlock, archival),
                 tags + instant OCR-index search, self-installing auto-update,
                 portable mode, settings export/import, crash reporter, touch mode,
                 4 nayi bhashayein (partial)]
=========================================================
Run with 32-bit Python (HP TWAIN driver is 32-bit):
    py -3.12-32 scanner_app_v10.py
One-click scan shortcut:
    py -3.12-32 scanner_app_v10.py --scan --profile "Documents"
"""

import io
import os
import re
import sys
import json
import shutil
import socket
import tempfile
import traceback
import datetime

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog

try:
    from PyQt5.QtMultimedia import QCamera, QCameraInfo, QCameraImageCapture
    from PyQt5.QtMultimediaWidgets import QCameraViewfinder
    HAS_CAMERA = True
except Exception:
    HAS_CAMERA = False

try:
    import twain
    HAS_TWAIN = True
except Exception:
    HAS_TWAIN = False

from PIL import Image, ImageChops, ImageEnhance, ImageOps, ImageDraw, ImageFont, ImageFilter

try:
    import numpy as np
    HAS_NUMPY = True
except Exception:
    HAS_NUMPY = False

try:
    import pytesseract
    from pypdf import PdfReader, PdfWriter
    HAS_OCR_LIBS = True
except Exception:
    HAS_OCR_LIBS = False

try:
    import fitz  # PyMuPDF — best PDF page rendering (optional)
    HAS_FITZ = True
except Exception:
    HAS_FITZ = False


def pdf_to_images(pdf_path, tmpdir, dpi=200):
    """Turn each page of a PDF into a PNG (for import). Uses PyMuPDF if installed
    (renders any PDF), else falls back to extracting embedded images with pypdf
    (works for scanned PDFs). Returns a list of PNG paths."""
    out = []
    if HAS_FITZ:
        try:
            doc = fitz.open(pdf_path)
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            for page in doc:
                pix = page.get_pixmap(matrix=mat)
                fd, png = tempfile.mkstemp(suffix=".png", dir=tmpdir); os.close(fd)
                pix.save(png)
                out.append(png)
            doc.close()
            if out:
                return out
        except Exception:
            out = []
    # fallback: pull the biggest embedded image from each page (scanned PDFs)
    try:
        from pypdf import PdfReader as _PR
        reader = _PR(pdf_path)
        for page in reader.pages:
            try:
                imgs = list(getattr(page, "images", []))
            except Exception:
                imgs = []
            if not imgs:
                continue
            big = max(imgs, key=lambda im: len(getattr(im, "data", b"")))
            fd, png = tempfile.mkstemp(suffix=".png", dir=tmpdir); os.close(fd)
            with open(png, "wb") as fh:
                fh.write(big.data)
            try:
                with Image.open(png) as im2:
                    im2.convert("RGB").save(png, "PNG")
            except Exception:
                pass
            out.append(png)
    except Exception:
        pass
    return out


def _configure_tesseract():
    """Point pytesseract at a Tesseract install even if it's not on PATH."""
    if not HAS_OCR_LIBS:
        return
    try:
        home = os.path.expanduser("~")
        for c in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                  r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                  os.path.join(home, "AppData", "Local", "Tesseract-OCR", "tesseract.exe"),
                  os.path.join(home, "AppData", "Local", "Programs", "Tesseract-OCR", "tesseract.exe")):
            if os.path.exists(c):
                pytesseract.pytesseract.tesseract_cmd = c
                break
    except Exception:
        pass


def tesseract_available():
    if not HAS_OCR_LIBS:
        return False
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


_configure_tesseract()

try:
    import openpyxl
    HAS_XLSX = True
except Exception:
    HAS_XLSX = False

try:
    from pyzbar.pyzbar import decode as _zbar_decode
    HAS_ZBAR = True
except Exception:
    HAS_ZBAR = False

try:
    import win32com.client as _w32
    import pythoncom as _pythoncom
    HAS_W32 = True
except Exception:
    HAS_W32 = False


APP_NAME = "ApneScan"
VERSION = "29"
UPDATE_API = "https://api.github.com/repos/Skaler2015/ApneScan/releases/latest"
DOWNLOAD_PAGE = "https://github.com/Skaler2015/ApneScan/releases/latest"
def _portable_dir():
    """PORTABLE MODE: exe ke saath 'portable.txt' naam ki khaali file rakh do —
    saari settings wahi folder me rahengi (pen-drive se chalao, settings saath)."""
    try:
        base = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                               else os.path.abspath(__file__))
        if os.path.exists(os.path.join(base, "portable.txt")):
            return base
    except Exception:
        pass
    return None


_PORTABLE = _portable_dir()
CONFIG_PATH = (os.path.join(_PORTABLE, "apnescan_config.json") if _PORTABLE
               else os.path.join(os.path.expanduser("~"), ".apnescan.json"))
CRASH_PATH = os.path.join(os.path.expanduser("~"), "apnescan_crash.txt")
PSTATS_PATH = os.path.join(os.path.expanduser("~"), ".apnescan_pstats.json")
_OLD_CONFIG = os.path.join(os.path.expanduser("~"), ".noble_doc_scanner.json")
if not os.path.exists(CONFIG_PATH) and os.path.exists(_OLD_CONFIG):
    try:
        import shutil as _sh
        _sh.copy2(_OLD_CONFIG, CONFIG_PATH)
    except Exception:
        pass
SCANNER_PORTS = (80, 443, 8080, 9100)

# Built-in ApneScan worldwide-stats endpoint (Google Apps Script). Every install
# reports scan COUNTS here (never any document/patient data).
DEFAULT_STATS_URL = "https://script.google.com/macros/s/AKfycbzwh2HQHXKe_09wgXpSh-NZQu-GbR7N-NVmtUEyRpl8oOa-MdoJ9ShqHB_i0QCcuqe2_w/exec"

_TESS_GUESSES = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]
if HAS_OCR_LIBS:
    for _p in _TESS_GUESSES:
        if os.path.exists(_p):
            pytesseract.pytesseract.tesseract_cmd = _p
            break

def resource_path(name):
    """Path to a bundled data file, works both from source and PyInstaller .exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))
    p = os.path.join(base, name)
    if os.path.exists(p):
        return p
    return os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), name)


# Embedded app icon (base64 PNG) so the ApneScan icon always shows in the title
# bar + taskbar even if the icon file wasn't bundled with the build.
_EMBEDDED_ICON_B64 = """\
iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAYAAAD0eNT6AAAQoklEQVR4nO3dS48c5d3G4Rqrl0gzlpA/iXdss3bWwQokMZFickAG
EnbhsCOBOOTEJuEYJ+v4O7DjkyAktyX2k4XduMeemT5V1XO4r0tqvYr0ClU/7q7/r57q6T4aqMrx735zWvoYAKbw8A9/OSp9DDzh
H2Nmx7/9tQEPcI6Hf/yrmTQjiz0hwx7gMKJgOhZ2RAY+wLQEwXgs5AEMfICyBMH+LNyODH2AOomB3VisLRj6AG0RA5tZoEsY/ABt
EwIXszBPMfQB+iQGzrIYjxn8ABmEwCPxi2DwA2RKD4HYJ2/wAzAMuSEQ96QNfgDOkxYCMU/W4AdgGykh0P2TNPgB2EfvIdDtkzP4
ARhDryFwpfQBTMHwB2Asvc6Urqqm138kAOrQ025AF0/E4AdgTj2EQPO3AAx/AObWw+xpOgB6+AcAoE2tz6AmtzBaX3QA+tLiLYHm
Dvj4TcMfgPo8/KCtCGjqFoDhD0CtWptRTdRKa4sKQLYWdgOqP8DjN39l+APQnIcf/K3qGVv1LQDDH4BW1T7Dqg2A2hcOADapeZZV
GQA1LxgA7KLWmVZdANS6UACwrxpnWzUfUKhxcQBgbLV8OLCKHQDDH4AUtcy84gFQy0IAwFxqmH1FA6CGBQCAEkrPwGIBUPqJA0Bp
JWdh8VsAAMD8igSAq38AeKTUTJw9AAx/ADirxGycNQAMfwA439wzcrYAMPwB4HJzzspZAsDwB4DtzDUzJw8Awx8AdjPH7PRngAAQ
aNIfJDh+w9U/AOzr4YfT/XDQZDsAhj8AHGbKWTpJABj+ADCOqWaqzwAAQKDRA8DVPwCMa4rZOmoAGP4AMI2xZ+xony48fuOXhj8A
TOzhh38fZXb7DAAABBolAFz9A8A8xpq5dgAAINDBAeDqHwDmNcbsPSgADH8AKOPQGewWAAAE2jsAXP0DQFmHzGI7AAAQaK8AcPUP
AHXYdybbAQCAQDsHgKt/AKjLPrPZDgAABNopAFz9A0Cddp3RdgAAINDWAeDqHwDqtsusXmz9XzX+AaAbW+0AHL/u6h8AWrDtzPYZ
AAAItDEAXP0DQFu2md12AAAgkAAAgECXBoDtfwBo06YZbgcAAAJdGACu/gGgbZfN8ku+CMj8B4BeuQUAAIHODYDj1191+Q8AHbho
ptsBAIBAAgAAAj0TALb/AaAv5812OwAAEEgAAEAgAQAAgc4EgPv/ANCnp2e8HQAACCQAACDQ2d8CcAMAACJ8vwNwfMf9fwDo2fqs
dwsAAAIJAAAIJAAAIJAAAIBAV4bBBwABIMVq5tsBAIBAAgAAAgkAAAgkAAAgkAAAgECPfwvAHwEAQBI7AAAQ6Oj4zm2X/wAQxg4A
AAQSAAAQSAAAQCABAACBBAAABBIAABBo4TuAACCPHQAACCQAACCQAACAQAIAAAIJAAAIJAAAIJAAAIBAAgAAAgkAAAgkAAAgkAAA
gEACAAACLUofAKxb3v1H6UOgIV8/+Kb0IczmB++9U/oQ6IwAoDhDH2B+i8HvAVPI8u7HpQ8BGuJczbh8BoAiDH/GcP3qtdKHAM0S
AMzO8GdMIgD2IwCYleHPFEQA7E4AMBvDnymJANiNAGAWhj9zEAGwPQHA5Ax/5iQCYDsLf1nClJZ/NvyZ3/Wr1/r7kiDnakZmBwDo
kp0AuJwAYDKu/ilNBMDFBADQNREA5xMAQPdEADxLADAJ2//URgTAWQIAiCEC4AkBAEQRAfCIAADiiAAQAEAoEUA6AQDEEgEkEwBA
NBFAKgEAxBMBJBIAAIMIII8AAHhMBJBEAACsEQGkEAAATxEBJBAAAOcQAfRuMQynpY8BoErXr14bvn7wTenDeMy5mnHZAQC4hJ0A
eiUAADYQAfRoYVcJYLPitwOcqxmZHQCALdkJoCcCAGAHIoBeCACAHYkAeiAAAPYgAmidAADYkwigZQIA4AAigFYJAIADiQBaJAAA
RiACaI0AABiJCKAlAgBgRCKAVggAgJGJAFogAAAmIAKonQAAmIgIoGYCAGBCIoBaCQCAiYkAaiQAAGYgAqiNAACYiQigJgIAYEYi
gFoIAICZiQBqIAAAChABlLYYTk9LHwNApOtXrw1fP/hmu/9n52pGZgcAoCA7AZQiAAAKEwGUIAAAKiACmJsAAKiECGBOAgCgIiKA
uQgAgMqIAOYgAAAqJAKY2qL0AQBwPhHAlOwAAEAgAQAAgQQAAAQSAAAQSAAAQCABAACBBAAABBIAABBIAABAIAEAAIEEAAAEEgAA
EEgAAECgxXBa+hAA2Mi5mpHZAQCAQAIAAAIt7CsBtMC5mnHZAQCAQAIAAAIJAAAItCh9AFDKyVuv3Sx9DLRh+f5H90ofA4xNABDH
4GdXq9eMEKAnbgEQxfDnEF4/9EQAEMPJmzF4HdELAUAEJ23G5PVEDwQAAAQSAHTP1RpT8LqidQIAAAIJAAAIJAAAIJAAAIBAAoDu
+fY2puB1ResEAAAEEgBEcLXGmLye6MGV4XQYPDxGf1TISZsxFHsdlX5Pe3T3sANAFBHAIbx+6ImfAybO6iTum9zYlsFPjwQAsZzU
gWRuAQBAIAEAAIEWjz4OCEDdnKsZlx0AAAgkAAAgkAAAgEACAAACCQAACCQAACCQAACAQAIAAAIJAAAIJAAAIJAAAIBAAgAAAgkA
AAgkAAAg0MIvTAI0wLmakdkBAIBAAgAAAgkAAAgkAAAg0KL0AUApJ2+9drP0MdCG5fsf3St9DDA2AUAcg59drV4zQoCeuAVAFMOf
Q3j90BMBQAwnb8bgdUQvrjz6dgkPj7EfdXHSZkxlXk+l39MevT3sAABAIAFA91z9MwWvK1onAAAgkAAAgEACAAACHR2/euu09EFQ
hx++c+fb0scATO9/79x9vvQxUJ4ACGfoQzYxkMstgGCGP+A8kMsOQCBveOA8dgOy2AEIY/gDF3F+yCIAgnhzA5s4T+RYHJU+AgCq
Yi5kODq57TMACW68q+qB7d1/2+cBeucWQADDH9iV80b/FqUPgPm88NyJogc2+uq7peEfwA5ACMMf2JbzRQYB0Lkb79751psZ2NUL
z5087zZA3wRA5wx/YF/OH30TAAAQSAAAQCABAACBBAAABBIAnfP3vMC+nD/6JgA6d//tu897EwO7+uq75be+Drhvi2HwUwAJvvpu
6fsAgK08uWgwH3rmq4CD2AkAYMUtgAD33/6TK39gJ84b/RMAIbyZgW05X2QQAAAQSAAEUfXAJs4TOQRAGG9u4CLOD1mOTm7/zN95
hLrx7uv+KgAw+EPZAQjmTQ84D+Q6OvmFHQAeufGeHQFIcP/3hj4CgIk8+PhfpQ8BunL19q3Sh0BnfBMgsU7eeu1m6WPoyfL9j+6V
PgZgewKAOAb/NFbrKgSgDT4ESBTDf3rWGNogAIhhMM3HWkP9BAARDKT5WXOomwAAgEACgO65Ei3H2kO9BAAABBIAABBIAABAIAEA
AIEEAN3zzXTlWHuolwAAgEBXhuF08PAY/1EXV6Lzs+ZjK/2e9ujtYQeAGAbSfKw11E8AEMVgmp41hjb4OWDirAaUb6kbl8EPbREA
xDKwgGRuAQBAoMVwWvoQANjIuZqR2QEAgEACAAACCQAACCQAACCQAACAQAIAAAIJAAAIJAAAIJAAAIBAAgAAAgkAAAgkAAAgkAAA
gEACAAACCQAACCQAACCQAACAQIthOC19DABs5FzNuOwAAEAgAQAAgQQAAARauK0E0ADnakZmBwAAAgkAAAgkAAAgkAAAgEACAAAC
CQAACCQAACCQAACAQAIAAAIJAAAIJAAAIJAAAIBAAgAAAgkAAAgkAAAgkAAAgEACAAACCQAACCQAACCQAACAQIvh9LT0MQCwiXM1
I7MDAACBBAAABBIAABBIAABAIAEAAIEEAAAEEgAAEEgAAEAgAQAAgQQAAAQSAAAQSAAAQCABAACBBAAABBIAABBIAABAIAEAAIEE
AAAEEgAAEEgAAECgxXBa+hAA2Mi5mpHZAQCAQAIAAAIt7CsBtMC5mnHZAQCAQAIAAAIJAAAIJAAAIJAAAIBAAgAAAgkAAAgkAAAg
kAAAgEACAAACCQAACCQAACCQAGASV3/+culDgG54PzEFAQAAgQQAAAQSAEzGtiUczvuIqQgAAAgkAJiUqxfYn/cPUzo6ufXSaemD
oH8P/vl56UOAplx9xfBnWnYAmIWTGWzP+4U5CABm46QGm3mfMBcBwKyc3OBi3h/MSQAwOyc5eJb3BXMTABThZAdPeD9QwtHJrR/7
KwCKevDPL0ofAhRx9ZWXSh8CwQQAVRED9M7QpxYCAAAC+QwAAAQSAAAQSAAAQCABAACBBAAABBIAABBIAABAIAEAAIEEAAAEEgAA
EEgAAEAgAQAAgRaDnwICgDh2AAAgkAAAgEACAAACCQAACCQAACDQleUnXx6VPggAYD7LT748sgMAAIEEAAAEWjz6P74NCACS2AEA
gEACAAACCQAACCQAACDQlWEYhuUn//ZdAAAQYDXz7QAAQCABAACBBAAABBIAABDo+wDwQUAA6Nv6rLcDAACBFmf+l58EAIAIdgAA
INCZAFh+6nMAANCjp2e8HQAACCQAACCQAACAQM8EgM8BAEBfzpvtdgAAIJAAAIBA5waA2wAA0IeLZrodAAAIJAAAINClW/0nP73p
1wEAoFHLT+9dOOftAABAoEsD4LJyAADqtWmG2wEAgEACAAACbQwAtwEAoC3bzG47AAAQaKsAsAsAAG3YdmbbAQCAQFsHgF0AAKjb
LrN6sdN/2fcCAkAXdroFsPzMLgAA1GjXGe0zAAAQaOcAsAsAAHXZZzbbAQCAQHsFgF0AAKjDvjPZDgAABNo7AOwCAEBZh8xiOwAA
EOigALALAABlHDqDD94BEAEAMK8xZq9bAAAQaJQAsAsAAPMYa+baAQCAQKMFgF0AAJjWmLN29KF98pMX/WgwAIxs+dl/Rp3Zo98C
GPsAASDdFLPVZwAAINAkAWAXAADGMdVMnWwHQAQAwGGmnKWT3gIQAQCwn6ln6GLK//gwDMPgbwIAoDqTfwhw+bldAADYxRyzc5a/
AhABALCduWbmbH8GKAIA4HJzzspZvwdABADA+eaekbN/EZAIAICzSszGIt8EKAIA4JFSM9FXAQNAoGIBYBcAgHQlZ2HRHQARAECq
0jOw+C2A0gsAAHOrYfYVD4BhqGMhAGAOtcy8Kg5i3cnLL/r1AAC6U8vgX6liB2BdbQsEAIeqcbZVFwDDUOdCAcA+ap1pVQbAMNS7
YACwrZpnWbUBMAx1LxwAXKb2GVZ1AAxD/QsIAE9rYXZVf4DrTl7+kb8QAKBay8//28xcrX4HYF1LCwtAltZmVFMBMAztLTAA/Wtx
NjV3wOtOXnJLAIByll+0N/hXmtsBWNfywgPQttZnUNMBMAzt/wMA0J4eZk/zT2CdWwIATKmHwb/SzRNZJwQAGFNPg3+l+VsA5+nx
HwqAMnqdKV0+qXV2AwDYR6+Df6XrJ7dOCACwjd4H/0rEk1wnBAA4T8rgX4l6suuEAADDkDf4VyKf9DohAJApdfCvRD/5dUIAIEP6
4F+xCE8RAgB9MvjPshiXEAMAbTP0L2ZhtiAEANpi8G9mgXYkBgDqZOjvxmIdQAwAlGXo78/CjUgQAEzLwB+PhZyQIAA4jIE/HQs7
M1EAcD7Dfl4WuzICAeiVAV+X/wMZzgP3ZyE7mwAAAABJRU5ErkJggg==
"""


def app_icon():
    for n in ("apnescan.ico", "apnescan_icon.png"):
        p = resource_path(n)
        if os.path.exists(p):
            return QtGui.QIcon(p)
    # fallback: decode the embedded icon
    try:
        import base64 as _b64
        pm = QtGui.QPixmap()
        pm.loadFromData(_b64.b64decode(_EMBEDDED_ICON_B64))
        if not pm.isNull():
            return QtGui.QIcon(pm)
    except Exception:
        pass
    return QtGui.QIcon()


COLOUR_MODES = {"Colour (rang)": "color", "Grayscale": "gray", "Black & White": "bw"}
RESOLUTIONS = ["150", "200", "300", "600"]

# (id, label, default key) — all reassignable via Settings -> Keyboard Shortcuts
SHORTCUTS = [
    ("scan", "Scan", "Return"),
    ("import", "Import images / PDF", "Ctrl+O"),
    ("rename", "Rename page", "F2"),
    ("save_all", "Save all as PDF", "Ctrl+S"),
    ("save_sel", "Save selected as PDF", "Space"),
    ("save_pw", "Save PDF (password)", ""),
    ("save_img", "Save images", "Ctrl+Shift+S"),
    ("ocr_text", "Export OCR text", "Ctrl+Alt+O"),
    ("print", "Print", "Ctrl+P"),
    ("rotate_left", "Rotate left", "Ctrl+Left"),
    ("rotate_right", "Rotate right", "Ctrl+Right"),
    ("bright_up", "Brightness +", ""),
    ("bright_dn", "Brightness -", ""),
    ("contrast_up", "Contrast +", ""),
    ("contrast_dn", "Contrast -", ""),
    ("autocrop", "Auto-crop page", ""),
    ("autoname", "Auto-name pages (OCR)", "Ctrl+Alt+N"),
    ("undo", "Undo delete", "Ctrl+Z"),
    ("delete", "Delete page", "Delete"),
    ("clear", "Clear all", "Ctrl+Shift+Delete"),
    ("move_up", "Move page up", "Ctrl+Up"),
    ("move_down", "Move page down", "Ctrl+Down"),
    ("search", "Search past PDFs", "Ctrl+F"),
    ("merge", "Merge PDFs", ""),
    ("split", "Split into PDFs", ""),
    ("report", "Monthly report", ""),
    ("zoom_in", "Zoom in (thumbnails)", "Ctrl+="),
    ("zoom_out", "Zoom out (thumbnails)", "Ctrl+-"),
    ("options", "Options / Settings", "Ctrl+,"),
    ("profiles", "Profiles", "Ctrl+L"),
]

DEFAULT_OPTIONS = {
    "auto_save": False,
    "save_folder": os.path.join(os.path.expanduser("~"), "Documents", "NobleScans"),
    "filename_template": "{claim}_{date}_{seq}",
    "make_claim_folder": False,
    "year_month_folders": False,
    "remove_blank": False,
    "blank_sensitivity": "normal",   # kam / normal / zyada
    "auto_crop": False,
    "deskew": False,
    "quality_enhance": False,
    "compress": False,
    "jpeg_quality": 60,
    "watermark": False,
    "watermark_text": "Noble Care Hospital",
    "batch_mode": False,
    "validate_claim": False,
    "claim_pattern": r"^[A-Za-z0-9\-]{4,}$",
    "barcode_autofill": False,
    "duplicate_check": False,
    "excel_log": False,
    "activity_log": False,
    "backup": False,
    "backup_folder": os.path.join(os.path.expanduser("~"), "Documents", "NobleScans_Backup"),
    "scanner_method": "twain",   # "twain" ya "wia"
    "twain_file_xfer": False,    # experimental: continuous ADF feed (TWAIN file transfer)
    "auto_orient": False,        # ulta page OCR se seedha karo
    "auto_colour": False,        # rangeen page colour me, baki gray me (chhoti file)
    "custom_page_mm": 600,       # "custom" page size ki lambai (mm)
    "touch_mode": False,         # bade buttons/font (touch / buzurg mode)
    "sign_image": "",            # sign/stamp wali image ka path
    "tags": {},                  # pdf path -> [tags]
    "show_files_panel": True,    # daayan "Meri Files" panel
    "show_left_panel": True,     # baayan scan-settings panel
    "fav_folders": [],           # panel ke ⭐ favourite folders
    "sidebar_stats": [],         # khaali = default set dikhega
    "ui_dashboard": True,        # khaali screen par bade action-cards
    "ui_fab": False,             # floating gol Scan button
    "ui_header": True,           # toolbar ke neeche status-patti
    "ui_graph": True,            # sidebar me 7-din ka graph
    "ui_preview": False,         # daayan preview panel
    "ui_jobs": False,            # job-chips patti
    "ui_kiosk": False,           # kiosk overlay
    "jobs": [],                  # job-chips: [{name,icon,profile,folder,template}]
    "ui_ribbon": False,          # ribbon toolbar (classic toolbar ki jagah)
    "wia_device_id": None,
    "language": "hi",            # "hi" ya "en"
    "simple_mode": False,
    "feedback_email": "",
    "fast_mode": False,
    "after_save": "nothing",   # nothing / open / folder
    "naps2_path": "",
    "naps2_profile": "",
    "scanner_ip": "",
    "theme": "light",             # light / dark
    "show_page_numbers": True,
    "save_images_too": False,     # also save each page as an image
    "auto_name": False,           # OCR the document title as the page label
    "shortcuts": {},              # id -> custom key (overrides default)
}


def load_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Image helpers
# ---------------------------------------------------------------------------

def is_blank_page(img, dark_fraction_threshold=0.0008, dark_level=75):
    # Count only VERY DARK ink (< dark_level, default 75). Real printed text / stamps
    # are near-black; fold-line shadows and faint bleed-through from the front are
    # medium/light grey (>75) and are therefore IGNORED. So a blank back side (only
    # creases + bleed-through) reads ~0% and gets dropped, while a page with real
    # (even sparse) dark text is kept. No erosion (that was thinning real text).
    try:
        g = img.convert("L")
        ink = g.point(lambda p: 255 if p < dark_level else 0)
        hist = ink.histogram()
        total = (g.width * g.height) or 1
        return (hist[-1] / total) < dark_fraction_threshold
    except Exception:
        return False


def whiten_dark_background(img, bright_thresh=160):
    """Scanner ki DARK backing (kala / navy / GRAY — sab) jo page ke chaaron
    taraf dikhti hai use WHITE karo — pixel-level par.

    Har row/column me KINARE se shuru karke pehle BRIGHT (paper) pixel tak ke
    pixels hi white hote hain. Isliye:
      - tedha page ho to bhi uska kona/kinara kabhi nahi 'katta' (purana code
        poori patti white pot deta tha jisme page ka kinara bhi chala jata
        tha — 'side se kata hua' wala bug),
      - gray backing bhi pakdi jaati hai (purana threshold sirf <90 tha,
        gray 90-160 wali backing reh jaati thi),
      - page ke ANDAR ka content (text, X-ray, stamp) kabhi nahi chhedha
        jaata — run pehle bright pixel par ruk jaata hai."""
    try:
        import numpy as _np
        g = _np.asarray(img.convert("L"))
        # fast early-out: kinare pehle se bright hain to backing hai hi nahi
        if (g[0].mean() > 150 and g[-1].mean() > 150
                and g[:, 0].mean() > 150 and g[:, -1].mean() > 150):
            return img
        bright = g >= bright_thresh
        H, W = g.shape
        arr = _np.asarray(img.convert("RGB")).copy()
        cols = _np.arange(W)[None, :]
        rows = _np.arange(H)[:, None]
        # Left/Right: har row me pehla/aakhri bright pixel — uske bahar sab white.
        any_row = bright.any(axis=1)
        first = _np.where(any_row, bright.argmax(axis=1), W).astype(_np.int64)[:, None]
        last = _np.where(any_row, W - 1 - bright[:, ::-1].argmax(axis=1), -1).astype(_np.int64)[:, None]
        mask = (cols < first) | (cols > last)
        # Poori-dark rows (koi bright pixel nahi) sirf tab white hon jab wo
        # upar/neeche ke KINARE se lagatar judi hon (backing band). Beech ki
        # dark rows (jaise X-ray film) kabhi nahi.
        no_bright_row = ~any_row
        top_run = _np.cumprod(no_bright_row).astype(bool)
        bot_run = _np.cumprod(no_bright_row[::-1]).astype(bool)[::-1]
        interior_dark_rows = no_bright_row & ~top_run & ~bot_run
        mask[interior_dark_rows, :] = False
        # Top/Bottom: har column me bhi wahi — kinare se pehle bright pixel tak.
        any_col = bright.any(axis=0)
        firstc = _np.where(any_col, bright.argmax(axis=0), H).astype(_np.int64)[None, :]
        lastc = _np.where(any_col, H - 1 - bright[::-1, :].argmax(axis=0), -1).astype(_np.int64)[None, :]
        cmask = (rows < firstc) | (rows > lastc)
        no_bright_col = ~any_col
        left_run = _np.cumprod(no_bright_col).astype(bool)
        right_run = _np.cumprod(no_bright_col[::-1]).astype(bool)[::-1]
        interior_dark_cols = no_bright_col & ~left_run & ~right_run
        cmask[:, interior_dark_cols] = False
        mask |= cmask
        arr[mask] = 255
        return Image.fromarray(arr)
    except Exception:
        return img


def trim_dark_borders(img, bright_thresh=120, min_paper_frac=0.04, pad=6):
    """Remove wide DARK margins (the scanner's backing — black OR dark-blue —
    showing around a narrow sheet like an ECG strip) so the page comes out to its
    real paper. A row/column is "paper" if enough of its pixels are BRIGHT
    (paper background), so dark backing is trimmed. Safe: only crops when a real
    dark border exists; leaves normal white pages and fully-dark pages untouched."""
    try:
        import numpy as _np
        g = img.convert("L")
        a = _np.asarray(g)
        bright = a > bright_thresh                       # paper background is bright
        row_has = bright.mean(axis=1) > min_paper_frac
        col_has = bright.mean(axis=0) > min_paper_frac
        rows = _np.where(row_has)[0]
        cols = _np.where(col_has)[0]
        if len(rows) == 0 or len(cols) == 0:
            return img                                   # no bright paper found
        t, b = int(rows[0]), int(rows[-1])
        l, r = int(cols[0]), int(cols[-1])
        t = max(0, t - pad); l = max(0, l - pad)
        b = min(a.shape[0] - 1, b + pad); r = min(a.shape[1] - 1, r + pad)
        # only crop if it removes a meaningful border (>1.5% on some side)
        if (r - l) < img.width * 0.985 or (b - t) < img.height * 0.985:
            return img.crop((l, t, r + 1, b + 1))
        return img
    except Exception:
        return img


def autocrop(img, border=20):
    try:
        rgb = img.convert("RGB")
        bg = Image.new("RGB", rgb.size, (255, 255, 255))
        bbox = ImageChops.difference(rgb, bg).getbbox()
        if not bbox:
            return img
        l, t, r, b = bbox
        l = max(0, l - border); t = max(0, t - border)
        r = min(img.width, r + border); b = min(img.height, b + border)
        return img.crop((l, t, r, b))
    except Exception:
        return img


def deskew(img):
    if not HAS_NUMPY:
        return img
    try:
        g = img.convert("L")
        small = g.resize((max(1, g.width // 4), max(1, g.height // 4)))
        best_angle, best_score = 0.0, -1.0
        for i in range(-10, 11):
            angle = i * 0.5
            rot = small.rotate(angle, resample=Image.BILINEAR, fillcolor=255)
            arr = np.asarray(rot, dtype=np.float32)
            score = float(np.var(np.diff(arr.sum(axis=1))))
            if score > best_score:
                best_score, best_angle = score, angle
        if abs(best_angle) < 0.25:
            return img
        return img.rotate(best_angle, resample=Image.BICUBIC, fillcolor=(255, 255, 255), expand=True)
    except Exception:
        return img


def auto_enhance(img):
    """Clean up faded / dull documents automatically."""
    try:
        rgb = img.convert("RGB")
        rgb = ImageOps.autocontrast(rgb, cutoff=1)
        rgb = ImageEnhance.Sharpness(rgb).enhance(1.4)
        rgb = ImageEnhance.Contrast(rgb).enhance(1.06)
        return rgb
    except Exception:
        return img


def flatten_photo_shadows(img):
    """Phone-photo ke shadow / roshni ke dabbe hatao: background ko blur karke
    usse divide karo — paper flat white ho jata hai, text/stamp waise hi rehte
    hain. numpy na ho to sirf autocontrast."""
    if not HAS_NUMPY:
        return ImageOps.autocontrast(img.convert("RGB"), cutoff=1)
    try:
        import numpy as _np
        rgb = img.convert("RGB")
        bg = rgb.convert("L").filter(ImageFilter.GaussianBlur(radius=40))
        b = _np.asarray(bg, dtype=_np.float32)
        b[b < 1] = 1
        a = _np.asarray(rgb, dtype=_np.float32)
        flat = _np.clip(a / b[..., None] * 230.0, 0, 255).astype("uint8")
        out = Image.fromarray(flat, "RGB")
        return ImageOps.autocontrast(out, cutoff=1)
    except Exception:
        return img


def clean_photo(img):
    """Phone se kheenchi photo ko scan-jaisa banao: EXIF ke hisaab se seedha,
    bahut badi ho to chhota, shadow hatao."""
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    img = img.convert("RGB")
    m = max(img.width, img.height)
    if m > 2600:
        s = 2600.0 / m
        img = img.resize((max(1, int(img.width * s)), max(1, int(img.height * s))),
                         Image.LANCZOS)
    return flatten_photo_shadows(img)


def _largest_gap(flags, min_gap):
    """Bool array me sabse badi False-run (start,end) jo kinaron ko na chhue."""
    best = None
    run = 0
    for i, v in enumerate(flags):
        if not v:
            run += 1
        else:
            if run >= min_gap and i - run > 0:
                if best is None or run > (best[1] - best[0]):
                    best = (i - run, i)
            run = 0
    return best


def detect_content_boxes(img, bg_thresh=235):
    """White background par rakhe alag-alag cards/documents ke boxes dhoondo
    (2-3 ID ek saath scan karne par). Bade khaali gap par recursively todta hai."""
    if not HAS_NUMPY:
        return [(0, 0, img.width, img.height)]
    import numpy as _np
    g = _np.asarray(img.convert("L"))
    mask = g < bg_thresh
    boxes = []

    def _split(t, b, l, r, depth):
        sub = mask[t:b, l:r]
        if sub.size == 0 or not sub.any():
            return
        rows = sub.any(axis=1)
        cols = sub.any(axis=0)
        t2 = t + int(_np.argmax(rows))
        b2 = b - int(_np.argmax(rows[::-1]))
        l2 = l + int(_np.argmax(cols))
        r2 = r - int(_np.argmax(cols[::-1]))
        if depth < 6:
            rows2 = mask[t2:b2, l2:r2].any(axis=1)
            gap = _largest_gap(rows2, max(14, (b2 - t2) // 20))
            if gap:
                _split(t2, t2 + gap[0], l2, r2, depth + 1)
                _split(t2 + gap[1], b2, l2, r2, depth + 1)
                return
            cols2 = mask[t2:b2, l2:r2].any(axis=0)
            gap = _largest_gap(cols2, max(14, (r2 - l2) // 20))
            if gap:
                _split(t2, b2, l2, l2 + gap[0], depth + 1)
                _split(t2, b2, l2 + gap[1], r2, depth + 1)
                return
        boxes.append((l2, t2, r2, b2))

    _split(0, 0 + mask.shape[0], 0, mask.shape[1], 0)
    return boxes


def auto_orient(img):
    """Ulta (90/180/270 ghuma hua) page OCR (Tesseract OSD) se pehchan kar
    seedha karo. Tesseract na ho ya samajh na aaye to page waise hi rehta hai."""
    try:
        osd = pytesseract.image_to_osd(img.convert("RGB"))
        m = re.search(r"Rotate:\s*(\d+)", osd or "")
        rot = int(m.group(1)) if m else 0
        if rot in (90, 180, 270):
            return img.rotate(-rot, expand=True)
    except Exception:
        pass
    return img


def colorfulness(img):
    """Page kitna rangeen hai (0 = pura B&W jaisa). Auto colour-detect ke liye."""
    try:
        import numpy as _np
        small = img.convert("RGB").resize(
            (max(1, img.width // 8), max(1, img.height // 8)))
        a = _np.asarray(small, dtype=_np.int16)
        rg = _np.abs(a[..., 0] - a[..., 1])
        yb = _np.abs((a[..., 0] + a[..., 1]) // 2 - a[..., 2])
        return float(rg.mean() + yb.mean())
    except Exception:
        return 999.0


def restore_photo(img):
    """Purani dhundhli/feeki photo ko sudharo: contrast, rang, sharpness."""
    try:
        rgb = img.convert("RGB")
        rgb = rgb.filter(ImageFilter.MedianFilter(3))          # halka noise saaf
        rgb = ImageOps.autocontrast(rgb, cutoff=2)
        rgb = ImageEnhance.Color(rgb).enhance(1.25)
        rgb = ImageEnhance.Sharpness(rgb).enhance(1.3)
        return rgb
    except Exception:
        return img


def save_image_keep_ext(img, path, quality=92):
    """Image ko uske extension ke hisaab se sahi format me save karo."""
    if path.lower().endswith((".jpg", ".jpeg")):
        img.convert("RGB").save(path, "JPEG", quality=quality)
    else:
        img.save(path, "PNG")


def apply_watermark(img, text):
    try:
        base = img.convert("RGBA")
        overlay = Image.new("RGBA", base.size, (255, 255, 255, 0))
        draw = ImageDraw.Draw(overlay)
        size = max(14, base.width // 45)
        try:
            font = ImageFont.truetype("arial.ttf", size)
        except Exception:
            font = ImageFont.load_default()
        stamp = "%s  |  %s" % (text, datetime.datetime.now().strftime("%d-%m-%Y"))
        try:
            bbox = draw.textbbox((0, 0), stamp, font=font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        except Exception:
            tw, th = draw.textsize(stamp, font=font)
        x = (base.width - tw) // 2
        y = base.height - th - max(10, base.height // 60)
        draw.rectangle([x - 8, y - 6, x + tw + 8, y + th + 6], fill=(255, 255, 255, 160))
        draw.text((x, y), stamp, fill=(120, 120, 120, 200), font=font)
        return Image.alpha_composite(base, overlay).convert("RGB")
    except Exception:
        return img


def enhance_image(path, brightness=1.0, contrast=1.0):
    try:
        with Image.open(path) as im:
            im = im.convert("RGB")
            if brightness != 1.0:
                im = ImageEnhance.Brightness(im).enhance(brightness)
            if contrast != 1.0:
                im = ImageEnhance.Contrast(im).enhance(contrast)
            im.save(path, "PNG")
        return True
    except Exception:
        return False


def read_barcode(path):
    """Return first barcode/QR text found in the image, else None."""
    if not HAS_ZBAR:
        return None
    try:
        with Image.open(path) as im:
            results = _zbar_decode(im.convert("RGB"))
        for r in results:
            try:
                val = r.data.decode("utf-8").strip()
            except Exception:
                val = str(r.data).strip()
            if val:
                return val
        return None
    except Exception:
        return None


def sanitize(text, fallback="scan"):
    safe = "".join(c for c in (text or "") if c.isalnum() or c in "-_")
    return safe or fallback


# ---------------------------------------------------------------------------
# Network + TWAIN backend
# ---------------------------------------------------------------------------

def tcp_reachable(ip, ports=SCANNER_PORTS, timeout=1.2):
    for p in ports:
        try:
            with socket.create_connection((ip, p), timeout=timeout):
                return True
        except Exception:
            continue
    return False


class ConnectionChecker(QtCore.QThread):
    result = QtCore.pyqtSignal(bool, str)

    def __init__(self, ip):
        super().__init__()
        self.ip = (ip or "").strip()

    def run(self):
        if not self.ip:
            self.result.emit(False, "Scanner ka IP address set nahi hai")
            return
        if tcp_reachable(self.ip):
            self.result.emit(True, "Connected — scanner network par mil gaya (%s)" % self.ip)
        else:
            self.result.emit(False, "Not connected — %s reachable nahi. Scanner ON hai? "
                                    "Same WiFi/LAN par hai?" % self.ip)


class ScannerStateChecker(QtCore.QThread):
    """Ask an eSCL scanner whether it's Idle (free) or Processing (busy)."""
    state = QtCore.pyqtSignal(str)   # "free" / "busy" / "unknown"

    def __init__(self, ip):
        super().__init__()
        self.ip = (ip or "").strip()

    def run(self):
        if not self.ip:
            self.state.emit("unknown"); return
        try:
            import urllib.request as _u
            r = _u.urlopen("http://%s/eSCL/ScannerStatus" % self.ip, timeout=5)
            xml = r.read().decode("utf-8", "ignore")
            m = re.search(r"State>\s*([A-Za-z]+)", xml)
            st = (m.group(1).lower() if m else "")
            if st == "idle":
                self.state.emit("free")
            elif st:
                self.state.emit("busy")
            else:
                self.state.emit("unknown")
        except Exception:
            self.state.emit("unknown")


class EsclTestWorker(QtCore.QThread):
    """Step-by-step eSCL probe that reports the ACTUAL response at each stage,
    so the real cause of a scan failure (busy job / bad settings / offline) is
    visible instead of just a generic error."""
    done = QtCore.pyqtSignal(str)

    def __init__(self, ip):
        super().__init__()
        self.ip = (ip or "").strip()

    def run(self):
        import urllib.request as U
        import urllib.error as UE
        lines = []
        def add(t=""):
            lines.append(t)
        add("===== ApneScan eSCL Test =====")
        add("Time: %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        add("Scanner IP: %s" % (self.ip or "(not set)"))
        base = "http://%s/eSCL" % self.ip

        add("")
        add("[1] TCP connect (port 80):")
        try:
            add("    %s" % ("OK - reachable" if tcp_reachable(self.ip) else "FAIL - not reachable (scanner ON? same WiFi/LAN?)"))
        except Exception as e:
            add("    ERROR: %s" % e)

        add("")
        add("[2] GET /eSCL/ScannerCapabilities:")
        try:
            r = U.urlopen(base + "/ScannerCapabilities", timeout=8)
            body = r.read().decode("utf-8", "ignore")
            add("    HTTP %s, %d bytes" % (getattr(r, "status", 200), len(body)))
            mw = re.findall(r"MaxWidth>\s*(\d+)", body)
            mh = re.findall(r"MaxHeight>\s*(\d+)", body)
            if mw and mh:
                add("    Max scan area: W=%s H=%s (1/300 inch)" % (max(mw), max(mh)))
            if "Duplex" in body:
                add("    Duplex: supported (eSCL)")
        except UE.HTTPError as e:
            add("    HTTP ERROR %s (eSCL shayad support nahi)" % e.code)
        except Exception as e:
            add("    ERROR: %s" % e)

        add("")
        add("[3] GET /eSCL/ScannerStatus (free/busy + open jobs):")
        try:
            r = U.urlopen(base + "/ScannerStatus", timeout=8)
            body = r.read().decode("utf-8", "ignore")
            st = re.search(r"State>\s*([A-Za-z]+)", body)
            add("    HTTP %s, State=%s" % (getattr(r, "status", 200), st.group(1) if st else "?"))
            jobs = re.findall(r"JobState>\s*([A-Za-z]+)", body)
            if jobs:
                add("    Open/last jobs: %s" % ", ".join(jobs))
                add("    (Agar koi 'Processing'/'Pending' job hai to scanner busy rahega.)")
            else:
                add("    No open jobs listed.")
        except UE.HTTPError as e:
            add("    HTTP ERROR %s" % e.code)
        except Exception as e:
            add("    ERROR: %s" % e)

        add("")
        add("[4] POST /eSCL/ScanJobs  (THE REAL TEST - minimal A4 simplex gray):")
        settings = (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<scan:ScanSettings xmlns:scan="http://schemas.hp.com/imaging/escl/2011/05/03" '
            'xmlns:pwg="http://www.pwg.org/schemas/2010/12/sm">\n'
            '<pwg:Version>2.6</pwg:Version>\n'
            '<pwg:ScanRegions pwg:MustHonor="false"><pwg:ScanRegion>\n'
            '<pwg:Height>3550</pwg:Height><pwg:Width>2480</pwg:Width>\n'
            '<pwg:XOffset>0</pwg:XOffset><pwg:YOffset>0</pwg:YOffset>\n'
            '</pwg:ScanRegion></pwg:ScanRegions>\n'
            '<pwg:InputSource>Feeder</pwg:InputSource>\n'
            '<scan:ColorMode>Grayscale8</scan:ColorMode>\n'
            '<scan:XResolution>200</scan:XResolution>\n'
            '<scan:YResolution>200</scan:YResolution>\n'
            '<pwg:DocumentFormat>image/jpeg</pwg:DocumentFormat>\n'
            '</scan:ScanSettings>'
        )
        job = None
        try:
            req = U.Request(base + "/ScanJobs", data=settings.encode("utf-8"),
                            headers={"Content-Type": "text/xml"}, method="POST")
            r = U.urlopen(req, timeout=15)
            job = r.headers.get("Location")
            add("    HTTP %s  ->  JOB CREATED" % getattr(r, "status", 201))
            add("    Location: %s" % job)
        except UE.HTTPError as e:
            add("    HTTP ERROR %s" % e.code)
            try:
                body = e.read().decode("utf-8", "ignore")
                if body.strip():
                    add("    Response body: %s" % body.replace("\n", " ")[:300])
            except Exception:
                pass
            if e.code == 503:
                add("    => 503 = scanner BUSY / ek purana job khula hai.")
                add("       Fix: scanner off/on karein; koi aur scan app band karein.")
            elif e.code in (400, 415, 409):
                add("    => %s = scan settings/format ka problem (is scanner ke liye tweak chahiye)." % e.code)
        except Exception as e:
            add("    ERROR: %s" % e)

        if job:
            if job.startswith("/"):
                job = "http://%s%s" % (self.ip, job)
            add("")
            add("[5] GET job/NextDocument (1 page fetch):")
            try:
                r = U.urlopen(job + "/NextDocument", timeout=30)
                d = r.read()
                add("    HTTP %s, %d bytes  ->  PAGE RECEIVED (scan works!)" % (getattr(r, "status", 200), len(d)))
            except UE.HTTPError as e:
                add("    HTTP %s (koi page nahi / done)" % e.code)
            except Exception as e:
                add("    ERROR: %s" % e)
            add("")
            add("[6] DELETE job (release):")
            try:
                U.urlopen(U.Request(job, method="DELETE"), timeout=8)
                add("    OK - job released (scanner free).")
            except Exception as e:
                add("    ERROR: %s" % e)

        add("")
        add("===== RESULT =====")
        add("Agar [4] me 'JOB CREATED' aaya -> scan theek hona chahiye.")
        add("Agar [4] me 503 aaya -> scanner busy: off/on karein + doosri scan app band karein.")
        add("Agar [1] FAIL -> scanner network par nahi mila (IP/WiFi check).")
        add("==================")
        self.done.emit("\n".join(lines))


class StatsWorker(QtCore.QThread):
    """Talk to the ApneScan stats server (Google Apps Script). action:
    'ping' (mark online + fetch), 'scan' (add scans + fetch), 'stats' (fetch).
    App version + country (sirf 2-akshar ka code) bhi jaata hai — kabhi koi
    document/naam nahi."""
    got = QtCore.pyqtSignal(int, int, int)   # total, today, online
    got_full = QtCore.pyqtSignal(dict)       # poora server payload
    failed = QtCore.pyqtSignal()

    def __init__(self, url, client, action="ping", n=0):
        super().__init__()
        self.url = url; self.client = client; self.action = action; self.n = n
        try:
            self.country = (QtCore.QLocale().name().split("_") + [""])[1][:2]
        except Exception:
            self.country = ""

    def run(self):
        try:
            import urllib.request as U
            import urllib.parse as P
            import json as J
            q = {"action": self.action, "client": self.client,
                 "v": VERSION, "c": self.country}
            if self.action == "scan":
                q["n"] = str(self.n)
            full = self.url + ("&" if "?" in self.url else "?") + P.urlencode(q)
            r = U.urlopen(full, timeout=12)
            data = J.loads(r.read().decode("utf-8", "ignore"))
            if data.get("ok"):
                self.got.emit(int(data.get("total", 0)),
                              int(data.get("today", 0)),
                              int(data.get("online", 0)))
                try:
                    self.got_full.emit(dict(data))
                except Exception:
                    pass
            else:
                self.failed.emit()
        except Exception:
            self.failed.emit()


class UpdateChecker(QtCore.QThread):
    """Latest version ka pata karo — pehle WEBSITE se (version.txt, jo har
    deploy par apne aap banta hai), na mile to GitHub Releases se."""
    result = QtCore.pyqtSignal(str, str)   # latest_tag ("v19"), download_url

    def run(self):
        import urllib.request as U
        try:
            req = U.Request("https://apnescan.apnesoft.com/version.txt",
                            headers={"User-Agent": "ApneScan"})
            tag = U.urlopen(req, timeout=10).read().decode("utf-8", "ignore").strip()[:20]
            if re.search(r"\d", tag or ""):
                self.result.emit(tag, DOWNLOAD_PAGE)
                return
        except Exception:
            pass
        try:
            import json as J
            req = U.Request(UPDATE_API, headers={"User-Agent": "ApneScan"})
            data = J.loads(U.urlopen(req, timeout=10).read().decode("utf-8", "ignore"))
            self.result.emit(str(data.get("tag_name") or ""),
                             str(data.get("html_url") or DOWNLOAD_PAGE))
        except Exception:
            self.result.emit("", "")


class ScannerFinder(QtCore.QThread):
    """Network par eSCL scanner khud dhoondo — poore /24 subnet ko jaanchta hai
    (IP hath se daalne ki zaroorat nahi)."""
    found = QtCore.pyqtSignal(list)   # [(ip, model), ...]

    def run(self):
        import socket as _s
        import concurrent.futures as _cf
        import urllib.request as _u
        results = []
        try:
            s = _s.socket(_s.AF_INET, _s.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            my_ip = s.getsockname()[0]
            s.close()
        except Exception:
            self.found.emit([])
            return
        base = my_ip.rsplit(".", 1)[0]

        def probe(i):
            ip = "%s.%d" % (base, i)
            for port in (80, 8080):
                try:
                    c = _s.create_connection((ip, port), timeout=0.25)
                    c.close()
                except Exception:
                    continue
                try:
                    host = ip if port == 80 else "%s:%d" % (ip, port)
                    r = _u.urlopen("http://%s/eSCL/ScannerCapabilities" % host, timeout=1.5)
                    xml = r.read(4000).decode("utf-8", "ignore")
                    m = re.search(r"MakeAndModel>\s*([^<]+)", xml)
                    return (ip, (m.group(1).strip() if m else "eSCL scanner"))
                except Exception:
                    pass
            return None

        try:
            with _cf.ThreadPoolExecutor(max_workers=64) as ex:
                for res in ex.map(probe, range(1, 255)):
                    if res:
                        results.append(res)
        except Exception:
            pass
        self.found.emit(results)


class ScannerError(Exception):
    pass


def list_sources(hwnd):
    if not HAS_TWAIN:
        raise ScannerError("TWAIN (pytwain) install nahi hai.")
    sm = twain.SourceManager(hwnd)
    try:
        try:
            return list(sm.source_list)
        except Exception:
            return list(sm.GetSourceList())
    finally:
        try:
            sm.close()
        except Exception:
            try:
                sm.destroy()
            except Exception:
                pass


class _FileXferUnsupported(Exception):
    pass


def _open_twain_source(sm, source_name):
    try:
        src = sm.open_source(source_name) if source_name else sm.open_source()
    except Exception:
        try:
            src = sm.OpenSource(source_name) if source_name else sm.OpenSource()
        except Exception:
            src = None
    if src is None:
        try:
            src = sm.open_source()
        except Exception:
            try:
                src = sm.OpenSource()
            except Exception:
                src = None
    return src


def _common_caps(src, dpi, pixel_type, duplex):
    def _try(fn):
        try:
            fn()
        except Exception:
            pass
    _try(lambda: src.set_capability(twain.ICAP_XRESOLUTION, twain.TWTY_FIX32, float(dpi)))
    _try(lambda: src.set_capability(twain.ICAP_YRESOLUTION, twain.TWTY_FIX32, float(dpi)))
    pt = {"color": twain.TWPT_RGB, "gray": twain.TWPT_GRAY, "bw": twain.TWPT_BW}[pixel_type]
    _try(lambda: src.set_capability(twain.ICAP_PIXELTYPE, twain.TWTY_UINT16, pt))
    _try(lambda: src.set_capability(twain.CAP_FEEDERENABLED, twain.TWTY_BOOL, True))
    _try(lambda: src.set_capability(twain.CAP_AUTOFEED, twain.TWTY_BOOL, True))
    _try(lambda: src.set_capability(twain.CAP_AUTOSCAN, twain.TWTY_BOOL, True))
    _try(lambda: src.set_capability(twain.CAP_XFERCOUNT, twain.TWTY_INT16, -1))
    _try(lambda: src.set_capability(twain.CAP_DUPLEXENABLED, twain.TWTY_BOOL, bool(duplex)))


def _scan_file(hwnd, source_name, dpi, pixel_type, duplex, on_page=None, should_stop=None):
    """File-transfer acquire: the driver writes each page to a file and (on most
    ADF scanners) keeps the feeder moving continuously. If the source doesn't
    support it, raises _FileXferUnsupported so the caller falls back to native.
    Only raises _FileXferUnsupported when ZERO pages were produced (so we never
    double-feed paper)."""
    count = 0
    MAX_PAGES = 1000   # hard guard: never loop forever
    sm = twain.SourceManager(hwnd)
    src = None
    try:
        src = _open_twain_source(sm, source_name)
        if src is None:
            raise _FileXferUnsupported()
        _common_caps(src, dpi, pixel_type, duplex)
        # Switch to FILE transfer mechanism. If unsupported -> fall back.
        try:
            src.set_capability(twain.ICAP_XFERMECH, twain.TWTY_UINT16, twain.TWSX_FILE)
        except Exception:
            raise _FileXferUnsupported()
        try:
            src.request_acquire(show_ui=False, modal_ui=False)
        except Exception:
            raise _FileXferUnsupported()

        while count < MAX_PAGES:
            if should_stop and should_stop():
                break
            fd, bmp = tempfile.mkstemp(suffix=".bmp")
            os.close(fd)
            try:
                os.remove(bmp)
            except Exception:
                pass
            # point the driver at our file
            try:
                src.file_xfer_params(bmp, twain.TWFF_BMP)
            except Exception:
                if count == 0:
                    raise _FileXferUnsupported()
                break
            # transfer one page to the file
            try:
                more = src.xfer_image_by_file()
            except Exception:
                # feeder empty / done (or unsupported on first try)
                if count == 0:
                    raise _FileXferUnsupported()
                break
            if not os.path.exists(bmp):
                if count == 0:
                    raise _FileXferUnsupported()
                break
            try:
                with Image.open(bmp) as im:
                    img = im.convert("RGB").copy()
            finally:
                try:
                    os.remove(bmp)
                except Exception:
                    pass
            count += 1
            if on_page:
                on_page(img)
            # 'more' may be a bool/int/pending-count depending on pytwain; stop when falsy
            if more in (None, 0, False):
                break
        return count
    finally:
        try:
            if src is not None:
                src.close()
        except Exception:
            pass
        try:
            sm.close()
        except Exception:
            try:
                sm.destroy()
            except Exception:
                pass


def scan_pages(hwnd, source_name, dpi, pixel_type, duplex, on_page=None, should_stop=None,
               file_xfer=False):
    """TWAIN scan. When file_xfer=True (Settings -> experimental continuous ADF
    feed) try the file-transfer path first and fall back to the proven native
    path only when zero pages were produced."""
    if file_xfer:
        try:
            n = _scan_file(hwnd, source_name, dpi, pixel_type, duplex, on_page, should_stop)
            if n > 0:
                return n
            raise _FileXferUnsupported()
        except _FileXferUnsupported:
            pass
        except ScannerError:
            raise
        except Exception:
            pass
    return _scan_native(hwnd, source_name, dpi, pixel_type, duplex, on_page, should_stop)


def _scan_native(hwnd, source_name, dpi, pixel_type, duplex, on_page=None, should_stop=None):
    if not HAS_TWAIN:
        raise ScannerError("TWAIN (pytwain) install nahi hai.")
    count = 0
    sm = twain.SourceManager(hwnd)
    src = None
    try:
        try:
            src = sm.open_source(source_name) if source_name else sm.open_source()
        except Exception:
            try:
                src = sm.OpenSource(source_name) if source_name else sm.OpenSource()
            except Exception:
                src = None
        if src is None:
            # Stored name might be a WIA name (when switching method) -> open default TWAIN source.
            try:
                src = sm.open_source()
            except Exception:
                try:
                    src = sm.OpenSource()
                except Exception:
                    src = None
        if src is None:
            raise ScannerError("Scanner open nahi hua. Profile me device chuno.")

        def _try(fn):
            try:
                fn()
            except Exception:
                pass

        _try(lambda: src.set_capability(twain.ICAP_XRESOLUTION, twain.TWTY_FIX32, float(dpi)))
        _try(lambda: src.set_capability(twain.ICAP_YRESOLUTION, twain.TWTY_FIX32, float(dpi)))
        pt = {"color": twain.TWPT_RGB, "gray": twain.TWPT_GRAY, "bw": twain.TWPT_BW}[pixel_type]
        _try(lambda: src.set_capability(twain.ICAP_PIXELTYPE, twain.TWTY_UINT16, pt))
        # Continuous ADF feed: pull EVERY page in a single acquire (no per-page gap).
        _try(lambda: src.set_capability(twain.CAP_FEEDERENABLED, twain.TWTY_BOOL, True))
        _try(lambda: src.set_capability(twain.CAP_AUTOFEED, twain.TWTY_BOOL, True))
        _try(lambda: src.set_capability(twain.CAP_AUTOSCAN, twain.TWTY_BOOL, True))
        _try(lambda: src.set_capability(twain.CAP_XFERCOUNT, twain.TWTY_INT16, -1))
        _try(lambda: src.set_capability(twain.CAP_DUPLEXENABLED, twain.TWTY_BOOL, bool(duplex)))
        _try(lambda: src.request_acquire(show_ui=False, modal_ui=False))

        while True:
            if should_stop and should_stop():
                break
            try:
                rv = src.xfer_image_natively()
            except Exception as exc:
                if count:
                    break
                raise ScannerError("Scan transfer fail: %s" % exc)
            if rv is None:
                break
            handle, more = rv
            fd, bmp = tempfile.mkstemp(suffix=".bmp")
            os.close(fd)
            try:
                twain.dib_to_bm_file(handle, bmp)
            finally:
                try:
                    twain.global_handle_free(handle)
                except Exception:
                    pass
            with Image.open(bmp) as im:
                img = im.convert("RGB").copy()
            try:
                os.remove(bmp)
            except Exception:
                pass
            count += 1
            if on_page:
                on_page(img)
            if not more:
                break
    finally:
        try:
            if src is not None:
                src.close()
        except Exception:
            pass
        try:
            sm.close()
        except Exception:
            try:
                sm.destroy()
            except Exception:
                pass

    if count == 0:
        raise ScannerError("Koi page scan nahi hua. Feeder me document rakha hai? Scanner ON hai?")
    return count


# ---------------------------------------------------------------------------
# WIA backend (alternative to TWAIN, for other users' scanners)
# ---------------------------------------------------------------------------

WIA_ERROR_PAPER_EMPTY = -2145320957   # 0x80210003


def list_wia_sources():
    if not HAS_W32:
        raise ScannerError("WIA (pywin32) install nahi hai.")
    dm = _w32.Dispatch("WIA.DeviceManager")
    out = []
    for i in range(1, dm.DeviceInfos.Count + 1):
        info = dm.DeviceInfos.Item(i)
        name = None
        try:
            for p in info.Properties:
                if p.Name == "Name":
                    name = p.Value
                    break
        except Exception:
            pass
        out.append((info.DeviceID, name or ("Scanner %d" % i)))
    return out


def wia_scan_pages(device_id, dpi, pixel_type, duplex, on_page=None, should_stop=None):
    wia_scan_pages.last_duplex_value = getattr(wia_scan_pages, "last_duplex_value", None)
    """Scan via Windows WIA. Kept simple: no forced properties (some drivers
    throw otherwise). Runs inside a thread that has called CoInitialize."""
    if not HAS_W32:
        raise ScannerError("WIA (pywin32) install nahi hai.")
    dm = _w32.Dispatch("WIA.DeviceManager")
    device = None
    for i in range(1, dm.DeviceInfos.Count + 1):
        info = dm.DeviceInfos.Item(i)
        if device_id is None or info.DeviceID == device_id:
            device = info.Connect()
            break
    if device is None:
        raise ScannerError("WIA scanner nahi mila. Settings me scanner chuno.")

    def _wia_set(props, pid, value):
        # Set one WIA property; True if it took, False if the driver rejected it.
        try:
            for p in props:
                if int(p.PropertyID) == pid:
                    p.Value = value
                    return True
        except Exception:
            return False
        return False

    def _wia_valid_dpi(props, want):
        # Pick a resolution the device actually supports (avoids "parameter incorrect").
        try:
            for p in props:
                if int(p.PropertyID) == 6147:
                    vals = None
                    try:
                        vals = [int(x) for x in p.SubType.Values]   # WIA_PROP_LIST
                    except Exception:
                        vals = None
                    if vals:
                        return min(vals, key=lambda v: abs(v - want))
        except Exception:
            pass
        return want

    def _apply_feeder(dev, want_duplex=None):
        # Feeder + (optional) DUPLEX + all pages. HP network drivers use different
        # duplex "select" values, so try a few and keep the first the driver accepts.
        wd = duplex if want_duplex is None else want_duplex
        try:
            dprops = dev.Properties
            if wd:
                for val in (5, 7, 0x8005, 4, 3):   # FEEDER|DUPLEX variants (0x8005=32773 HP)
                    if _wia_set(dprops, 3088, val):
                        wia_scan_pages.last_duplex_value = val
                        break
                else:
                    wia_scan_pages.last_duplex_value = None
            else:
                _wia_set(dprops, 3088, 1)          # feeder only (single side)
            _wia_set(dprops, 3096, 0)              # all pages
        except Exception:
            pass

    def _apply_quality(it):
        try:
            props = it.Properties
            _wia_set(props, 4104, {"bw": 0, "gray": 2, "color": 3}.get(pixel_type, 3))
            dv = _wia_valid_dpi(props, int(dpi))
            _wia_set(props, 6147, dv)
            _wia_set(props, 6148, dv)
        except Exception:
            pass

    _apply_feeder(device)          # set FEEDER|DUPLEX on the DEVICE first
    item = device.Items[1]         # enumerate the page item AFTER duplex is active
    _apply_quality(item)
    _props_ok = True
    # BMP format id (widely accepted by HP WIA drivers).
    WIA_FMT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"

    def _transfer_one(it):
        # Try Transfer(format) first (robust on HP), then plain Transfer().
        try:
            return it.Transfer(WIA_FMT_BMP)
        except Exception:
            return it.Transfer()

    count = 0
    retry_stage = 0
    while True:
        if should_stop and should_stop():
            break
        try:
            image = _transfer_one(item)
        except Exception as exc:
            hres = exc.args[0] if getattr(exc, "args", None) else None
            # End-of-feeder conditions -> stop cleanly.
            if hres in (WIA_ERROR_PAPER_EMPTY, -2145320957) or count > 0:
                break
            # First page failed: retry once with a fresh item and NO custom props.
            if retry_stage < 2:
                stage = retry_stage
                retry_stage += 1
                try:
                    device2 = None
                    dm2 = _w32.Dispatch("WIA.DeviceManager")
                    for i in range(1, dm2.DeviceInfos.Count + 1):
                        info = dm2.DeviceInfos.Item(i)
                        if device_id is None or info.DeviceID == device_id:
                            device2 = info.Connect(); break
                    if device2 is not None:
                        if stage == 0:
                            _apply_feeder(device2)          # retry: duplex, no quality props
                        # stage 1: apply nothing -> guaranteed to scan (single side)
                        item = device2.Items[1]
                except Exception:
                    pass
                continue
            raise ScannerError("WIA scan fail: %s" % exc)
        fd, raw = tempfile.mkstemp(suffix=".bmp")
        os.close(fd)
        try:
            os.remove(raw)
        except Exception:
            pass
        try:
            image.SaveFile(raw)
        except Exception as exc:
            raise ScannerError("WIA save fail: %s" % exc)
        with Image.open(raw) as im:
            img = im.convert("RGB").copy()
        try:
            os.remove(raw)
        except Exception:
            pass
        count += 1
        if on_page:
            on_page(img)
    if count == 0:
        raise ScannerError("Koi page scan nahi hua (WIA).")
    return count


# ---------------------------------------------------------------------------
# Friendly error messages (technical -> simple)
# ---------------------------------------------------------------------------

def friendly_error(text, lang="hi"):
    t = str(text)
    low = t.lower()
    hi = {
        "not_found": "Scanner nahi mila. Scanner ON karein aur PC se juda hai ya nahi check karein.",
        "busy": "Scanner abhi busy hai. Thodi der baad dobara try karein.",
        "empty": "Feeder me koi document nahi hai. Page rakh kar dobara scan karein.",
        "jam": "Paper jam hua lagta hai. Feeder check karein.",
        "driver": "Scanner driver ki dikkat. Doosri scan-method (Settings me TWAIN/WIA) try karein.",
        "twain": "TWAIN scanner nahi khula. Settings me WIA method try karein.",
        "generic": "Scan nahi ho paya. Scanner ON hai aur juda hai? Dobara try karein.",
    }
    en = {
        "not_found": "Scanner not found. Turn it on and check it's connected.",
        "busy": "Scanner is busy. Please try again in a moment.",
        "empty": "No document in the feeder. Add a page and scan again.",
        "jam": "Looks like a paper jam. Check the feeder.",
        "driver": "Scanner driver issue. Try the other method (TWAIN/WIA) in Settings.",
        "twain": "TWAIN scanner didn't open. Try WIA method in Settings.",
        "generic": "Scanning failed. Is the scanner on and connected? Try again.",
    }
    m = hi if lang == "hi" else en
    if "paper_empty" in low or "0x80210003" in low or "-2145320957" in low or "feeder me" in low:
        return m["empty"]
    if "busy" in low or "0x80210006" in low:
        return m["busy"]
    if "jam" in low:
        return m["jam"]
    if "device driver threw" in low or "exception_in_driver" in low or "-2145320946" in low:
        return m["driver"]
    if "nahi mila" in low or "not found" in low or "no such" in low or "device" in low and "nahi" in low:
        return m["not_found"]
    if "twain" in low and ("open" in low or "khula" in low or "install" in low):
        return m["twain"]
    return m["generic"] + "\n\n(" + t[:200] + ")"


# ---------------------------------------------------------------------------
# Small translation layer (English / Hindi) for the main interface
# ---------------------------------------------------------------------------

T = {
    "scan": {"hi": "Scan", "en": "Scan"},
    "profile": {"hi": "Profile:", "en": "Profile:"},
    "profiles": {"hi": "Profiles…", "en": "Profiles…"},
    "claim_no": {"hi": "Claim No.:", "en": "Claim No.:"},
    "import": {"hi": "Import", "en": "Import"},
    "save_pdf": {"hi": "Save PDF", "en": "Save PDF"},
    "check": {"hi": "Check", "en": "Check"},
    "up": {"hi": "Upar", "en": "Up"},
    "down": {"hi": "Neeche", "en": "Down"},
    "delete": {"hi": "Delete", "en": "Delete"},
    "clear": {"hi": "Clear", "en": "Clear"},
    "rotate_l": {"hi": "Rotate ⟲", "en": "Rotate ⟲"},
    "rotate_r": {"hi": "Rotate ⟳", "en": "Rotate ⟳"},
    "menu_file": {"hi": "File", "en": "File"},
    "menu_edit": {"hi": "Edit", "en": "Edit"},
    "menu_tools": {"hi": "Tools", "en": "Tools"},
    "menu_settings": {"hi": "Settings", "en": "Settings"},
    "menu_help": {"hi": "Help", "en": "Help"},
    "help_guide": {"hi": "Guide / Madad", "en": "Guide / Help"},
    "about": {"hi": "About", "en": "About"},
    "whatsnew": {"hi": "Naya kya hai", "en": "What's new"},
    "feedback": {"hi": "Feedback / Sujhav", "en": "Feedback"},
    "language": {"hi": "Language / Bhasha", "en": "Language"},
    "scan_method": {"hi": "Scan method (TWAIN/WIA)", "en": "Scan method (TWAIN/WIA)"},
    "options": {"hi": "Options…", "en": "Options…"},
    "pages_hint": {"hi": "Pages (drag se order badlein)", "en": "Pages (drag to reorder)"},
    "simple_on": {"hi": "Simple mode (aasan)", "en": "Simple mode"},
    "scanner_ip": {"hi": "Scanner IP:", "en": "Scanner IP:"},
    "st_pages": {"hi": "Pages", "en": "Pages"},
    "st_profile": {"hi": "Profile", "en": "Profile"},
    "none_profile": {"hi": "koi profile nahi", "en": "no profile"},
    "checking": {"hi": "Connection check ho raha hai...", "en": "Checking connection..."},
    "select_first": {"hi": "Pehle koi page select karein.", "en": "Select a page first."},
    "scan_first": {"hi": "Pehle koi page scan/import karein.", "en": "Scan or import a page first."},
    "fast": {"hi": "Fast", "en": "Fast"},
    "save_all": {"hi": "Sab pages ki PDF", "en": "Save all pages (PDF)"},
    "save_sel": {"hi": "Selected pages ki PDF", "en": "Save selected pages (PDF)"},
    "no_selection": {"hi": "Koi page select nahi hai. Thumbnails me Ctrl/Shift se pages chuno.", "en": "No pages selected. Use Ctrl/Shift to select pages."},
}

# Aur bhashayein (adhuri — jo key translate nahi hui wo Hindi me dikhegi)
_EXTRA_LANGS = {
    "mr": {"up": "वर", "down": "खाली", "help_guide": "मदत / Guide", "whatsnew": "नवीन काय आहे",
           "none_profile": "प्रोफाइल नाही", "checking": "कनेक्शन तपासत आहे...",
           "select_first": "आधी एखादे पान निवडा.", "scan_first": "आधी पान scan/import करा.",
           "save_all": "सर्व पानांची PDF", "save_sel": "निवडलेल्या पानांची PDF",
           "pages_hint": "Pages (क्रम बदलण्यासाठी drag करा)",
           "no_selection": "कोणतेही पान निवडलेले नाही. Ctrl/Shift ने निवडा."},
    "gu": {"up": "ઉપર", "down": "નીચે", "help_guide": "મદદ / Guide", "whatsnew": "નવું શું છે",
           "none_profile": "કોઈ પ્રોફાઇલ નથી", "checking": "કનેક્શન તપાસી રહ્યાં છીએ...",
           "select_first": "પહેલા કોઈ પાનું પસંદ કરો.", "scan_first": "પહેલા પાનું scan/import કરો.",
           "save_all": "બધા પાનાની PDF", "save_sel": "પસંદ કરેલા પાનાની PDF",
           "pages_hint": "Pages (ક્રમ બદલવા drag કરો)",
           "no_selection": "કોઈ પાનું પસંદ નથી. Ctrl/Shift થી પસંદ કરો."},
    "pa": {"up": "ਉੱਪਰ", "down": "ਹੇਠਾਂ", "help_guide": "ਮਦਦ / Guide", "whatsnew": "ਨਵਾਂ ਕੀ ਹੈ",
           "none_profile": "ਕੋਈ ਪ੍ਰੋਫਾਈਲ ਨਹੀਂ", "checking": "ਕਨੈਕਸ਼ਨ ਚੈੱਕ ਹੋ ਰਿਹਾ ਹੈ...",
           "select_first": "ਪਹਿਲਾਂ ਕੋਈ ਪੰਨਾ ਚੁਣੋ।", "scan_first": "ਪਹਿਲਾਂ ਪੰਨਾ scan/import ਕਰੋ।",
           "save_all": "ਸਾਰੇ ਪੰਨਿਆਂ ਦੀ PDF", "save_sel": "ਚੁਣੇ ਪੰਨਿਆਂ ਦੀ PDF",
           "pages_hint": "Pages (ਕ੍ਰਮ ਬਦਲਣ ਲਈ drag ਕਰੋ)",
           "no_selection": "ਕੋਈ ਪੰਨਾ ਨਹੀਂ ਚੁਣਿਆ। Ctrl/Shift ਨਾਲ ਚੁਣੋ।"},
    "ta": {"up": "மேலே", "down": "கீழே", "help_guide": "உதவி / Guide", "whatsnew": "புதிதாக என்ன",
           "none_profile": "சுயவிவரம் இல்லை", "checking": "இணைப்பு சரிபார்க்கப்படுகிறது...",
           "select_first": "முதலில் ஒரு பக்கத்தைத் தேர்வு செய்யவும்.",
           "scan_first": "முதலில் பக்கத்தை scan/import செய்யவும்.",
           "save_all": "எல்லா பக்கங்களின் PDF", "save_sel": "தேர்ந்த பக்கங்களின் PDF",
           "pages_hint": "Pages (வரிசை மாற்ற drag)",
           "no_selection": "எந்தப் பக்கமும் தேர்வில்லை. Ctrl/Shift பயன்படுத்தவும்."},
}
for _lc, _tab in _EXTRA_LANGS.items():
    for _k, _v in _tab.items():
        T.setdefault(_k, {})[_lc] = _v


def tr(key, lang="hi"):
    e = T.get(key)
    if not e:
        return key
    return e.get(lang, e.get("hi", key))


HELP_TEXT_HI = """<h3>ApneScan — Madad</h3>
<b>Shuru kaise karein</b><br>
1. <b>Settings → Profiles…</b> me ek profile banayein: <i>Choose device</i> se apna
scanner chuno, DPI/Colour/Duplex set karo.<br>
2. Upar profile chuno aur <b>Scan</b> dabao (ya F5).<br>
3. Pages left me dikhenge — drag se order badlo, rotate/delete karo.<br>
4. <b>Save PDF</b> se PDF banao (OCR ka option milega).<br><br>
<b>Scanner nahi mil raha?</b><br>
Settings → <i>Scan method</i> me TWAIN aur WIA dono try karo. Scanner ka driver
install hona chahiye, aur woh ON + connected ho.<br><br>
<b>Fast scan ke liye:</b> Profile me DPI 200, Mode Black & White.<br><br>
<b>Auto-save, register, backup, watermark, split, merge, password PDF</b> —
sab <b>Settings → Options</b> aur <b>Tools</b> menu me hain.<br><br>
Keyboard: F5=Scan, Ctrl+S=Save PDF, Ctrl+P=Print, Ctrl+F=Search, Delete=Delete page.
"""

CHANGELOG_HTML = """<h3>ApneScan — Naya kya hai / What's new</h3>
<b>Version 1.0</b><br>
- TWAIN + WIA dono scanners support<br>
- Setup wizard (pehli baar aasan setup)<br>
- Profiles, auto-save PDF, auto file-naming, claim folders<br>
- OCR searchable PDF (Hindi+English), OCR text export<br>
- Blank-page removal, auto-crop, deskew, quality enhance<br>
- PDF compress, password PDF, split, merge, print<br>
- Excel register, activity log, duplicate check, daily backup<br>
- Barcode/QR se claim number, monthly report<br>
- Hindi / English interface, Simple mode, Feedback, in-app Help<br>
"""


HELP_TEXT_EN = """<h3>ApneScan — Help</h3>
<b>Getting started</b><br>
1. Go to <b>Settings → Profiles…</b> and create a profile: click <i>Choose device</i>
to pick your scanner, then set DPI/Colour/Duplex.<br>
2. Select the profile at the top and click <b>Scan</b> (or press F5).<br>
3. Pages appear on the left — drag to reorder, rotate/delete.<br>
4. Click <b>Save PDF</b> (you'll be asked about OCR).<br><br>
<b>Scanner not detected?</b><br>
In Settings → <i>Scan method</i>, try both TWAIN and WIA. The scanner's driver
must be installed, and the device must be on and connected.<br><br>
<b>For faster scans:</b> set DPI 200 and Black & White in the profile.<br><br>
<b>Auto-save, register, backup, watermark, split, merge, password PDF</b> are all
under <b>Settings → Options</b> and the <b>Tools</b> menu.<br><br>
Shortcuts: F5=Scan, Ctrl+S=Save PDF, Ctrl+P=Print, Ctrl+F=Search, Delete=Delete page.
"""



# ---------------------------------------------------------------------------
# NAPS2 engine backend (uses NAPS2's console to scan — reliable fast + duplex)
# ---------------------------------------------------------------------------

NAPS2_GUESSES = [
    r"C:\\Program Files\\NAPS2\\NAPS2.Console.exe",
    r"C:\\Program Files (x86)\\NAPS2\\NAPS2.Console.exe",
    r"C:\\Program Files\\NAPS2\\NAPS2.exe",
    r"C:\\Program Files (x86)\\NAPS2\\NAPS2.exe",
]


def find_naps2():
    for p in NAPS2_GUESSES:
        if os.path.exists(p):
            return p
    return ""


def scan_via_naps2(naps2_exe, profile, tmpdir, duplex, on_page=None, should_stop=None):
    import subprocess
    import glob as _glob
    import re as _re
    if not naps2_exe or not os.path.exists(naps2_exe):
        raise ScannerError("NAPS2 nahi mila. Settings me NAPS2 ka path set karo.")
    # clean old temp scans
    for f in _glob.glob(os.path.join(tmpdir, "naps2_*.jpg")):
        try:
            os.remove(f)
        except Exception:
            pass
    out_pattern = os.path.join(tmpdir, "naps2_$(n).jpg")
    cmd = [naps2_exe, "-o", out_pattern, "--split", "-f", "-v"]
    if profile:
        cmd += ["-p", profile]
    if duplex:
        cmd += ["--source", "duplex"]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=600,
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired:
        raise ScannerError("NAPS2 scan me bahut time lag gaya (timeout).")
    except Exception as exc:
        raise ScannerError("NAPS2 chalane me dikkat: %s" % exc)

    files = _glob.glob(os.path.join(tmpdir, "naps2_*.jpg"))

    def _num(p):
        m = _re.search(r"naps2_(\d+)\.jpg$", os.path.basename(p))
        return int(m.group(1)) if m else 0
    files.sort(key=_num)

    if not files:
        msg = (proc.stderr or proc.stdout or "").strip()
        raise ScannerError("NAPS2 se koi page nahi aaya.\nProfile naam sahi hai? Scanner ON hai?\n\n%s"
                           % msg[:300])
    count = 0
    for f in files:
        if should_stop and should_stop():
            break
        try:
            with Image.open(f) as im:
                img = im.convert("RGB").copy()
        except Exception:
            continue
        try:
            os.remove(f)
        except Exception:
            pass
        count += 1
        if on_page:
            on_page(img)
    if count == 0:
        raise ScannerError("NAPS2 se koi page nahi aaya.")
    return count


# ---------------------------------------------------------------------------
# eSCL (AirScan) network backend — pure Python, talks to the scanner's IP over
# HTTP. Supports DUPLEX over the network with NO other software needed.
# ---------------------------------------------------------------------------

def scan_via_escl(ip, dpi, color, duplex, on_page=None, should_stop=None, page_size="auto"):
    import urllib.request as _u
    import urllib.error as _ue
    import io as _io
    import time as _t

    if not ip:
        raise ScannerError("Scanner IP set nahi hai. Settings \u2192 Scanner IP me IP daalo (jaise 192.168.1.8).")
    ip = ip.strip()
    base = "http://%s/eSCL" % ip
    # NOTE: "BlackAndWhite1" + image/jpeg ek saath kaam NAHI karta (JPEG 1-bit
    # ho hi nahi sakta). HP N4000 aisi job me pages kheenchta rehta hai par data
    # kabhi taiyar nahi karta — page feeder me atak jata tha aur app ko 0 page
    # milte the (Fast/B&W mode ka jam wala bug). Isliye B&W bhi Grayscale8 me
    # scan hota hai; asli 1-bit conversion app me hota hai (ScanWorker save step).
    cmode = {"color": "RGB24", "gray": "Grayscale8", "bw": "Grayscale8"}.get(color, "RGB24")

    # eSCL region units are 1/300 inch. Fixed sizes request that exact sheet.
    # "auto" requests the FULL BED WIDTH (so sides are never cut) with a Legal-length
    # height, and MustHonor="false" so the ADF stops at each sheet's real trailing
    # edge (no forced long scan = no feeder jam). The dark backing that appears
    # around a narrow sheet is painted white by whiten_dark_background.
    ps = (page_size or "auto").lower()
    if ps.startswith("a4"):
        w, h = 2480, 3550
    elif ps.startswith("letter"):
        w, h = 2550, 3300
    elif ps.startswith("legal"):
        w, h = 2550, 4200
    elif ps.startswith("a5"):
        w, h = 1748, 2480
    elif ps.startswith("custom"):
        # "custom:800" => 800mm lambi parchi (receipt/bahi-khata). Width full bed.
        mm = 600
        _m = re.findall(r"(\d+)", ps)
        if _m:
            mm = max(100, min(3000, int(_m[0])))
        w, h = 2550, int(round(mm * 300 / 25.4))
    else:  # auto -> full width, Legal length, let ADF stop at paper end
        w, h = 2550, 4200
        try:
            _rr = _u.urlopen(base + "/ScannerCapabilities", timeout=8)
            _cap = _rr.read().decode("utf-8", "ignore")
            _mw = re.findall(r"MaxWidth>\s*(\d+)", _cap)
            if _mw:
                w = max(int(x) for x in _mw)   # full bed width (no side cut)
        except Exception:
            pass

    settings = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<scan:ScanSettings xmlns:scan="http://schemas.hp.com/imaging/escl/2011/05/03"'
        ' xmlns:pwg="http://www.pwg.org/schemas/2010/12/sm">\n'
        '<pwg:Version>2.6</pwg:Version>\n'
        '<scan:Intent>Document</scan:Intent>\n'
        '<pwg:ScanRegions pwg:MustHonor="false"><pwg:ScanRegion>\n'
        '<pwg:Height>%d</pwg:Height><pwg:Width>%d</pwg:Width>\n'
        '<pwg:XOffset>0</pwg:XOffset><pwg:YOffset>0</pwg:YOffset>\n'
        '</pwg:ScanRegion></pwg:ScanRegions>\n'
        '<pwg:InputSource>Feeder</pwg:InputSource>\n'
        '<scan:Duplex>%s</scan:Duplex>\n'
        '<scan:ColorMode>%s</scan:ColorMode>\n'
        '<scan:XResolution>%d</scan:XResolution>\n'
        '<scan:YResolution>%d</scan:YResolution>\n'
        '<pwg:DocumentFormat>image/jpeg</pwg:DocumentFormat>\n'
        '</scan:ScanSettings>'
    ) % (h, w, "true" if duplex else "false", cmode, dpi, dpi)

    # Before starting, clear any STUCK job (e.g. a "Processing" job left over from
    # an earlier interrupted scan). Such a job makes every new ScanJobs POST return
    # 503 even though ScannerStatus shows Idle. We read ScannerStatus, find job
    # URIs/UUIDs whose state is not finished, and DELETE them.
    def _clear_stuck_jobs():
        try:
            rr = _u.urlopen(base + "/ScannerStatus", timeout=8)
            xml = rr.read().decode("utf-8", "ignore")
        except Exception:
            return
        blocks = re.findall(r"<[a-zA-Z]*:?JobInfo>.*?</[a-zA-Z]*:?JobInfo>", xml, re.S)
        for blk in blocks:
            stm = re.search(r"JobState>\s*([A-Za-z]+)", blk)
            state = (stm.group(1).lower() if stm else "")
            if state in ("completed", "canceled", "aborted", ""):
                continue                          # already finished -> leave it
            uri = re.search(r"JobUri>\s*([^<\s]+)", blk)
            uuid = re.search(r"JobUuid>\s*([^<\s]+)", blk)
            target = None
            if uri:
                target = uri.group(1)
                if target.startswith("/"):
                    target = "http://%s%s" % (ip, target)
            elif uuid:
                target = "%s/ScanJobs/%s" % (base, uuid.group(1))
            if target:
                try:
                    _u.urlopen(_u.Request(target, method="DELETE"), timeout=8)
                    _t.sleep(0.5)
                except Exception:
                    pass

    _clear_stuck_jobs()

    # Create the scan job. Try immediately first (this worked before); only if the
    # scanner reports busy (503/409) do we wait briefly and retry a couple of times.
    def _post_job():
        req = _u.Request(base + "/ScanJobs", data=settings.encode("utf-8"),
                         headers={"Content-Type": "text/xml"}, method="POST")
        return _u.urlopen(req, timeout=45)

    resp = None
    last_code = None
    for attempt in range(10):                    # be patient: scanner is briefly busy at start
        if should_stop and should_stop():
            break
        try:
            resp = _post_job()
            break
        except _ue.HTTPError as e:
            last_code = e.code
            if e.code in (503, 409, 429):        # busy -> wait and keep trying
                _t.sleep(1.5)
                continue
            raise ScannerError("eSCL job error: HTTP %s. Scanner network-scan (eSCL) support karta hai?" % e.code)
        except Exception as e:
            last_code = str(e)
            _t.sleep(1.2)
            continue
    if resp is None:
        raise ScannerError(
            "Scanner kaafi der se busy bata raha hai (HTTP %s). Agar koi AUR scan app "
            "(NAPS2 / HP Scan / purani ApneScan window) khuli hai to use band karein, "
            "ya scanner ko ek baar off/on kar dein." % last_code)

    job = resp.headers.get("Location")
    if not job:
        raise ScannerError("eSCL job location nahi mila (scanner ne job start nahi kiya).")
    if job.startswith("/"):
        job = "http://%s%s" % (ip, job)

    count = 0
    empties = 0
    hiccups = 0
    try:
        while True:
            if should_stop and should_stop():
                break
            try:
                r = _u.urlopen(job + "/NextDocument", timeout=90)
            except _ue.HTTPError as e:
                if e.code in (404, 410):
                    break                       # no more pages -> done
                if e.code == 503:
                    _t.sleep(1); empties += 1
                    if empties > 30:
                        break
                    continue
                break
            except Exception:
                # network hiccup (timeout/reset) — turant give-up mat karo, warna
                # job DELETE ho jata hai aur beech ka page feeder me atak jata hai
                hiccups += 1
                if hiccups > 3:
                    break
                _t.sleep(1.0)
                continue
            data = r.read()
            if not data:
                break
            try:
                with Image.open(_io.BytesIO(data)) as im:
                    img = im.convert("RGB").copy()
            except Exception:
                break
            count += 1
            empties = 0
            hiccups = 0
            if on_page:
                on_page(img)
    finally:
        # ALWAYS release the scan job (even on error/interrupt) so the scanner is
        # free for the next scan. A job left open is the #1 cause of later "busy"/503.
        try:
            _u.urlopen(_u.Request(job, method="DELETE"), timeout=10)
        except Exception:
            pass

    if count == 0:
        # Sach-sach batao kya hua: jam / khaali feeder / kuch aur. (Pehle har
        # zero-page case par "feeder khaali" dikhta tha — jam me bhi.)
        st = ""
        try:
            rr = _u.urlopen(base + "/ScannerStatus", timeout=8)
            st = rr.read().decode("utf-8", "ignore").lower()
        except Exception:
            pass
        if "jam" in st:
            raise ScannerError(
                "Paper jam: ek page scanner ke andar atka hai. Use aaram se nikaalein, "
                "feeder me pages seedhe lagayein, phir dobara scan karein.")
        if "adfempty" in st.replace(" ", ""):
            raise ScannerError("PAPER_EMPTY: feeder khaali hai.")
        raise ScannerError(
            "eSCL: scanner ne koi page taiyar nahi kiya. Scanner ko ek baar off/on "
            "karke dobara try karein; problem rahe to Help > eSCL Test chalayein.")
    return count


# Role for storing a custom (document-name) label on a thumbnail item.
TITLE_ROLE = int(QtCore.Qt.UserRole) + 1
NAMEKEY_ROLE = int(QtCore.Qt.UserRole) + 2   # normalized key for remembering saved names


def underscore_name(s):
    """Filename-friendly: spaces -> underscore, drop odd chars, collapse repeats."""
    s = re.sub(r"[^\w\s.\-]", "", s or "")
    s = re.sub(r"\s+", "_", s.strip())
    s = re.sub(r"_+", "_", s)
    return s.strip("_.")


def name_key(s):
    """Normalized key to match the 'same' document title across scans."""
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


# Document TYPE classifier — recognise WHAT the document is (by tell-tale keywords)
# instead of naming it after whatever text is printed on top (which is usually just
# the hospital name and would be the same for every page). Tuned for ECHS/RGHS
# hospital paperwork. Order = priority when tie-broken.
DOC_TYPES = [
    # Only STRONG, distinctive phrases (avoid generic words that appear on every
    # hospital form like self/spouse/patient/hospital). If nothing here matches
    # clearly, the page stays "Page N" (better than a wrong name).
    ("Discharge_Summary", ["discharge summary"]),
    ("Bill",              ["cash receipt", "final bill", "bill of supply", "tax invoice",
                           "bill no.", "receipt no.", "total amount", "net payable"]),
    ("Referral",          ["referral form", "referral slip", "individual referred",
                           "referred hospital", "empanelled hospital"]),
    ("ECHS_Card",         ["semi-pvt", "echs card", "echs smart card", "ex-serviceman card",
                           "esm :"]),
    ("Aadhaar",           ["unique identification authority", "aadhaar", "आधार",
                           "government of india"]),
    ("Prescription",      ["provisional diagnosis", "chief complaint", "on examination",
                           "treatment advised", "diagnosis :", "complaints :"]),
    ("Lab_Report",        ["laboratory report", "pathology report", "haematology",
                           "biochemistry", "specimen", "sample collected", "test result"]),
    ("Xray_Radiology",    ["x-ray", "radiology report", "mri scan", "ct scan",
                           "ultrasound report", "sonography", "impression :"]),
    ("Registration",      ["registration no.", "uhid no", "op registration"]),
    ("Consent_Form",      ["consent form", "i hereby give my consent", "informed consent"]),
    ("Discharge_Card",    ["discharge card"]),
]


def page_ocr_text(path, frac=0.55):
    """OCR the top `frac` of a page and return the raw text (used by both the
    document-type classifier and the learned-name matcher)."""
    if not HAS_OCR_LIBS:
        return ""
    try:
        with Image.open(path) as im:
            w, h = im.size
            crop = im.crop((0, 0, w, max(1, int(h * frac)))).convert("L")
            if crop.width > 1500:
                r = 1500.0 / crop.width
                crop = crop.resize((1500, max(1, int(crop.height * r))))
        try:
            return pytesseract.image_to_string(crop, lang="eng+hin")
        except Exception:
            return pytesseract.image_to_string(crop)
    except Exception:
        return ""


def sig_words(text):
    """Distinctive words that identify a form's FIXED printed layout (letterhead,
    field labels, form title). Patient-specific data (names, numbers) mostly drops
    out, so two scans of the SAME form share most of these words."""
    words = re.findall(r"[a-z]{4,}", (text or "").lower())
    stop = {"name", "date", "time", "age", "years", "year", "male", "female",
            "address", "sign", "signature", "page", "self", "spouse", "mobile",
            "phone", "number", "the", "and", "for", "with", "this", "that"}
    return set(w for w in words if w not in stop)


def _jaccard(a, b):
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / union if union else 0.0


def learned_name_for(text, learned, thr=0.45):
    """learned = list of (wordset, name). Return the name whose stored signature
    best overlaps this page's words (>= thr), else None."""
    sig = sig_words(text)
    if not sig:
        return None
    best_name, best_score = None, 0.0
    for words, name in learned:
        sc = _jaccard(sig, words)
        if sc > best_score:
            best_score, best_name = sc, name
    return best_name if best_score >= thr else None


def classify_from_text(text):
    t = " " + re.sub(r"\s+", " ", (text or "").lower()) + " "
    scores = []
    for name, kws in DOC_TYPES:
        n = sum(1 for kw in kws if kw in t)
        if n > 0:
            scores.append((n, name))
    if not scores:
        return None
    scores.sort(reverse=True)
    # Accept only a CLEAR winner (top strictly beats runner-up); ties -> None.
    if len(scores) >= 2 and scores[0][0] == scores[1][0]:
        return None
    return scores[0][1]


def classify_document(path):
    """Return a clean document-TYPE name (e.g. Discharge_Summary, Bill, ECHS_Card)
    by matching keywords in the page text. None if no type is confidently found."""
    if not HAS_OCR_LIBS:
        return None
    return classify_from_text(page_ocr_text(path, 0.6))


def ocr_page_title(path):
    """Read the top of a scanned page and return a likely PRINTED document title
    (e.g. 'DISCHARGE SUMMARY'). Uses OCR word-confidence so low-confidence
    handwriting/garbage is rejected. None if nothing reliable is found."""
    if not HAS_OCR_LIBS:
        return None

    def _clean(ln):
        ln = re.sub(r"[^A-Za-z0-9 &.\-/,]", "", ln)
        return re.sub(r"\s+", " ", ln).strip()

    def _looks_garbage(ln):
        # a single alphabetic run longer than 15 chars with no space = OCR garbage
        for tok in ln.split():
            if tok.isalpha() and len(tok) > 15:
                return True
        return False

    def _title_from(fraction):
        try:
            with Image.open(path) as im:
                w, h = im.size
                strip = im.crop((0, 0, w, max(1, int(h * fraction)))).convert("L")
                if strip.width > 1400:
                    r = 1400.0 / strip.width
                    strip = strip.resize((1400, max(1, int(strip.height * r))))
            # Confidence-aware: build lines only from words OCR is confident about.
            try:
                try:
                    data = pytesseract.image_to_data(strip, lang="eng+hin",
                                                     output_type=pytesseract.Output.DICT)
                except Exception:
                    data = pytesseract.image_to_data(strip,
                                                     output_type=pytesseract.Output.DICT)
                lines = {}
                for i in range(len(data.get("text", []))):
                    txt = (data["text"][i] or "").strip()
                    try:
                        conf = float(data["conf"][i])
                    except Exception:
                        conf = -1
                    if not txt or conf < 55:
                        continue
                    if sum(c.isalnum() for c in txt) < 2:
                        continue
                    key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
                    lines.setdefault(key, []).append((txt, conf))
                cands = []
                for words in lines.values():
                    text = _clean(" ".join(wtc[0] for wtc in words))
                    if len(text) < 4 or _looks_garbage(text):
                        continue
                    avg = sum(wtc[1] for wtc in words) / len(words)
                    cands.append((text, avg))
                if cands:
                    def score(tc):
                        ln, cf = tc
                        letters = sum(c.isalpha() for c in ln)
                        caps = sum(1 for c in ln if c.isupper())
                        return cf * 0.5 + letters + caps * 0.4 - abs(len(ln) - 22) * 0.05
                    best = max(cands, key=score)[0]
                    if sum(c.isalpha() for c in best) >= 3:
                        return best[:38]
            except Exception:
                pass
            return None
        except Exception:
            return None

    return _title_from(0.22) or _title_from(0.45)


class NameWorker(QtCore.QThread):
    named = QtCore.pyqtSignal(int, str)   # (row, title)
    finished_all = QtCore.pyqtSignal()

    def __init__(self, items, learned=None):
        super().__init__()
        self.items = items                # list of (row, path)
        self.learned = learned or []      # list of (wordset, name)

    def run(self):
        for row, path in self.items:
            if self.isInterruptionRequested():
                break
            text = page_ocr_text(path, 0.6)
            # priority: name YOU taught it > document-type > printed heading
            title = (learned_name_for(text, self.learned)
                     or classify_from_text(text)
                     or ocr_page_title(path))
            if title:
                self.named.emit(row, title)
        self.finished_all.emit()


class FuncWorker(QtCore.QThread):
    """Koi bhi bhaari kaam (OCR/network) background me chala kar result de —
    UI kabhi nahi rukti. Result ya Exception, dono `done` signal me aate hain."""
    done = QtCore.pyqtSignal(object)

    def __init__(self, fn):
        super().__init__()
        self.fn = fn

    def run(self):
        try:
            self.done.emit(self.fn())
        except Exception as e:
            self.done.emit(e)


class LearnWorker(QtCore.QThread):
    """Rename ke baad naam 'seekhne' ka OCR background me — pehle ye UI thread
    par chalta tha aur har rename par app 3-10 second jam jaati thi."""
    got = QtCore.pyqtSignal(str, list)   # (name, words)

    def __init__(self, path, name):
        super().__init__()
        self.path = path
        self.name = name

    def run(self):
        try:
            words = sorted(sig_words(page_ocr_text(self.path, 0.6)))
        except Exception:
            words = []
        self.got.emit(self.name, words)


class ImportWorker(QtCore.QThread):
    """Import (images/PDF/photo-cleanup) background me — UI kabhi nahi rukti.
    Thumbnail bhi yahin chhota (QImage worker thread me safe hai)."""
    page_ready = QtCore.pyqtSignal(str, QtGui.QImage)
    done = QtCore.pyqtSignal(int)

    def __init__(self, files, tmpdir, mode="normal", thumb_w=360, thumb_h=480):
        super().__init__()
        self.files = list(files)
        self.tmpdir = tmpdir
        self.mode = mode          # "normal" ya "photo" (phone-photo cleanup)
        self.thumb_w = thumb_w
        self.thumb_h = thumb_h

    def _thumb(self, path):
        r = QtGui.QImageReader(path)
        r.setAutoTransform(True)
        sz = r.size()
        if sz.isValid() and sz.width() > 0 and sz.height() > 0:
            s = min(float(self.thumb_w) / sz.width(),
                    float(self.thumb_h) / sz.height(), 1.0)
            r.setScaledSize(QtCore.QSize(max(1, int(sz.width() * s)),
                                         max(1, int(sz.height() * s))))
        return r.read()

    def run(self):
        count = 0
        for f in self.files:
            if self.isInterruptionRequested():
                break
            try:
                if self.mode == "photo":
                    with Image.open(f) as im:
                        img = clean_photo(im)
                    fd, out = tempfile.mkstemp(suffix=".jpg", dir=self.tmpdir)
                    os.close(fd)
                    img.save(out, "JPEG", quality=90)
                    outs = [out]
                elif f.lower().endswith(".pdf"):
                    outs = pdf_to_images(f, self.tmpdir) or []
                else:
                    with Image.open(f) as im:
                        fd, out = tempfile.mkstemp(suffix=".png", dir=self.tmpdir)
                        os.close(fd)
                        im.convert("RGB").save(out, "PNG")
                    outs = [out]
                for p in outs:
                    count += 1
                    self.page_ready.emit(p, self._thumb(p))
            except Exception:
                pass
        self.done.emit(count)


class ScanWorker(QtCore.QThread):
    page_done = QtCore.pyqtSignal(str)
    done = QtCore.pyqtSignal(int, int)
    failed = QtCore.pyqtSignal(str)

    def __init__(self, hwnd, source_name, dpi, pixel_type, duplex, tmpdir, opts):
        super().__init__()
        self.hwnd = hwnd
        self.source_name = source_name
        self.dpi = dpi
        self.pixel_type = pixel_type
        self.duplex = duplex
        self.tmpdir = tmpdir
        self.opts = opts or {}
        self.kept = 0
        self.skipped = 0

    def run(self):
        import queue as _q, threading as _th
        pageq = _q.Queue()

        def _consumer():
            # Runs in parallel: saves/processes each page while the scanner
            # is already pulling the NEXT one (this removes the per-page pause).
            while True:
                img = pageq.get()
                if img is None:
                    break
                try:
                    if self.opts.get("remove_blank"):
                        _sens = {"kam": 0.0003, "normal": 0.0008, "zyada": 0.004}
                        _thr = _sens.get(self.opts.get("blank_sensitivity", "normal"), 0.0008)
                        if is_blank_page(img, _thr):
                            self.skipped += 1
                            continue
                    # Backing (kala/navy/gray) HAMESHA white karo — kisi bhi page
                    # size me sheet chhoti ho to backing dikhti hai; print me
                    # kala bilkul nahi aana chahiye.
                    img = whiten_dark_background(img)
                    if self.opts.get("auto_crop"):
                        img = autocrop(img)
                    if self.opts.get("deskew"):
                        img = deskew(img)
                    if self.opts.get("quality_enhance"):
                        img = auto_enhance(img)
                    if self.opts.get("auto_orient") and tesseract_available():
                        img = auto_orient(img)
                    # Rang-heen page ko gray bana do (chhoti file) — sirf colour
                    # scan me, aur sirf jab option ON ho.
                    if (self.opts.get("auto_colour") and self.pixel_type == "color"
                            and colorfulness(img) < 6.0):
                        img = img.convert("L")
                    # Save the page. JPEG encodes ~5x faster than PNG (big scan-speed
                    # win); use it for colour/grey. Keep PNG for 1-bit black&white.
                    if self.pixel_type == "bw":
                        fd, out = tempfile.mkstemp(suffix=".png", dir=self.tmpdir)
                        os.close(fd)
                        # eSCL me B&W ab grayscale me aata hai (BlackAndWhite1+JPEG
                        # scanner par jam karta tha) — asli 1-bit yahan banao.
                        try:
                            if img.mode != "1":
                                img = img.convert("L").point(
                                    lambda v: 255 if v >= 160 else 0, mode="1")
                        except Exception:
                            pass
                        try:
                            img.save(out, "PNG", compress_level=1)
                        except Exception:
                            img.save(out, "PNG")
                    else:
                        fd, out = tempfile.mkstemp(suffix=".jpg", dir=self.tmpdir)
                        os.close(fd)
                        try:
                            if img.mode == "L":
                                img.save(out, "JPEG", quality=88)
                            else:
                                img.convert("RGB").save(out, "JPEG", quality=88)
                        except Exception:
                            img.convert("RGB").save(out, "JPEG", quality=85)
                    self.kept += 1
                    self.page_done.emit(out)
                except Exception:
                    pass

        consumer = _th.Thread(target=_consumer, daemon=True)
        consumer.start()

        def _on_page(img):
            # Producer: hand the page off instantly and let the scanner keep going.
            pageq.put(img)

        err = None
        method = self.opts.get("scanner_method", "twain")
        try:
            _ps = str(self.opts.get("page_size", "auto"))
            if _ps.startswith("custom"):
                _ps = "custom:%d" % int(self.opts.get("custom_page_mm", 600) or 600)
            if method == "escl":
                scan_via_escl(self.opts.get("scanner_ip"), self.dpi, self.pixel_type,
                              self.duplex, _on_page, should_stop=self.isInterruptionRequested,
                              page_size=_ps)
            elif method == "naps2":
                scan_via_naps2(self.opts.get("naps2_path") or find_naps2(),
                               self.opts.get("naps2_profile"), self.tmpdir,
                               self.duplex, _on_page,
                               should_stop=self.isInterruptionRequested)
            elif method == "wia":
                try:
                    _pythoncom.CoInitialize()
                except Exception:
                    pass
                try:
                    wia_scan_pages(self.opts.get("wia_device_id"), self.dpi,
                                   self.pixel_type, self.duplex, _on_page,
                                   should_stop=self.isInterruptionRequested)
                finally:
                    try:
                        _pythoncom.CoUninitialize()
                    except Exception:
                        pass
            else:
                scan_pages(self.hwnd, self.source_name, self.dpi,
                           self.pixel_type, self.duplex, _on_page,
                           should_stop=self.isInterruptionRequested,
                           file_xfer=bool(self.opts.get("twain_file_xfer")))
        except ScannerError as exc:
            err = str(exc)
        except Exception:
            err = "Scan error:\n%s" % traceback.format_exc()

        pageq.put(None)            # tell the saver thread to finish
        consumer.join(timeout=60)  # wait until all pages are saved

        if err is not None and self.kept == 0:
            self.failed.emit(err)
            return
        if self.kept == 0 and self.skipped == 0:
            self.failed.emit(err or "Koi page scan nahi hua.")
            return
        self.done.emit(self.kept, self.skipped)


# ---------------------------------------------------------------------------
# Dialogs
# ---------------------------------------------------------------------------

class EditProfileDialog(QtWidgets.QDialog):
    def __init__(self, parent, profile=None):
        super().__init__(parent)
        self.setWindowTitle("Profile settings")
        self.setMinimumWidth(420)
        self.profile = dict(profile) if profile else {
            "name": "", "source_name": None, "dpi": 200, "color": "gray", "duplex": False}
        form = QtWidgets.QFormLayout(self)

        def qh(text, tip):
            # "?" help-icon wala row-label (Hindi + English hover)
            h = QtWidgets.QLabel("?")
            h.setToolTip(tip); h.setCursor(QtCore.Qt.WhatsThisCursor)
            h.setStyleSheet("QLabel{color:#0f766e;border:1px solid #0f766e;border-radius:9px;"
                            "min-width:18px;max-width:18px;min-height:18px;max-height:18px;"
                            "font-weight:bold;qproperty-alignment:AlignCenter;}")
            r = QtWidgets.QHBoxLayout(); r.setContentsMargins(0, 0, 0, 0); r.setSpacing(6)
            r.addWidget(QtWidgets.QLabel(text)); r.addWidget(h); r.addStretch(1)
            w = QtWidgets.QWidget(); w.setLayout(r); return w
        self.name_edit = QtWidgets.QLineEdit(self.profile.get("name", ""))
        self.name_edit.setPlaceholderText("jaise: Documents fast")
        self.name_edit.setToolTip("हिन्दी: Is profile ka naam (jaise 'Fast B&W', 'Colour A4'). Toolbar me isi naam se chunoge.\nEnglish: A name for this profile (e.g. 'Fast B&W'); you'll pick it by this name.")
        form.addRow(qh("Display Name:", self.name_edit.toolTip()), self.name_edit)
        dev_row = QtWidgets.QHBoxLayout()
        self.device_label = QtWidgets.QLabel(self.profile.get("source_name") or "(koi device nahi)")
        btn_dev = QtWidgets.QPushButton("Choose device")
        btn_dev.clicked.connect(self._choose_device)
        dev_row.addWidget(self.device_label, 1); dev_row.addWidget(btn_dev)
        dw = QtWidgets.QWidget(); dw.setLayout(dev_row)
        form.addRow(qh("Device:", "हिन्दी: Kaunsa scanner is profile me use hoga (TWAIN device chuno).\nEnglish: Which scanner this profile uses (choose a TWAIN device)."), dw)
        self.cmb_dpi = QtWidgets.QComboBox(); self.cmb_dpi.addItems(RESOLUTIONS)
        self.cmb_dpi.setCurrentText(str(self.profile.get("dpi", 200)))
        form.addRow(qh("Resolution (DPI):", "हिन्दी: Scan ki safai. 200 = tez aur text ke liye theek; 300 = behtar (dhima); 600 = photo ke liye.\nEnglish: Scan sharpness. 200 = fast, fine for text; 300 = better (slower); 600 = for photos."), self.cmb_dpi)
        self.cmb_color = QtWidgets.QComboBox(); self.cmb_color.addItems(list(COLOUR_MODES.keys()))
        for label, code in COLOUR_MODES.items():
            if code == self.profile.get("color", "gray"):
                self.cmb_color.setCurrentText(label)
        form.addRow(qh("Colour mode:", "हिन्दी: Rangeen = colour, Grayscale = kaala-safed shades, B&W = sirf kaala/safed (sabse chhoti file, tez).\nEnglish: Colour, Grayscale, or pure Black & White (smallest & fastest)."), self.cmb_color)
        self.chk_duplex = QtWidgets.QCheckBox("Duplex (dono taraf)")
        self.chk_duplex.setChecked(bool(self.profile.get("duplex")))
        self.chk_duplex.setToolTip("हिन्दी: ON: kaagaz ke dono taraf apne aap scan (duplex ADF scanner me).\nEnglish: ON: scan both sides of the paper automatically (duplex ADF).")
        _dxr = QtWidgets.QHBoxLayout(); _dxr.setContentsMargins(0, 0, 0, 0); _dxr.setSpacing(6)
        _dxh = QtWidgets.QLabel("?"); _dxh.setToolTip(self.chk_duplex.toolTip())
        _dxh.setCursor(QtCore.Qt.WhatsThisCursor)
        _dxh.setStyleSheet("QLabel{color:#0f766e;border:1px solid #0f766e;border-radius:9px;"
                           "min-width:18px;max-width:18px;min-height:18px;max-height:18px;"
                           "font-weight:bold;qproperty-alignment:AlignCenter;}")
        _dxr.addWidget(self.chk_duplex); _dxr.addWidget(_dxh); _dxr.addStretch(1)
        _dxw = QtWidgets.QWidget(); _dxw.setLayout(_dxr)
        form.addRow("", _dxw)
        self.cmb_psize = QtWidgets.QComboBox()
        self._PSIZES = [("Auto (alag-alag size khud pakde)", "auto"),
                        ("A4 (210x297 mm)", "a4"), ("Letter", "letter"),
                        ("Legal", "legal"), ("A5", "a5"),
                        ("Lambi parchi / custom (length Settings me)", "custom")]
        self.cmb_psize.addItems([t for t, _c in self._PSIZES])
        _ps = (self.profile.get("page_size") or "auto").lower()
        _idx = next((k for k, (_t, c) in enumerate(self._PSIZES) if _ps.startswith(c[:4])), 0)
        self.cmb_psize.setCurrentIndex(_idx)
        form.addRow(qh("Page size:", "हिन्दी: Auto = har page ki asli size khud pakde (mixed/ID/aadha page bhi poora). A4/Letter/Legal/A5 = fixed. Custom = lambi parchi (length Settings me).\nEnglish: Auto = detect each page's real size; A4/Letter/Legal/A5 = fixed; Custom = long receipts."), self.cmb_psize)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self._ok); btns.rejected.connect(self.reject)
        form.addRow(btns)

    def _choose_device(self):
        if not HAS_TWAIN:
            QtWidgets.QMessageBox.warning(self, "Error", "TWAIN install nahi hai."); return
        try:
            names = list_sources(int(self.winId()))
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Error", "Scanner list nahi mili:\n%s" % exc); return
        if not names:
            QtWidgets.QMessageBox.warning(self, "Error", "Koi scanner nahi mila."); return
        name, ok = QtWidgets.QInputDialog.getItem(self, "Device chuno", "Scanner:", names, 0, False)
        if ok and name:
            self.profile["source_name"] = name; self.device_label.setText(name)

    def _ok(self):
        if not self.name_edit.text().strip():
            QtWidgets.QMessageBox.warning(self, "Error", "Display Name daalein."); return
        self.accept()

    def get_profile(self):
        self.profile["name"] = self.name_edit.text().strip() or "Profile"
        self.profile["dpi"] = int(self.cmb_dpi.currentText())
        self.profile["color"] = COLOUR_MODES[self.cmb_color.currentText()]
        self.profile["duplex"] = self.chk_duplex.isChecked()
        self.profile["page_size"] = self._PSIZES[self.cmb_psize.currentIndex()][1]
        return self.profile


class ProfileManagerDialog(QtWidgets.QDialog):
    def __init__(self, parent, profiles):
        super().__init__(parent)
        self.setWindowTitle("Profiles"); self.resize(420, 320)
        self.profiles = [dict(p) for p in profiles]
        lay = QtWidgets.QVBoxLayout(self)
        lay.addWidget(QtWidgets.QLabel("Scan profiles:"))
        self.listw = QtWidgets.QListWidget()
        self.listw.itemDoubleClicked.connect(lambda *_: self._edit())
        lay.addWidget(self.listw, 1); self._refresh()
        row = QtWidgets.QHBoxLayout()
        for t, s in [("New", self._new), ("Edit", self._edit), ("Delete", self._delete)]:
            b = QtWidgets.QPushButton(t); b.clicked.connect(s); row.addWidget(b)
        row.addStretch(1)
        done = QtWidgets.QPushButton("Done"); done.setObjectName("primary")
        done.clicked.connect(self.accept); row.addWidget(done)
        lay.addLayout(row)

    def _refresh(self):
        self.listw.clear()
        for p in self.profiles:
            self.listw.addItem(p.get("name", "Profile"))

    def _new(self):
        dlg = EditProfileDialog(self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self.profiles.append(dlg.get_profile()); self._refresh()
            self.listw.setCurrentRow(len(self.profiles) - 1)

    def _edit(self):
        i = self.listw.currentRow()
        if i < 0:
            return
        dlg = EditProfileDialog(self, self.profiles[i])
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self.profiles[i] = dlg.get_profile(); self._refresh(); self.listw.setCurrentRow(i)

    def _delete(self):
        i = self.listw.currentRow()
        if i < 0:
            return
        del self.profiles[i]; self._refresh()


class OptionsDialog(QtWidgets.QDialog):
    def __init__(self, parent, opts):
        super().__init__(parent)
        self.setWindowTitle("Options / Settings")
        self.setMinimumWidth(560)
        self.opts = dict(DEFAULT_OPTIONS); self.opts.update(opts or {})

        outer = QtWidgets.QVBoxLayout(self)
        scroll = QtWidgets.QScrollArea(); scroll.setWidgetResizable(True)
        inner = QtWidgets.QWidget(); form = QtWidgets.QFormLayout(inner)
        scroll.setWidget(inner); outer.addWidget(scroll, 1)

        def header(t):
            lbl = QtWidgets.QLabel("<b>%s</b>" % t); form.addRow(lbl)

        def qhelp(tip):
            h = QtWidgets.QLabel("?")
            h.setToolTip(tip); h.setCursor(QtCore.Qt.WhatsThisCursor)
            h.setStyleSheet("QLabel{color:#0f766e; border:1px solid #0f766e; border-radius:9px;"
                            "min-width:18px; max-width:18px; min-height:18px; max-height:18px;"
                            "font-weight:bold; qproperty-alignment:AlignCenter;}")
            return h

        def chkrow(chk, tip):
            chk.setToolTip(tip)
            row = QtWidgets.QHBoxLayout(); row.setContentsMargins(0, 0, 0, 0); row.setSpacing(6)
            row.addWidget(chk); row.addWidget(qhelp(tip)); row.addStretch(1)
            w = QtWidgets.QWidget(); w.setLayout(row); return w

        def lblhelp(text, tip):
            row = QtWidgets.QHBoxLayout(); row.setContentsMargins(0, 0, 0, 0); row.setSpacing(6)
            row.addWidget(QtWidgets.QLabel(text)); row.addWidget(qhelp(tip)); row.addStretch(1)
            w = QtWidgets.QWidget(); w.setLayout(row); return w

        header("Interface")
        self.cmb_theme = QtWidgets.QComboBox(); self.cmb_theme.addItems(["Light", "Dark", "Dark Pro"])
        self.cmb_theme.setCurrentIndex({"light": 0, "dark": 1, "darkpro": 2}.get(self.opts.get("theme"), 0))
        form.addRow(lblhelp("Theme:", 'हिन्दी: App ka rang-roop. Dark = kaala theme (aankhon ko aaram, raat me achha). Light = safed.\nEnglish: App look. Dark = dark theme (easy on eyes at night). Light = white.'), self.cmb_theme)
        self.chk_pagenum = QtWidgets.QCheckBox("Page number thumbnails ke neeche dikhao")
        self.chk_pagenum.setChecked(self.opts.get("show_page_numbers", True)); form.addRow(chkrow(self.chk_pagenum, "हिन्दी: Har thumbnail ke neeche 'Page 1, 2..' dikhana. Band karne par number nahi dikhenge.\nEnglish: Show 'Page 1, 2..' under each thumbnail. Off = no numbers."))
        self.chk_autoname = QtWidgets.QCheckBox("Document ka naam apne aap padho (OCR)")
        self.chk_autoname.setChecked(self.opts.get("auto_name", False))
        form.addRow(chkrow(self.chk_autoname, "हिन्दी: ON: scan ke baad har page ka naam (document ka title, jaise DISCHARGE SUMMARY / RECEIPT) apne aap padh kar likh dega — 'Page 1,2' ke bajay. (OCR/Tesseract chahiye.)\nEnglish: ON: after scanning, auto-labels each page with its document title (needs OCR)."))

        header("Save")
        self.chk_autosave = QtWidgets.QCheckBox("Scan ke baad PDF apne aap save karo")
        self.chk_autosave.setChecked(self.opts["auto_save"]); form.addRow(chkrow(self.chk_autosave, 'हिन्दी: ON: scan khatam hote hi PDF apne aap save ho jayegi (har baar save dabana nahi padega). Roz bahut scan karte ho to ON rakho.\nEnglish: ON: PDF saves automatically after each scan. Handy if you scan a lot.'))
        fr = QtWidgets.QHBoxLayout()
        self.folder_edit = QtWidgets.QLineEdit(self.opts["save_folder"])
        b = QtWidgets.QPushButton("…"); b.setFixedWidth(36); b.clicked.connect(self._pick_folder)
        fr.addWidget(self.folder_edit, 1); fr.addWidget(b)
        fw = QtWidgets.QWidget(); fw.setLayout(fr); form.addRow(lblhelp("Save folder:", 'हिन्दी: PDF kahaan save hon wo folder.\nEnglish: Folder where PDFs get saved.'), fw)
        self.tmpl_edit = QtWidgets.QLineEdit(self.opts["filename_template"])
        form.addRow(lblhelp("Filename template:", 'हिन्दी: File ka naam kaise bane. {claim}=claim number, {date}=tareekh, {time}=samay, {seq}=kram sankhya.\nEnglish: How filenames are built: {claim} {date} {time} {seq}.'), self.tmpl_edit)
        form.addRow("", QtWidgets.QLabel("Tags: {claim} {date} {time} {seq}"))
        self.chk_claimfolder = QtWidgets.QCheckBox("Claim number ka alag folder")
        self.chk_claimfolder.setChecked(self.opts["make_claim_folder"]); form.addRow(chkrow(self.chk_claimfolder, 'हिन्दी: ON: har claim number ka alag folder banega (ek claim ke saare pages ek jagah).\nEnglish: ON: a separate folder per claim number.'))
        self.chk_ymfolder = QtWidgets.QCheckBox("Saal/Mahine ke folder (2026/07/...)")
        self.chk_ymfolder.setChecked(self.opts["year_month_folders"]); form.addRow(chkrow(self.chk_ymfolder, 'हिन्दी: ON: saal/mahine ke folder (2026/07/...) — purane scans aasani se milen.\nEnglish: ON: year/month folders (2026/07/...) for easy filing.'))
        self.cmb_after = QtWidgets.QComboBox(); self.cmb_after.addItems(["Kuch nahi", "PDF kholo", "Folder kholo"])
        self.cmb_after.setCurrentIndex({"nothing": 0, "open": 1, "folder": 2}.get(self.opts.get("after_save", "nothing"), 0))
        form.addRow(lblhelp("Save ke baad:", 'हिन्दी: Save ke baad kya ho — kuch nahi / PDF khule / folder khule.\nEnglish: After save — do nothing / open the PDF / open the folder.'), self.cmb_after)
        self.chk_imgtoo = QtWidgets.QCheckBox("Har page ki alag image (JPG) bhi save karo")
        self.chk_imgtoo.setChecked(self.opts.get("save_images_too", False)); form.addRow(chkrow(self.chk_imgtoo, 'हिन्दी: ON: PDF ke saath har page ki alag JPG image bhi banegi.\nEnglish: ON: also save each page as a separate JPG image.'))

        header("Image cleanup")
        self.chk_blank = QtWidgets.QCheckBox("Khaali (blank) pages hatao")
        self.chk_blank.setChecked(self.opts["remove_blank"]); form.addRow(chkrow(self.chk_blank, 'हिन्दी: ON: khaali (blank) pages apne aap hat jayenge — jaise duplex me peeche ka khaali side. (NAPS2 jaisa.)\nEnglish: ON: blank pages are removed automatically (e.g. blank back side in duplex).'))
        self.cmb_blank_sens = QtWidgets.QComboBox(); self.cmb_blank_sens.addItems(["Kam (safe)", "Normal", "Zyada (aggressive)"])
        self.cmb_blank_sens.setCurrentIndex({"kam": 0, "normal": 1, "zyada": 2}.get(self.opts.get("blank_sensitivity", "normal"), 1))
        form.addRow(lblhelp("Blank hatane ki sensitivity:", 'हिन्दी: Kitni sakhti se blank hataye. "Zyada" = fold-line/halke stamp wale peeche ke khaali page bhi hat jayenge (par kabhi-kabhi kam content wala asli page bhi hat sakta hai). "Kam" = sirf bilkul khaali. Normal beech ka.\nEnglish: How aggressively to drop blanks. Zyada = also removes back sides with fold lines/faint marks; Kam = only truly empty; Normal = balanced.'), self.cmb_blank_sens)
        self.chk_crop = QtWidgets.QCheckBox("Border auto-crop")
        self.chk_crop.setChecked(self.opts["auto_crop"]); form.addRow(chkrow(self.chk_crop, 'हिन्दी: ON: page ke aas-paas ki khaali safed border apne aap kat jayegi.\nEnglish: ON: auto-trims the empty white border around the page.'))
        self.chk_deskew = QtWidgets.QCheckBox("Auto-deskew (tedha seedha)")
        self.chk_deskew.setChecked(self.opts["deskew"])
        if not HAS_NUMPY:
            self.chk_deskew.setEnabled(False)
            self.chk_deskew.setText("Auto-deskew (numpy install karein)")
        form.addRow(chkrow(self.chk_deskew, 'हिन्दी: ON: tedha scan hua page apne aap seedha ho jayega.\nEnglish: ON: straightens a tilted/skewed page automatically.'))
        self.chk_enhance = QtWidgets.QCheckBox("Quality auto-sudhar (faded documents saaf)")
        self.chk_enhance.setChecked(self.opts["quality_enhance"]); form.addRow(chkrow(self.chk_enhance, 'हिन्दी: ON: feeke/halke documents saaf aur gehre dikhenge.\nEnglish: ON: brightens & sharpens faded documents.'))
        self.chk_orient = QtWidgets.QCheckBox("Ulta page khud seedha karo (OCR se)")
        self.chk_orient.setChecked(bool(self.opts.get("auto_orient")))
        form.addRow(chkrow(self.chk_orient, 'हिन्दी: ON: ulta (90/180/270) scan hua page OCR se pehchan kar apne aap seedha ho jayega (Tesseract chahiye; scan thoda dheema hota hai).\nEnglish: ON: auto-rotates upside-down/sideways pages using OCR (needs Tesseract; slightly slower).'))
        self.chk_autocolour = QtWidgets.QCheckBox("Auto colour-detect (rang-heen page gray me)")
        self.chk_autocolour.setChecked(bool(self.opts.get("auto_colour")))
        form.addRow(chkrow(self.chk_autocolour, 'हिन्दी: ON (sirf Colour scan me): jis page par rang nahi hai wo apne aap gray me save hoga — chhoti file, saaf print.\nEnglish: ON (colour scans only): pages with no real colour are saved as grayscale — smaller files.'))
        self.spin_custlen = QtWidgets.QSpinBox(); self.spin_custlen.setRange(100, 3000)
        self.spin_custlen.setValue(int(self.opts.get("custom_page_mm", 600) or 600))
        self.spin_custlen.setSuffix(" mm")
        form.addRow(lblhelp("Custom page ki lambai:", 'हिन्दी: Profile me page size "Lambi parchi/custom" chunne par itni lambi patti scan hogi (receipt/bahi-khata).\nEnglish: Length used when the profile page size is "custom" (long receipts).'), self.spin_custlen)
        self.chk_filexfer = QtWidgets.QCheckBox("TWAIN continuous feed (experimental)")
        self.chk_filexfer.setChecked(bool(self.opts.get("twain_file_xfer")))
        form.addRow(chkrow(self.chk_filexfer, 'हिन्दी: ON (sirf TWAIN method me): scanner ka feeder bina ruke chalta rahega (file-transfer mode). Agar scanner support na kare to app khud purane tareeke par aa jayegi.\nEnglish: ON (TWAIN method only): keeps the ADF feeding continuously via file-transfer mode; falls back to the normal path automatically if unsupported.'))

        header("Output")
        self.chk_compress = QtWidgets.QCheckBox("PDF compress (chhoti file)")
        self.chk_compress.setChecked(self.opts["compress"]); form.addRow(chkrow(self.chk_compress, 'हिन्दी: ON: PDF ki file chhoti banegi (email/upload aasan); quality thodi kam.\nEnglish: ON: smaller PDF (easier to email/upload); slightly lower quality.'))
        self.q_spin = QtWidgets.QSpinBox(); self.q_spin.setRange(20, 95)
        self.q_spin.setValue(int(self.opts["jpeg_quality"]))
        form.addRow(lblhelp("Compress quality:", 'हिन्दी: Compress ki quality (zyada = behtar pic par badi file). 60-80 theek hai.\nEnglish: Compression quality (higher = better image, larger file). 60-80 is fine.'), self.q_spin)
        self.chk_wm = QtWidgets.QCheckBox("Watermark/stamp har page par")
        self.chk_wm.setChecked(self.opts["watermark"]); form.addRow(chkrow(self.chk_wm, 'हिन्दी: ON: har page par aapka text/stamp (jaise hospital ka naam) chhapega.\nEnglish: ON: stamps your text (e.g. hospital name) on every page.'))
        self.wm_edit = QtWidgets.QLineEdit(self.opts["watermark_text"])
        form.addRow(lblhelp("Watermark text:", 'हिन्दी: Watermark me kya likha ho.\nEnglish: The watermark text.'), self.wm_edit)

        header("Workflow")
        self.chk_batch = QtWidgets.QCheckBox("Batch mode (save ke baad next claim ready)")
        self.chk_batch.setChecked(self.opts["batch_mode"]); form.addRow(chkrow(self.chk_batch, 'हिन्दी: ON: ek claim save hote hi agla claim scan karne ke liye screen saaf ho jayegi (tezi se ek ke baad ek).\nEnglish: ON: after saving, screen clears for the next claim (fast back-to-back).'))
        self.chk_validate = QtWidgets.QCheckBox("Claim number validate karo")
        self.chk_validate.setChecked(self.opts["validate_claim"]); form.addRow(chkrow(self.chk_validate, 'हिन्दी: ON: galat/adhoora claim number daalne par chetavni dega.\nEnglish: ON: warns if the claim number looks wrong/incomplete.'))
        self.pat_edit = QtWidgets.QLineEdit(self.opts["claim_pattern"])
        form.addRow(lblhelp("Claim pattern (regex):", 'हिन्दी: Claim number ka sahi roop (regex). Aam taur par badalne ki zaroorat nahi.\nEnglish: The valid claim-number pattern (regex). Usually leave as is.'), self.pat_edit)
        self.chk_barcode = QtWidgets.QCheckBox("Barcode/QR se claim number khud bharo")
        self.chk_barcode.setChecked(self.opts["barcode_autofill"])
        if not HAS_ZBAR:
            self.chk_barcode.setEnabled(False)
            self.chk_barcode.setText("Barcode/QR (pyzbar install karein)")
        form.addRow(chkrow(self.chk_barcode, 'हिन्दी: ON: page par barcode/QR ho to claim number apne aap bhar jayega.\nEnglish: ON: auto-fills the claim number from a barcode/QR on the page.'))
        self.chk_dup = QtWidgets.QCheckBox("Duplicate claim number par chetavni")
        self.chk_dup.setChecked(self.opts["duplicate_check"]); form.addRow(chkrow(self.chk_dup, 'हिन्दी: ON: wahi claim number dobara ho to chetavni (double entry se bachav).\nEnglish: ON: warns on a duplicate claim number.'))

        header("Records & safety")
        self.chk_excel = QtWidgets.QCheckBox("Excel register me entry (register.xlsx)")
        self.chk_excel.setChecked(self.opts["excel_log"])
        if not HAS_XLSX:
            self.chk_excel.setEnabled(False)
            self.chk_excel.setText("Excel register (openpyxl install karein)")
        form.addRow(chkrow(self.chk_excel, 'हिन्दी: ON: har scan ki entry ek Excel register (register.xlsx) me judegi — record ke liye.\nEnglish: ON: logs each scan into an Excel register (register.xlsx).'))
        self.chk_log = QtWidgets.QCheckBox("Activity log (activity_log.txt)")
        self.chk_log.setChecked(self.opts["activity_log"]); form.addRow(chkrow(self.chk_log, 'हिन्दी: ON: kab kya scan/save hua iska text log rakhega.\nEnglish: ON: keeps a text activity log.'))
        self.chk_backup = QtWidgets.QCheckBox("Har save ka backup copy")
        self.chk_backup.setChecked(self.opts["backup"]); form.addRow(chkrow(self.chk_backup, 'हिन्दी: ON: har save ki ek backup copy alag folder me bhi rakhega (surakhsha).\nEnglish: ON: keeps a backup copy of every save in another folder.'))
        br = QtWidgets.QHBoxLayout()
        self.backup_edit = QtWidgets.QLineEdit(self.opts["backup_folder"])
        bb = QtWidgets.QPushButton("…"); bb.setFixedWidth(36); bb.clicked.connect(self._pick_backup)
        br.addWidget(self.backup_edit, 1); br.addWidget(bb)
        bw = QtWidgets.QWidget(); bw.setLayout(br); form.addRow(lblhelp("Backup folder:", 'हिन्दी: Backup kahaan rakhe wo folder.\nEnglish: Folder for backups.'), bw)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject)
        outer.addWidget(btns)

    def _pick_folder(self):
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "Folder", self.folder_edit.text())
        if d:
            self.folder_edit.setText(d)

    def _pick_backup(self):
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "Backup folder", self.backup_edit.text())
        if d:
            self.backup_edit.setText(d)

    def get_opts(self):
        o = self.opts
        o["theme"] = ["light", "dark", "darkpro"][self.cmb_theme.currentIndex()]
        o["show_page_numbers"] = self.chk_pagenum.isChecked()
        o["auto_name"] = self.chk_autoname.isChecked()
        o["save_images_too"] = self.chk_imgtoo.isChecked()
        o["auto_save"] = self.chk_autosave.isChecked()
        o["save_folder"] = self.folder_edit.text().strip()
        o["filename_template"] = self.tmpl_edit.text().strip() or "{date}_{seq}"
        o["make_claim_folder"] = self.chk_claimfolder.isChecked()
        o["after_save"] = ["nothing", "open", "folder"][self.cmb_after.currentIndex()]
        o["year_month_folders"] = self.chk_ymfolder.isChecked()
        o["remove_blank"] = self.chk_blank.isChecked()
        o["blank_sensitivity"] = {0: "kam", 1: "normal", 2: "zyada"}.get(self.cmb_blank_sens.currentIndex(), "normal")
        o["auto_crop"] = self.chk_crop.isChecked()
        o["deskew"] = self.chk_deskew.isChecked()
        o["quality_enhance"] = self.chk_enhance.isChecked()
        o["twain_file_xfer"] = self.chk_filexfer.isChecked()
        o["auto_orient"] = self.chk_orient.isChecked()
        o["auto_colour"] = self.chk_autocolour.isChecked()
        o["custom_page_mm"] = int(self.spin_custlen.value())
        o["compress"] = self.chk_compress.isChecked()
        o["jpeg_quality"] = int(self.q_spin.value())
        o["watermark"] = self.chk_wm.isChecked()
        o["watermark_text"] = self.wm_edit.text().strip() or "Noble Care Hospital"
        o["batch_mode"] = self.chk_batch.isChecked()
        o["validate_claim"] = self.chk_validate.isChecked()
        o["claim_pattern"] = self.pat_edit.text().strip() or r"^[A-Za-z0-9\-]{4,}$"
        o["barcode_autofill"] = self.chk_barcode.isChecked()
        o["duplicate_check"] = self.chk_dup.isChecked()
        o["excel_log"] = self.chk_excel.isChecked()
        o["activity_log"] = self.chk_log.isChecked()
        o["backup"] = self.chk_backup.isChecked()
        o["backup_folder"] = self.backup_edit.text().strip()
        return o


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Setup wizard + Help (public-friendly)
# ---------------------------------------------------------------------------

class ScanProgressDialog(QtWidgets.QDialog):
    """NAPS2-style scanning box: page counter + Run in Background + Cancel."""
    cancelled = QtCore.pyqtSignal()

    def __init__(self, parent, title, lang="hi"):
        super().__init__(parent)
        self._lang = lang
        self.setWindowTitle(title or "Scanning")
        self.setModal(False)
        self.setMinimumWidth(430)
        lay = QtWidgets.QVBoxLayout(self)
        self.lbl = QtWidgets.QLabel(self._page_text(1))
        lay.addWidget(self.lbl)
        self.bar = QtWidgets.QProgressBar(); self.bar.setRange(0, 0)
        lay.addWidget(self.bar)
        row = QtWidgets.QHBoxLayout(); row.addStretch(1)
        self.btn_bg = QtWidgets.QPushButton("Run in Background" if lang == "en" else "Background me chalao")
        self.btn_bg.clicked.connect(self.hide)
        self.btn_cancel = QtWidgets.QPushButton("Cancel")
        self.btn_cancel.clicked.connect(lambda: self.cancelled.emit())
        row.addWidget(self.btn_bg); row.addWidget(self.btn_cancel)
        lay.addLayout(row)
        self.show()

    def _page_text(self, n):
        return ("Scanning page %d..." % n) if self._lang == "en" else ("Page %d scan ho raha hai..." % n)

    def set_page(self, n):
        self.lbl.setText(self._page_text(n))

    def closeEvent(self, e):
        e.accept()


class SetupWizard(QtWidgets.QWizard):
    def __init__(self, parent, lang="hi"):
        super().__init__(parent)
        self.lang = lang
        self.result_method = "twain"
        self.result_device_name = None
        self.result_wia_id = None
        self.setWindowTitle("Setup — ApneScan")
        self.resize(580, 440)
        self.addPage(self._welcome())
        self.addPage(self._device_page())
        self.addPage(self._finish())

    def _welcome(self):
        p = QtWidgets.QWizardPage()
        p.setTitle("Swagat hai" if self.lang == "hi" else "Welcome")
        lay = QtWidgets.QVBoxLayout(p)
        txt = ("Ye chhota setup aapke scanner ko jodega.\nNext dabao."
               if self.lang == "hi"
               else "This quick setup will connect your scanner.\nClick Next.")
        lbl = QtWidgets.QLabel(txt); lbl.setWordWrap(True); lay.addWidget(lbl)
        return p

    def _device_page(self):
        p = QtWidgets.QWizardPage()
        p.setTitle("Scanner chuno" if self.lang == "hi" else "Choose scanner")
        lay = QtWidgets.QVBoxLayout(p)
        self.rb_twain = QtWidgets.QRadioButton("TWAIN (zyadatar scanners)")
        self.rb_wia = QtWidgets.QRadioButton("WIA (Windows built-in)")
        self.rb_twain.setChecked(True)
        lay.addWidget(self.rb_twain); lay.addWidget(self.rb_wia)
        row = QtWidgets.QHBoxLayout()
        self.dev_lbl = QtWidgets.QLabel("(koi scanner nahi chuna)")
        btn = QtWidgets.QPushButton("Choose scanner"); btn.clicked.connect(self._choose)
        row.addWidget(self.dev_lbl, 1); row.addWidget(btn)
        w = QtWidgets.QWidget(); w.setLayout(row); lay.addWidget(w)
        hint = QtWidgets.QLabel(
            "Scanner list me na dikhe to doosra option (TWAIN/WIA) chuno."
            if self.lang == "hi"
            else "If your scanner isn't listed, try the other option (TWAIN/WIA).")
        hint.setWordWrap(True); lay.addWidget(hint)
        return p

    def _choose(self):
        if self.rb_wia.isChecked():
            self.result_method = "wia"
            try:
                devs = list_wia_sources()
            except Exception as exc:
                QtWidgets.QMessageBox.warning(self, "Error", friendly_error(exc, self.lang)); return
            if not devs:
                QtWidgets.QMessageBox.warning(self, "Error", "Koi WIA scanner nahi mila."); return
            names = [n for _i, n in devs]
            name, ok = QtWidgets.QInputDialog.getItem(self, "Scanner", "Scanner:", names, 0, False)
            if ok and name:
                for _id, n in devs:
                    if n == name:
                        self.result_wia_id = _id; self.result_device_name = n; break
                self.dev_lbl.setText(name)
        else:
            self.result_method = "twain"
            try:
                names = list_sources(int(self.winId()))
            except Exception as exc:
                QtWidgets.QMessageBox.warning(self, "Error", friendly_error(exc, self.lang)); return
            if not names:
                QtWidgets.QMessageBox.warning(self, "Error", "Koi TWAIN scanner nahi mila."); return
            name, ok = QtWidgets.QInputDialog.getItem(self, "Scanner", "Scanner:", names, 0, False)
            if ok and name:
                self.result_device_name = name; self.dev_lbl.setText(name)

    def _finish(self):
        p = QtWidgets.QWizardPage()
        p.setTitle("Ho gaya!" if self.lang == "hi" else "Done!")
        lay = QtWidgets.QVBoxLayout(p)
        lbl = QtWidgets.QLabel(
            "Finish dabao, phir 'Scan' dabao. Bas!"
            if self.lang == "hi" else "Click Finish, then press 'Scan'. That's it!")
        lbl.setWordWrap(True); lay.addWidget(lbl)
        return p


class HelpDialog(QtWidgets.QDialog):
    def __init__(self, parent, lang="hi"):
        super().__init__(parent)
        self.setWindowTitle("Guide / Help"); self.resize(620, 520)
        lay = QtWidgets.QVBoxLayout(self)
        br = QtWidgets.QTextBrowser()
        br.setHtml(HELP_TEXT_HI if lang == "hi" else HELP_TEXT_EN)
        lay.addWidget(br)
        b = QtWidgets.QPushButton("Band karo" if lang == "hi" else "Close")
        b.clicked.connect(self.accept); lay.addWidget(b)



def _make_icon(kind, color="#0f766e"):
    """Draw a small flat icon (28x28) for the toolbar. No external files needed."""
    from PyQt5 import QtGui, QtCore
    pm = QtGui.QPixmap(30, 30); pm.fill(QtCore.Qt.transparent)
    p = QtGui.QPainter(pm); p.setRenderHint(QtGui.QPainter.Antialiasing)
    pen = QtGui.QPen(QtGui.QColor(color)); pen.setWidth(2); pen.setJoinStyle(QtCore.Qt.RoundJoin)
    pen.setCapStyle(QtCore.Qt.RoundCap); p.setPen(pen)
    br = QtGui.QColor(color)
    def doc():
        p.drawRoundedRect(8, 5, 14, 20, 2, 2)
        p.drawLine(11, 11, 19, 11); p.drawLine(11, 15, 19, 15); p.drawLine(11, 19, 16, 19)
    if kind == "scan":
        doc(); p.setBrush(QtGui.QColor("#5eead4")); p.drawRoundedRect(5, 13, 20, 4, 2, 2)
    elif kind == "profiles":
        p.drawRoundedRect(6, 7, 18, 5, 2, 2); p.drawRoundedRect(6, 14, 18, 5, 2, 2); p.drawRoundedRect(6, 21, 12, 5, 2, 2)
    elif kind == "fast":
        p.setBrush(QtGui.QColor("#f59e0b")); p.setPen(QtGui.QColor("#f59e0b"))
        p.drawPolygon(QtGui.QPolygon([QtCore.QPoint(17,4),QtCore.QPoint(9,17),QtCore.QPoint(14,17),QtCore.QPoint(12,26),QtCore.QPoint(21,12),QtCore.QPoint(15,12)]))
    elif kind == "ocr":
        f = p.font(); f.setBold(True); f.setPixelSize(11); p.setFont(f); p.drawText(pm.rect(), QtCore.Qt.AlignCenter, "OCR")
    elif kind == "import":
        doc(); p.drawLine(15, 26, 15, 19); p.drawLine(12, 22, 15, 19); p.drawLine(18, 22, 15, 19)
    elif kind == "savepdf":
        doc(); f=p.font(); f.setBold(True); f.setPixelSize(7); p.setFont(f); p.drawText(QtCore.QRect(8,17,14,8), QtCore.Qt.AlignCenter, "PDF")
    elif kind == "images":
        p.drawRoundedRect(6, 8, 18, 14, 2, 2); p.setBrush(br); p.drawEllipse(10, 12, 3, 3)
        p.drawLine(8, 20, 13, 15); p.drawLine(13, 20, 18, 14)
    elif kind == "print":
        p.drawRoundedRect(8, 5, 14, 7, 1, 1); p.drawRoundedRect(5, 12, 20, 9, 2, 2); p.drawRoundedRect(9, 19, 12, 7, 1, 1)
    elif kind == "rotate":
        p.drawArc(7, 7, 16, 16, 30*16, 280*16); p.drawLine(22, 8, 22, 13); p.drawLine(22, 13, 17, 13)
    elif kind == "up":
        p.drawLine(15, 22, 15, 9); p.drawLine(15, 9, 10, 14); p.drawLine(15, 9, 20, 14)
    elif kind == "down":
        p.drawLine(15, 8, 15, 21); p.drawLine(15, 21, 10, 16); p.drawLine(15, 21, 20, 16)
    elif kind == "delete":
        p.drawLine(9, 9, 21, 21); p.drawLine(21, 9, 9, 21)
    elif kind == "clear":
        p.drawRoundedRect(8, 8, 14, 16, 2, 2); p.drawLine(6, 8, 24, 8); p.drawLine(13, 5, 17, 5)
    elif kind == "language":
        p.drawEllipse(6, 6, 18, 18); p.drawLine(6, 15, 24, 15); p.drawArc(11, 6, 8, 18, 0, 360*16)
    elif kind == "about":
        p.drawEllipse(6, 6, 18, 18); f=p.font(); f.setBold(True); f.setPixelSize(13); p.setFont(f); p.drawText(pm.rect(), QtCore.Qt.AlignCenter, "i")
    else:
        doc()
    p.end()
    return QtGui.QIcon(pm)


class PreviewDialog(QtWidgets.QDialog):
    """NAPS2-style full preview: navigate pages, zoom, rotate, rename, delete,
    save — all from one window."""
    def __init__(self, win, row):
        super().__init__(win)
        self.win = win
        self.row = row
        self.zoom = 1.0
        self.fit = True
        self.setWindowTitle("Preview")
        self.resize(940, 900)
        v = QtWidgets.QVBoxLayout(self)
        tb = QtWidgets.QHBoxLayout()

        def b(txt, fn, tip=""):
            btn = QtWidgets.QToolButton()
            btn.setText(txt); btn.setToolTip(tip)
            btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextOnly)
            btn.clicked.connect(fn)
            btn.setStyleSheet("QToolButton{padding:5px 9px; font-size:13px;}")
            tb.addWidget(btn)
            return btn

        b("\u25c0", self.prev, "Pichhla page (Left arrow)")
        self.lbl_count = QtWidgets.QLabel(); self.lbl_count.setStyleSheet("font-weight:bold;")
        tb.addWidget(self.lbl_count)
        b("\u25b6", self.next, "Agla page (Right arrow)")
        tb.addSpacing(14)
        b("\u2796", self.zoom_out, "Zoom out ( - )")
        b("Fit", self.fit_view, "Poora page dikhao")
        b("\u2795", self.zoom_in, "Zoom in ( + )")
        b("100%", self.actual_size, "Asli size")
        tb.addSpacing(14)
        b("\u21ba", lambda: self.rotate(-90), "Baayein ghumao")
        b("\u21bb", lambda: self.rotate(90), "Dayein ghumao")
        b("\u2712 Rename", self.rename, "Naam badlo (F2)")
        b("\U0001f5d1 Delete", self.delete, "Is page ko hatao")
        b("\U0001f4be Save", self.save_one, "Sirf is page ko PDF me save karo")
        tb.addStretch(1)
        b("\u2715 Close", self.accept, "Band karo (Esc)")
        v.addLayout(tb)

        self.lbl = QtWidgets.QLabel(); self.lbl.setAlignment(QtCore.Qt.AlignCenter)
        self.scr = QtWidgets.QScrollArea(); self.scr.setWidgetResizable(False)
        self.scr.setAlignment(QtCore.Qt.AlignCenter); self.scr.setWidget(self.lbl)
        v.addWidget(self.scr, 1)
        QtCore.QTimer.singleShot(0, self._load)

    def _path(self):
        it = self.win.list.item(self.row)
        return it.data(QtCore.Qt.UserRole) if it else None

    def _load(self):
        n = self.win.list.count()
        if n == 0:
            self.accept(); return
        self.row = max(0, min(self.row, n - 1))
        it = self.win.list.item(self.row)
        title = it.data(TITLE_ROLE) or it.text() or ("Page %d" % (self.row + 1))
        self.lbl_count.setText("  %d / %d  \u2014 %s  " % (self.row + 1, n, title))
        self.win.list.setCurrentRow(self.row)
        pix = QtGui.QPixmap(self._path())
        if pix.isNull():
            self.lbl.clear(); return
        if self.fit:
            area = self.scr.viewport().size()
            pix = pix.scaled(max(50, area.width() - 6), max(50, area.height() - 6),
                             QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        else:
            pix = pix.scaled(int(pix.width() * self.zoom), int(pix.height() * self.zoom),
                             QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.lbl.setPixmap(pix); self.lbl.resize(pix.size())

    def prev(self):
        self.row -= 1; self.fit = True; self._load()

    def next(self):
        self.row += 1; self.fit = True; self._load()

    def zoom_in(self):
        self.fit = False; self.zoom = min(6.0, self.zoom * 1.25); self._load()

    def zoom_out(self):
        self.fit = False; self.zoom = max(0.1, self.zoom / 1.25); self._load()

    def actual_size(self):
        self.fit = False; self.zoom = 1.0; self._load()

    def fit_view(self):
        self.fit = True; self._load()

    def rotate(self, ang):
        p = self._path()
        try:
            with Image.open(p) as im:
                im.rotate(-ang, expand=True).save(p)
            self.win._refresh_item(self.win.list.item(self.row))
            self.win._dirty = True
        except Exception:
            pass
        self.fit = True; self._load()

    def rename(self):
        self.win.list.setCurrentRow(self.row)
        self.win.rename_current_page()
        self._load()

    def delete(self):
        it = self.win.list.item(self.row)
        if it is None:
            return
        self.win.list.clearSelection(); it.setSelected(True)
        self.win.list.setCurrentRow(self.row)
        self.win.delete_page()
        self._load()

    def save_one(self):
        p = self._path()
        if p:
            self.win.save_pdf([p])

    def keyPressEvent(self, e):
        k = e.key()
        if k == QtCore.Qt.Key_Left:
            self.prev()
        elif k == QtCore.Qt.Key_Right:
            self.next()
        elif k in (QtCore.Qt.Key_Plus, QtCore.Qt.Key_Equal):
            self.zoom_in()
        elif k == QtCore.Qt.Key_Minus:
            self.zoom_out()
        elif k == QtCore.Qt.Key_F2:
            self.rename()
        else:
            super().keyPressEvent(e)


class CameraDialog(QtWidgets.QDialog):
    """Webcam/USB-camera se document capture — scanner na ho to bhi PDF banao.
    Capture ki gayi har photo apne aap saaf (shadow hatana/seedha) hoti hai."""

    def __init__(self, parent, tmpdir):
        super().__init__(parent)
        self.tmpdir = tmpdir
        self.captured = []          # saved temp image paths
        self.setWindowTitle("📷 Camera se scan")
        self.resize(720, 560)
        v = QtWidgets.QVBoxLayout(self)
        self.viewf = QCameraViewfinder()
        v.addWidget(self.viewf, 1)
        row = QtWidgets.QHBoxLayout()
        self.cmb = QtWidgets.QComboBox()
        self._cams = QCameraInfo.availableCameras()
        for c in self._cams:
            self.cmb.addItem(c.description())
        row.addWidget(self.cmb, 1)
        self.lbl = QtWidgets.QLabel("0 captured")
        row.addWidget(self.lbl)
        v.addLayout(row)
        brow = QtWidgets.QHBoxLayout()
        self.b_shot = QtWidgets.QPushButton("📸 Capture (Space)")
        self.b_shot.setObjectName("primary"); self.b_shot.setMinimumHeight(40)
        self.b_shot.clicked.connect(self._capture)
        b_done = QtWidgets.QPushButton("✔ Done — pages add karo")
        b_done.clicked.connect(self.accept)
        b_cancel = QtWidgets.QPushButton("Cancel")
        b_cancel.clicked.connect(self.reject)
        brow.addWidget(self.b_shot, 2); brow.addWidget(b_done, 1); brow.addWidget(b_cancel)
        v.addLayout(brow)
        QtWidgets.QShortcut(QtGui.QKeySequence("Space"), self, self._capture)
        self.camera = None
        self.cap = None
        self.cmb.currentIndexChanged.connect(self._start_camera)
        if self._cams:
            self._start_camera(0)
        else:
            self.b_shot.setEnabled(False)
            self.viewf.setStyleSheet("background:#111;")

    def _start_camera(self, idx):
        try:
            if self.camera:
                self.camera.stop()
            self.camera = QCamera(self._cams[idx])
            self.camera.setViewfinder(self.viewf)
            self.cap = QCameraImageCapture(self.camera)
            self.cap.imageCaptured.connect(self._on_captured)
            self.camera.start()
        except Exception:
            self.b_shot.setEnabled(False)

    def _capture(self):
        if not self.cap:
            return
        try:
            self.cap.capture()          # imageCaptured signal me frame aayega
        except Exception:
            pass

    def _on_captured(self, _id, qimg):
        try:
            fd, png = tempfile.mkstemp(suffix=".jpg", dir=self.tmpdir)
            os.close(fd)
            qimg.save(png, "JPG", 92)
            # scan-jaisa saaf: shadow hatana + seedha (clean_photo module-level hai)
            try:
                with Image.open(png) as im:
                    clean_photo(im).save(png, "JPEG", quality=90)
            except Exception:
                pass
            self.captured.append(png)
            self.lbl.setText("%d captured" % len(self.captured))
        except Exception:
            pass

    def closeEvent(self, e):
        try:
            if self.camera:
                self.camera.stop()
        except Exception:
            pass
        super().closeEvent(e)


class SparkBars(QtWidgets.QWidget):
    """Chhota bar-chart (bina kisi library ke) — dashboard ke liye."""

    def __init__(self, values, labels=None, color="#0f766e"):
        super().__init__()
        self.v = [max(0, int(x or 0)) for x in (values or [0])]
        self.labels = labels
        self.c = QtGui.QColor(color)
        self.setMinimumHeight(96)

    def set_values(self, values, labels=None):
        self.v = [max(0, int(x or 0)) for x in (values or [0])]
        if labels is not None:
            self.labels = labels
        self.update()

    def paintEvent(self, _e):
        p = QtGui.QPainter(self)
        try:
            W, H = self.width(), self.height() - 18
            mx = max(self.v) or 1
            n = max(1, len(self.v))
            gap = W // n
            bw = max(6, int(gap * 0.6))
            p.setPen(QtCore.Qt.NoPen)
            p.setBrush(self.c)
            for i, val in enumerate(self.v):
                h = int((H - 14) * val / mx)
                x = i * gap + (gap - bw) // 2
                p.drawRect(x, H - h, bw, h)
            p.setPen(QtGui.QColor("#64748b"))
            f = p.font(); f.setPointSize(7); p.setFont(f)
            for i, val in enumerate(self.v):
                p.drawText(i * gap, 0, gap, 12, QtCore.Qt.AlignCenter, str(val))
            if self.labels:
                for i, lb in enumerate(self.labels):
                    p.drawText(i * gap, H + 2, gap, 14, QtCore.Qt.AlignCenter, str(lb)[:3])
        finally:
            p.end()


class LibraryModel(QtWidgets.QFileSystemModel):
    """'Meri Files' panel ka model — AAJ banayi files hari dikhti hain."""

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if role == QtCore.Qt.ForegroundRole and index.column() == 0:
            try:
                fi = self.fileInfo(index)
                if fi.isFile() and fi.lastModified().date() == QtCore.QDate.currentDate():
                    return QtGui.QBrush(QtGui.QColor("#16a34a"))
            except Exception:
                pass
        return super().data(index, role)


class FilesTree(QtWidgets.QTreeView):
    """Folder tree jo pages ka DROP le leta hai — page ko kheench kar folder
    par chhodo = seedha wahan save."""

    def __init__(self, on_drop):
        super().__init__()
        self._on_drop = on_drop
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        e.acceptProposedAction()

    def dragMoveEvent(self, e):
        e.acceptProposedAction()

    def dropEvent(self, e):
        try:
            self._on_drop(self.indexAt(e.pos()))
            e.acceptProposedAction()
        except Exception:
            e.ignore()


class ScannerWindow(QtWidgets.QMainWindow):
    THUMB_W = 150
    THUMB_H = 200

    def __init__(self, auto_scan_profile=None):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        try:
            self.setWindowIcon(app_icon())
        except Exception:
            pass
        self.resize(1220, 820)
        self._tmpdir = tempfile.mkdtemp(prefix="nobledoc_")
        self._config = load_config()
        self._profiles = self._config.get("profiles", [])
        self._opts = dict(DEFAULT_OPTIONS)
        self._opts.update(self._config.get("options", {}))
        self._recent = self._config.get("recent", [])
        self._used_claims = set(self._config.get("used_claims", []))
        self._barcode_tried = False
        self._dirty = False
        self._undo_stack = []
        self._lang = self._opts.get("language", "hi")

        self._build_menu()
        self._build_ui()
        self._apply_style()
        self._refresh_profile_combo()
        self._load_profile_to_panel(self._selected_profile())
        self._refresh_method_label()
        self._refresh_recent_menu()
        self._update_status()

        QtCore.QTimer.singleShot(400, self.check_connection)
        self._conn_timer = QtCore.QTimer(self)
        self._conn_timer.setInterval(30000)
        self._conn_timer.timeout.connect(self.check_connection)
        self._conn_timer.start()

        if not self._config.get("setup_done"):
            QtCore.QTimer.singleShot(600, self._run_wizard)

        self._apply_simple_mode()

        if auto_scan_profile:
            i = self.cmb_profile.findText(auto_scan_profile)
            if i >= 0:
                self.cmb_profile.setCurrentIndex(i)
            QtCore.QTimer.singleShot(800, self.do_scan)

    # ---- config helpers ----
    def _save_opts(self):
        self._config["options"] = self._opts; save_config(self._config)

    def _save_profiles(self):
        self._config["profiles"] = self._profiles; save_config(self._config)

    def _add_recent(self, path):
        self._recent = [path] + [r for r in self._recent if r != path]
        self._recent = self._recent[:12]
        self._config["recent"] = self._recent; save_config(self._config)
        self._refresh_recent_menu()

    # ---- menu + shortcuts ----
    def _ma(self, menu, text, slot, tip="", sc=None):
        # add a menu action with a "?"-prefixed label + a Hindi/English hover tip
        a = menu.addAction(("\u2753 " + text) if tip else text, slot)
        if tip:
            a.setToolTip(tip); a.setStatusTip(tip)
        return a

    def _build_menu(self):
        mb = self.menuBar()
        mf = mb.addMenu(tr("menu_file", self._lang)); mf.setToolTipsVisible(True)
        self._ma(mf, tr("save_all", self._lang), self.save_pdf_all,
                 "हिन्दी: Sabhi pages ki ek PDF banao.\nEnglish: Save all pages as one PDF.", "Ctrl+S")
        self._ma(mf, tr("save_sel", self._lang), self.save_pdf_selected,
                 "हिन्दी: Sirf chune hue (Ctrl/Shift se) pages ki PDF.\nEnglish: PDF of only the selected pages.")
        self._ma(mf, "Save PDF (password)…", self.save_pdf_password,
                 "हिन्दी: Password se surakshit PDF (khole ke liye password lagega).\nEnglish: Password-protected PDF.")
        self._ma(mf, "Save Images…", self.save_images,
                 "हिन्दी: Pages ko JPG/PNG image me save karo (PDF ke bajay).\nEnglish: Save pages as JPG/PNG images.", "Ctrl+Shift+S")
        self._ma(mf, "Export OCR text…", self.export_ocr_text,
                 "हिन्दी: Page ke text ko padh kar .txt me nikaalo (OCR).\nEnglish: Extract page text to .txt via OCR.")
        pmenu = mf.addMenu("Print")
        pmenu.setToolTipsVisible(True)
        self._ma(pmenu, "All print (sabhi pages)", self.print_all,
                 "हिन्दी: Saare pages print karo.\nEnglish: Print all pages.")
        self._ma(pmenu, "Selected print (chune hue)", self.print_selected,
                 "हिन्दी: Sirf chune hue (Ctrl/Shift se) pages print karo.\nEnglish: Print only the selected pages.")
        self._ma(pmenu, "ID print (2 ID ek page par)", self.print_ids,
                 "हिन्दी: ID cards ko ek A4 page par 2-2 karke print karo (kaagaz bachega).\nEnglish: Print ID cards two-per-A4-sheet to save paper.")
        self._ma(pmenu, "ID print - sirf selected", self.print_ids_selected,
                 "हिन्दी: Sirf chune hue IDs ko 2-per-page print karo.\nEnglish: Print only selected IDs, two per sheet.")
        shmenu = mf.addMenu("Share / Bhejo")
        shmenu.setToolTipsVisible(True)
        self._ma(shmenu, "WhatsApp par bhejo…", self.share_whatsapp,
                 "हिन्दी: Aakhri save ki hui PDF WhatsApp par bhejo (file copy ho jati hai, chat me Ctrl+V ya drag se attach).\nEnglish: Send the last saved PDF via WhatsApp (file is copied; paste or drag into the chat).")
        self._ma(shmenu, "Email se bhejo…", self.share_email,
                 "हिन्दी: PDF ko Email se bhejo (Outlook ho to attachment ke saath draft khulta hai).\nEnglish: Send the PDF by Email (opens an Outlook draft with the attachment if available).")
        self.recent_menu = mf.addMenu("Recent PDFs")
        self.recent_menu.setToolTipsVisible(True)
        mf.addSeparator()
        self._ma(mf, "Exit", self.close, "हिन्दी: App band karo. (Bina save kiye pages hon to chetavni aayegi.)\nEnglish: Close the app. (Warns if there are unsaved pages.)")

        me = mb.addMenu(tr("menu_edit", self._lang)); me.setToolTipsVisible(True)
        self._ma(me, "Rotate left", self.rotate_left, "हिन्दी: Selected page ko baayein ghumao.\nEnglish: Rotate the selected page left.")
        self._ma(me, "Rotate right", self.rotate_right, "हिन्दी: Selected page ko daayein ghumao.\nEnglish: Rotate the selected page right.")
        self._ma(me, "Brightness +", lambda: self._enhance_current(1.12, 1.0), "हिन्दी: Page ko halka (bright) karo.\nEnglish: Make the page brighter.")
        self._ma(me, "Brightness -", lambda: self._enhance_current(0.9, 1.0), "हिन्दी: Page ko gehra (dim) karo.\nEnglish: Make the page darker.")
        self._ma(me, "Contrast +", lambda: self._enhance_current(1.0, 1.15), "हिन्दी: Text aur saaf/gehra dikhe.\nEnglish: Increase contrast (sharper text).")
        self._ma(me, "Contrast -", lambda: self._enhance_current(1.0, 0.88), "हिन्दी: Contrast kam karo.\nEnglish: Decrease contrast.")
        self._ma(me, "Auto-crop page", self.autocrop_current, "हिन्दी: Page ke aas-paas ki khaali border kaato.\nEnglish: Trim the empty border around the page.")
        self._ma(me, "Page ka text copy karo", self.copy_page_text, "हिन्दी: Is page ka poora text padh kar copy kar lo (kahin bhi paste karo).\nEnglish: OCR this page's text to the clipboard.")
        self._ma(me, "Page translate karo (Hindi ↔ English)…", self.translate_page, "हिन्दी: Page ka text padh kar Hindi/English me translate karo (internet chahiye).\nEnglish: Translate the page's text between Hindi and English (needs internet).")
        self._ma(me, "Undo delete", self.undo_delete, "हिन्दी: Galti se delete hua page wapas laao.\nEnglish: Restore a deleted page.", "Ctrl+Z")
        me.addSeparator()
        self._ma(me, "Delete page", self.delete_page, "हिन्दी: Selected page hatao.\nEnglish: Delete the selected page.", "Delete")
        self._ma(me, "Clear all", self.clear_all, "हिन्दी: Saare pages hatao (khaali karo).\nEnglish: Remove all pages.")

        mt = mb.addMenu(tr("menu_tools", self._lang)); mt.setToolTipsVisible(True)
        self._ma(mt, "Bundle → alag-alag PDFs (blank separator)…", self.bundle_split_save, "हिन्दी: Poora bundle ek saath feeder me daalo, documents ke beech KHAALI page rakho — ye khud alag-alag PDF bana dega. (Settings me 'blank hatao' OFF rakhein.)\nEnglish: Scan a whole bundle with a blank page between documents; this splits and saves separate PDFs automatically.")
        self._ma(mt, "Book page → 2 pages me kaato", self.split_book_page, "हिन्दी: Khuli kitab ke scan ko beech se kaat kar do alag pages bana do.\nEnglish: Split an open-book scan into left and right pages.")
        self._ma(mt, "Business cards → contacts…", self.business_cards, "हिन्दी: Visiting cards ke scan se naam/phone/email padh kar contact files (.vcf) + Excel banao.\nEnglish: Read visiting cards into contact (.vcf) files and an Excel sheet.")
        self._ma(mt, "Purani photo sudharo (restore)", self.restore_photo_current, "हिन्दी: Feeki/dhundhli purani photo ka rang-roop sudharo.\nEnglish: Restore faded/dull old photos.")
        self._ma(mt, "Scan History…", self.show_history, "हिन्दी: Ab tak ki saari saved PDFs — nayi se purani, filter ke saath.\nEnglish: All saved PDFs, newest first, with quick filter.")
        self._ma(mt, "📊 Stats Dashboard…", self.show_stats_dashboard, "हिन्दी: Aapki + duniya bhar ki poori statistics — graphs ke saath. (Sidebar ke stats-box par click karke bhi khulta hai.)\nEnglish: Full personal + worldwide statistics with charts.")
        self._ma(mt, "📷 Camera se scan (webcam)…", self.scan_from_camera, "हिन्दी: Scanner na ho to bhi — webcam/USB camera se document capture karke PDF banao (photo apne aap saaf hoti hai).\nEnglish: No scanner? Capture documents with a webcam/USB camera (auto-cleaned).")
        self._ma(mt, "Phone-photo se PDF (photo import)…", self.import_photos, "हिन्दी: Phone se kheenchi document-photos ko saaf karke pages banao (shadow hatana, seedha karna) — phir PDF save karo.\nEnglish: Clean up phone photos of documents (remove shadows, straighten) and add them as pages.")
        self._ma(mt, "ID cards alag karo (is page se)…", self.split_id_cards, "हिन्दी: Ek page par 2-3 ID cards scan kiye hain? Ye unhe alag-alag pages me kaat dega.\nEnglish: Scanned 2-3 ID cards on one page? This splits them into separate pages.")
        self._ma(mt, "Search past PDFs…", self.search_pdfs, "हिन्दी: Purani save ki hui PDF dhoondo (claim/naam/tag se, ya PDF ke andar ke text se).\nEnglish: Search your saved PDFs by name, tag or text content.", "Ctrl+F")
        self._ma(mt, "Search index banao/refresh…", self.build_search_index, "हिन्दी: Saari PDFs ka text ek baar padh kar index bana lo — phir andar-ke-text wali search TURANT hogi.\nEnglish: Build a one-time text index so in-PDF search becomes instant.")
        self._ma(mt, "Tag lagao (kisi PDF par)…", self.tag_pdf, "हिन्दी: PDF par apne tags lagao (jaise Aadhaar, School, Bijli-bill) — baad me tag se turant dhoondo.\nEnglish: Put your own tags on a PDF for quick finding later.")
        self._ma(mt, "Tag se dhoondo…", self.search_by_tag, "हिन्दी: Lagaye hue tag se files ki list dekho aur kholo.\nEnglish: List and open files by tag.")
        self._ma(mt, "Merge PDFs…", self.merge_pdfs, "हिन्दी: Kai PDF ko jodkar ek PDF banao.\nEnglish: Merge several PDFs into one.")
        self._ma(mt, "Split into multiple PDFs…", self.split_pdfs, "हिन्दी: Ek scan ko kai alag PDF me baanto.\nEnglish: Split into multiple PDFs.")
        self._ma(mt, "PDF chhota karo (compress)…", self.compress_pdf_tool, "हिन्दी: Abhi ke pages ya koi purani PDF ko 200KB/500KB/1MB/2MB tak chhota karo (portal upload ke liye).\nEnglish: Shrink current pages or any PDF to a 200KB/500KB/1MB/2MB target for portal uploads.")
        pdft = mt.addMenu("PDF Tools")
        pdft.setToolTipsVisible(True)
        self._ma(pdft, "PDF page editor (kram/ghumao/hatao)…", self.pdf_page_editor,
                 "हिन्दी: Kisi bhi PDF ke pages ka kram badlo, ghumao ya hatao — bina quality kharaab kiye (lossless).\nEnglish: Reorder, rotate or remove pages of any PDF, losslessly.")
        self._ma(pdft, "Sign/Stamp lagao (is page par)…", self.place_sign,
                 "हिन्दी: Apne signature/mohar ki image current page par lagao (safed background apne aap transparent).\nEnglish: Place your signature/stamp image on the current page (white background auto-transparent).")
        self._ma(pdft, "Page numbers lagao (sab pages par)…", self.add_page_numbers,
                 "हिन्दी: Har page par 'Page 1/5' aur chaaho to upar apna header chhapo.\nEnglish: Print 'Page 1/5' on every page, with an optional header.")
        self._ma(pdft, "Kisi PDF par watermark…", self.watermark_pdf_tool,
                 "हिन्दी: Kisi purani PDF par apna text-watermark/stamp chhapo.\nEnglish: Stamp a text watermark onto any existing PDF.")
        self._ma(pdft, "PDF ka password hatao…", self.remove_pdf_password,
                 "हिन्दी: Password pata ho to us PDF ki bina-password copy banao.\nEnglish: If you know the password, make a password-free copy of the PDF.")
        self._ma(pdft, "PDF → Word (.docx)…", self.pdf_to_word,
                 "हिन्दी: PDF/pages ka text OCR karke Word file banao (edit karne layak).\nEnglish: OCR the text into an editable Word document.")
        self._ma(pdft, "PDF → Excel (.xlsx)…", self.pdf_to_excel,
                 "हिन्दी: Bill/table wale pages ko OCR karke Excel me nikaalo (best-effort).\nEnglish: Extract bill/table pages into an Excel sheet (best-effort).")
        self._ma(pdft, "PDF → JPG images…", self.pdf_to_jpgs,
                 "हिन्दी: Kisi PDF ke har page ko alag JPG image me nikaalo.\nEnglish: Export each page of a PDF as a separate JPG image.")
        self._ma(pdft, "Folder ki images → ek PDF…", self.folder_to_pdf,
                 "हिन्दी: Ek folder ki saari images (naam ke kram me) ek PDF me jodo.\nEnglish: Combine all images in a folder (name order) into one PDF.")
        self._ma(pdft, "Archival PDF (300dpi + metadata)…", self.save_archival_pdf,
                 "हिन्दी: High-quality PDF (300dpi) poore metadata (title/date) ke saath — lambe samay tak sambhalne ke liye.\nEnglish: High-quality 300dpi PDF with full metadata for long-term archiving.")
        self._ma(mt, "Monthly report…", self.monthly_report, "हिन्दी: Mahine ka scan/claim report banao.\nEnglish: Generate a monthly report.")
        self._ma(mt, "Create desktop shortcut…", self.create_shortcut, "हिन्दी: Desktop par ek-click scan ka shortcut banao.\nEnglish: Make a one-click desktop scan shortcut.")
        self._ma(mt, "Auto-name pages (document ka naam)", self.auto_name_pages, "हिन्दी: Har page ko padh kar uska naam (jaise DISCHARGE SUMMARY, RECEIPT) thumbnail ke neeche likhe. 'Page 1,2' ke bajay asli naam.\nEnglish: Read each page and label it with its document title instead of 'Page 1,2'.")

        ms = mb.addMenu(tr("menu_settings", self._lang)); ms.setToolTipsVisible(True)
        self._ma(ms, tr("options", self._lang), self.open_options, "हिन्दी: App ki saari settings (auto-save, blank hatao, backup, waghera).\nEnglish: All app settings.")
        self._ma(ms, tr("profiles", self._lang), self.open_profiles, "हिन्दी: Scan profiles banao/badlo (device, dpi, colour, duplex).\nEnglish: Create/edit scan profiles.")
        self._ma(ms, tr("scan_method", self._lang) + "…", self.choose_scan_method, "हिन्दी: Scan ka tareeka: escl (network duplex), twain (USB duplex), ya wia.\nEnglish: Scan method: escl (network duplex), twain (USB), or wia.")
        self._ma(ms, tr("language", self._lang) + "…", self.choose_language, "हिन्दी: App ki bhasha badlo (Hindi/English).\nEnglish: Change the app language.")
        self._ma(ms, "Stats server URL…", self.set_stats_url, "हिन्दी: Worldwide stats ke liye Google Apps Script ka URL daalein (kitne scan hue, kitne online).\nEnglish: Set the stats server URL (worldwide scan counts + online users).")
        self._ma(ms, "Scanner khud dhoondo (network)…", self.find_scanners, "हिन्दी: IP yaad rakhne ki zaroorat nahi — network par scanner khud dhoondh kar set karo.\nEnglish: Auto-discover eSCL scanners on the network and set the IP.")
        self._ma(ms, "Scanner IP…", self.set_scanner_ip, "हिन्दी: Network scanner ka IP set karo (jaise 192.168.1.8).\nEnglish: Set the network scanner IP.")
        self._ma(ms, "Keyboard Shortcuts…", self.show_shortcuts, "हिन्दी: Keyboard ke shortcuts ki list dekho.\nEnglish: View keyboard shortcuts.")
        self.act_simple = self._ma(ms, tr("simple_on", self._lang), self.toggle_simple_mode, "हिन्दी: Simple mode: sirf zaroori buttons dikhein (naye users ke liye aasan).\nEnglish: Simple mode: show only the essential buttons.")
        self.act_simple.setCheckable(True)
        self.act_simple.setChecked(bool(self._opts.get("simple_mode")))
        self._ma(ms, self.L("Left sidebar dikhao/chhupao", "Show/hide left sidebar"), self.toggle_left_panel, "हिन्दी: Baayin taraf ka scan-settings panel on/off (zyada jagah ke liye).\nEnglish: Show/hide the left scan-settings sidebar for more space.", "F9")
        self.act_files_panel = self._ma(ms, self.L("Right sidebar (Meri Files) dikhao/chhupao", "Show/hide right sidebar (My Files)"), self.toggle_files_panel, "हिन्दी: Daayin taraf ka folders-wala panel on/off karo.\nEnglish: Show/hide the right-side files panel.", "F10")
        self._ma(ms, "Sidebar stats chuno…", self.choose_sidebar_stats, "हिन्दी: Sidebar ke stats-box me kaun-kaun si ginti dikhe — aap khud chuno (worldwide + personal).\nEnglish: Choose which stats appear in the sidebar box.")
        self._ma(ms, "🎨 UI customize karo…", self.customize_ui, "हिन्दी: App ka look apne hisaab se: dashboard, floating Scan button, status-patti, sidebar graph, Dark Pro theme — jo chaho on/off karo.\nEnglish: Customize the UI: dashboard, floating Scan button, status bar, sidebar graph, Dark Pro theme.")
        self.act_touch = self._ma(ms, "Touch / bade-button mode", self.toggle_touch_mode, "हिन्दी: Buttons/likhai badi ho jayegi — touch screen ya buzurgon ke liye aasan.\nEnglish: Bigger buttons and text for touch screens or elderly users.")
        self.act_touch.setCheckable(True)
        self.act_touch.setChecked(bool(self._opts.get("touch_mode")))
        ms.addSeparator()
        self._ma(ms, "Settings export karo…", self.export_settings, "हिन्दी: Saari settings ek file me — naye PC par le jaane ke liye.\nEnglish: Export all settings to a file for another PC.")
        self._ma(ms, "Settings import karo…", self.import_settings, "हिन्दी: Export ki hui settings file se sab wapas le aao.\nEnglish: Import settings from an exported file.")

        mh = mb.addMenu(tr("menu_help", self._lang)); mh.setToolTipsVisible(True)
        self._ma(mh, tr("help_guide", self._lang), self.show_help, "हिन्दी: App istemal karne ki guide.\nEnglish: How-to guide.")
        self._ma(mh, "Setup wizard", self._run_wizard, "हिन्दी: Pehli baar wala setup dobara chalao.\nEnglish: Re-run the first-time setup.")
        self._ma(mh, tr("whatsnew", self._lang), self.show_whatsnew, "हिन्दी: Naye badlav/features.\nEnglish: What's new.")
        self._ma(mh, "Test / Diagnostics", self.run_diagnostics, "हिन्दी: Scanner/app ki jaankari + error report (share karne ke liye).\nEnglish: Scanner/app info + error report.")
        self._ma(mh, "Duplex Test (both-side)", self.run_duplex_test, "हिन्दी: Jaancho ki dono taraf (duplex) scan ho raha hai ya nahi.\nEnglish: Test whether both-side (duplex) scanning works.")
        self._ma(mh, "eSCL Test (network scan jaanch)", self.run_escl_test, "हिन्दी: Network scan (eSCL) ko step-by-step jaanch kar asli problem batata hai (connect / status / job).\nEnglish: Step-by-step eSCL network-scan test that shows the real problem.")
        self._ma(mh, "Update check karo…", lambda: self.check_updates(False), "हिन्दी: Naya version aaya ho to app use khud download karke install kar legi.\nEnglish: If a newer version exists the app downloads and installs it itself.")
        self._ma(mh, "Error report dekho…", self.open_crash_report, "हिन्दी: Agar app kabhi crash hui ho to uski report kholo (feedback me bhejne ke liye).\nEnglish: Open the saved crash report, if any.")
        self._ma(mh, tr("feedback", self._lang), self.send_feedback, "हिन्दी: Sujhav/shikayat bhejo.\nEnglish: Send feedback.")
        self._ma(mh, tr("about", self._lang), self.show_about, "हिन्दी: App ke baare me.\nEnglish: About this app.")

        # AUTOMATIC: har menu/submenu ke har option par "?" + Hindi/English hover
        # PAKKA karo — aage koi naya option seedhe addAction se bhi joda jaye to
        # bhi ye use chhoot-ne nahi dega (chhupa hua safety-jaal).
        self._finalize_menu_help()
        self._build_shortcuts()

    def _finalize_menu_help(self):
        """Poore menu-tree ko scan karke jis bhi option par abhi tak '?' /
        tooltip nahi hai, uspar khud laga do. Ye function baar-baar chalana
        safe hai (dobara prefix nahi lagta)."""
        MARK = "❓ "

        def walk(menu):
            try:
                menu.setToolTipsVisible(True)
            except Exception:
                pass
            for act in menu.actions():
                if act.isSeparator():
                    continue
                sub = act.menu()
                if sub is not None:
                    walk(sub)
                    continue
                t = act.text()
                if not t:
                    continue
                clean = t[len(MARK):] if t.startswith(MARK) else t
                if not t.startswith(MARK):
                    act.setText(MARK + t)
                # tooltip na ho (ya sirf label ki copy ho) to ek default bilingual
                # explanation laga do
                if not act.toolTip() or act.toolTip() in (t, clean):
                    label = clean.replace("&", "").replace("…", "").strip()
                    act.setToolTip("हिन्दी: %s — is option ko chalata hai.\n"
                                   "English: Runs the \"%s\" action." % (label, label))
        try:
            for act in self.menuBar().actions():
                if act.menu() is not None:
                    walk(act.menu())
        except Exception:
            pass

    def _refresh_recent_menu(self):
        self.recent_menu.clear()
        if not self._recent:
            a = self.recent_menu.addAction("(koi nahi)"); a.setEnabled(False); return
        for path in self._recent:
            self.recent_menu.addAction(os.path.basename(path),
                                       lambda checked=False, p=path: self._open_path(p))

    def _open_path(self, path):
        try:
            os.startfile(path)
        except Exception as exc:
            self._warn("File nahi khuli:\n%s" % exc)

    # ---- profiles / options ----
    def _refresh_profile_combo(self):
        self.cmb_profile.blockSignals(True); self.cmb_profile.clear()
        for p in self._profiles:
            self.cmb_profile.addItem(p.get("name", "Profile"))
        sel = self._config.get("selected_profile")
        if sel:
            i = self.cmb_profile.findText(sel)
            if i >= 0:
                self.cmb_profile.setCurrentIndex(i)
        self.cmb_profile.blockSignals(False)

    def _selected_profile(self):
        name = self.cmb_profile.currentText()
        for p in self._profiles:
            if p.get("name") == name:
                return p
        return None

    def _on_fast_toggled(self, on):
        self._opts["fast_mode"] = bool(on); self._save_opts()
        if on:
            # reflect fast settings in the panel so the user sees them
            self.cmb_dpi.setCurrentText("200 dpi")
            self.cmb_depth.setCurrentText("Black & White")
        self.status.showMessage("Fast mode ON (200 dpi, B&W)" if on else "Fast mode OFF", 3000)

    def _on_profile_changed(self, name):
        self._config["selected_profile"] = name; save_config(self._config)
        self._load_profile_to_panel(self._selected_profile())


    def open_profiles(self):
        dlg = ProfileManagerDialog(self, self._profiles)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self._profiles = dlg.profiles; self._save_profiles(); self._refresh_profile_combo()

    def open_options(self):
        dlg = OptionsDialog(self, self._opts)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self._opts = dlg.get_opts(); self._save_opts(); self._update_status()
            self._apply_style(); self._refresh_page_numbers()
            self._refresh_files_root()   # save-folder badla ho to panel bhi

    def _refresh_page_numbers(self):
        show = self._opts.get("show_page_numbers", True)
        for i in range(self.list.count()):
            it = self.list.item(i)
            title = it.data(TITLE_ROLE)
            if title:
                it.setText(title)
            else:
                it.setText(("Page %d" % (i + 1)) if show else "")

    def _vsep(self):
        f = QtWidgets.QFrame(); f.setObjectName("hr"); f.setFrameShape(QtWidgets.QFrame.VLine); return f

    def _panel_scan_params(self, prof):
        try:
            dpi = int(self.cmb_dpi.currentText().split()[0])
        except Exception:
            dpi = int(prof.get("dpi", 200))
        d = self.cmb_depth.currentText()
        color = "color" if d.startswith("24") else ("gray" if d.startswith("Gray") else "bw")
        duplex = "Both" in self.cmb_sides.currentText()
        return dpi, color, duplex

    def _refresh_method_label(self):
        if not hasattr(self, "method_lbl"):
            return
        m = self._opts.get("scanner_method", "twain")
        color = "#0f766e" if m in ("wia", "naps2", "escl") else "#b45309"
        self.method_lbl.setText('<b>Connected via:</b> <span style="color:%s">%s</span>'
                                % (color, m.upper()))

    def _load_profile_to_panel(self, prof):
        if prof is None or not hasattr(self, "cmb_dpi"):
            return
        self.dev_lbl.setText(prof.get("source_name") or "(koi device nahi)")
        self.cmb_dpi.setCurrentText(str(prof.get("dpi", 200)) + " dpi")
        c = prof.get("color", "gray")
        self.cmb_depth.setCurrentText("24-bit Colour" if c == "color"
                                      else ("Grayscale" if c == "gray" else "Black & White"))
        self.cmb_source.setCurrentText("Feeder (ADF)")
        self.cmb_sides.setCurrentText("Both side (dono taraf)" if prof.get("duplex") else "Single side (ek taraf)")
        # page size (per-profile)
        if hasattr(self, "cmb_pagesize"):
            _ps = (prof.get("page_size") or "auto").lower()
            for k in range(self.cmb_pagesize.count()):
                if self.cmb_pagesize.itemText(k).lower().startswith(_ps[:4]):
                    self.cmb_pagesize.setCurrentIndex(k); break

    def _quick_new_profile(self):
        dlg = EditProfileDialog(self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self._profiles.append(dlg.get_profile()); self._save_profiles(); self._refresh_profile_combo()
            self.cmb_profile.setCurrentIndex(self.cmb_profile.count() - 1)

    def _quick_edit_profile(self):
        name = self.cmb_profile.currentText()
        idx = next((k for k, p in enumerate(self._profiles) if p.get("name") == name), -1)
        if idx < 0:
            self._quick_new_profile(); return
        dlg = EditProfileDialog(self, self._profiles[idx])
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            self._profiles[idx] = dlg.get_profile(); self._save_profiles(); self._refresh_profile_combo()
            i = self.cmb_profile.findText(self._profiles[idx].get("name"))
            if i >= 0:
                self.cmb_profile.setCurrentIndex(i)
            self._load_profile_to_panel(self._selected_profile())

    def _open_preview_dialog(self, item):
        if item is None:
            return
        row = self.list.row(item)
        PreviewDialog(self, row).exec_()

    def auto_name_pages(self):
        if not tesseract_available():
            self._warn(
                "Document ka naam padhne ke liye 'Tesseract OCR' program install karna zaroori hai.\n\n"
                "Download (free): https://github.com/UB-Mannheim/tesseract/wiki\n"
                "Install karke ApneScan dobara kholein. Phir Tools \u2192 Auto-name pages chalega."); return
        items = [(i, self.list.item(i).data(QtCore.Qt.UserRole)) for i in range(self.list.count())]
        if not items:
            self._warn("Pehle koi page scan/import karein."); return
        self._named_count = 0
        self._named_total = len(items)
        self.status.showMessage("Document ke naam padhe ja rahe hain... (thoda ruko)", 0)
        self._namer = NameWorker(items, learned=self._learned_names())
        self._namer.named.connect(self._apply_page_title)
        self._namer.finished_all.connect(self._on_naming_done)
        self._namer.start()

    def _learned_names(self):
        out = []
        for e in (self._config.get("learned_names", []) or []):
            try:
                out.append((set(e.get("words", [])), e.get("name", "")))
            except Exception:
                pass
        return out

    def _apply_page_title(self, row, title):
        if 0 <= row < self.list.count():
            it = self.list.item(row)
            key = name_key(title)
            # if this same document was saved before, reuse THAT saved name
            remembered = (self._config.get("doc_names", {}) or {}).get(key)
            label = remembered or underscore_name(title)
            it.setData(TITLE_ROLE, label)
            it.setData(NAMEKEY_ROLE, key)
            it.setText(label)
            self._named_count = getattr(self, "_named_count", 0) + 1
            self._pstats_bump(ocr_named=1)

    def _on_naming_done(self):
        got = getattr(self, "_named_count", 0)
        total = getattr(self, "_named_total", 0)
        if got == 0:
            self.status.clearMessage()
            self._warn(
                "OCR chala, par kisi bhi page ka naam nahi mil paaya.\n\n"
                "Iske 2 kaaran ho sakte hain:\n"
                "1) Tesseract theek se nahi mila \u2014 Help \u2192 Test/Diagnostics kholein aur "
                "'Tesseract version' line dekhein (WORKS likha hai ya NOT FOUND).\n"
                "2) Page ke UPAR saaf chhapa hua heading nahi hai (jaise haath se likhe page) \u2014 "
                "aise pages ka naam nahi aata, wo normal hai.")
        else:
            self.status.showMessage("%d / %d pages ke naam mil gaye." % (got, total), 6000)

    def _build_shortcuts(self):
        methods = {
            "scan": self.do_scan,
            "import": self.import_images,
            "rename": self.rename_current_page,
            "save_all": self.save_pdf_all,
            "save_sel": self.save_pdf_selected,
            "save_pw": self.save_pdf_password,
            "save_img": self.save_images,
            "ocr_text": self.export_ocr_text,
            "print": self.print_pages,
            "rotate_left": self.rotate_left,
            "rotate_right": self.rotate_right,
            "bright_up": lambda: self._enhance_current(1.12, 1.0),
            "bright_dn": lambda: self._enhance_current(0.9, 1.0),
            "contrast_up": lambda: self._enhance_current(1.0, 1.15),
            "contrast_dn": lambda: self._enhance_current(1.0, 0.88),
            "autocrop": self.autocrop_current,
            "autoname": self.auto_name_pages,
            "undo": self.undo_delete,
            "delete": self.delete_page,
            "clear": self.clear_all,
            "move_up": self.move_up,
            "move_down": self.move_down,
            "search": self.search_pdfs,
            "merge": self.merge_pdfs,
            "split": self.split_pdfs,
            "report": self.monthly_report,
            "zoom_in": lambda: self._zoom_thumbs(1.15),
            "zoom_out": lambda: self._zoom_thumbs(0.87),
            "options": self.open_options,
            "profiles": self.open_profiles,
        }
        self._sc = {}
        custom = self._opts.get("shortcuts", {}) or {}
        for sid, label, default in SHORTCUTS:
            fn = methods.get(sid)
            if fn is None:
                continue
            scut = QtWidgets.QShortcut(self)
            scut.activated.connect(fn)
            key = custom.get(sid, default)
            if key:
                scut.setKey(QtGui.QKeySequence(key))
            self._sc[sid] = scut

    def _apply_shortcut(self, sid, key):
        scut = self._sc.get(sid)
        if scut is not None:
            scut.setKey(QtGui.QKeySequence(key) if key else QtGui.QKeySequence())

    def _cur_shortcut(self, sid, default):
        return (self._opts.get("shortcuts", {}) or {}).get(sid, default)

    def show_shortcuts(self):
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Keyboard Shortcuts"); dlg.resize(560, 560)
        lay = QtWidgets.QVBoxLayout(dlg)
        lay.addWidget(QtWidgets.QLabel(
            "Shortcut badalne ke liye: row chuno \u2192 'Naya' box me key dabao \u2192 Assign.\n"
            "To change: pick a row, press the new key in 'New', then Assign."))
        tbl = QtWidgets.QTableWidget(len(SHORTCUTS), 2)
        tbl.setHorizontalHeaderLabels(["Action", "Shortcut"])
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        tbl.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        tbl.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        tbl.horizontalHeader().setStretchLastSection(True); tbl.setColumnWidth(0, 300)

        def refill():
            for r, (sid, label, default) in enumerate(SHORTCUTS):
                tbl.setItem(r, 0, QtWidgets.QTableWidgetItem(label))
                tbl.setItem(r, 1, QtWidgets.QTableWidgetItem(self._cur_shortcut(sid, default) or "—"))
        refill()
        lay.addWidget(tbl, 1)

        row = QtWidgets.QHBoxLayout()
        kseq = QtWidgets.QKeySequenceEdit()
        row.addWidget(QtWidgets.QLabel("Naya / New:")); row.addWidget(kseq, 1)
        b_assign = QtWidgets.QPushButton("Assign"); b_unassign = QtWidgets.QPushButton("Hatao / Unassign")
        row.addWidget(b_assign); row.addWidget(b_unassign)
        lay.addLayout(row)

        def sel():
            r = tbl.currentRow()
            return (SHORTCUTS[r][0], r) if r >= 0 else (None, -1)

        def do_assign():
            sid, r = sel()
            if sid is None:
                return
            key = kseq.keySequence().toString()
            if not key:
                return
            sm = self._opts.setdefault("shortcuts", {})
            for osid, olabel, odef in SHORTCUTS:      # clear conflicts
                if osid != sid and self._cur_shortcut(osid, odef) == key:
                    sm[osid] = ""; self._apply_shortcut(osid, "")
            sm[sid] = key; self._apply_shortcut(sid, key); self._save_opts()
            kseq.clear(); refill()

        def do_unassign():
            sid, r = sel()
            if sid is None:
                return
            self._opts.setdefault("shortcuts", {})[sid] = ""
            self._apply_shortcut(sid, ""); self._save_opts(); refill()

        b_assign.clicked.connect(do_assign); b_unassign.clicked.connect(do_unassign)

        bottom = QtWidgets.QHBoxLayout()
        b_reset = QtWidgets.QPushButton("Restore Defaults")

        def do_reset():
            self._opts["shortcuts"] = {}
            for sid, label, default in SHORTCUTS:
                self._apply_shortcut(sid, default)
            self._save_opts(); refill()
        b_reset.clicked.connect(do_reset)
        b_ok = QtWidgets.QPushButton("OK"); b_ok.clicked.connect(dlg.accept)
        bottom.addWidget(b_reset); bottom.addStretch(1); bottom.addWidget(b_ok)
        lay.addLayout(bottom)
        dlg.exec_()

    def _stats_url(self):
        if "stats_url" in self._config:
            return self._config["stats_url"]      # user override (may be "" to disable)
        return DEFAULT_STATS_URL

    def _get_client_id(self):
        cid = self._config.get("client_id")
        if not cid:
            import uuid
            cid = uuid.uuid4().hex[:16]
            self._config["client_id"] = cid
            try:
                save_config(self._config)
            except Exception:
                pass
        return cid

    def _set_stats_display(self, stats):
        if stats and isinstance(stats, tuple) and len(stats) == 3:
            ws = getattr(self, "_world_stats", {}) or {}
            ws.update({"total": stats[0], "today": stats[1], "online": stats[2]})
            self._world_stats = ws
        self._update_sidebar_stats()

    def _stats_failed(self):
        # network na ho to bhi personal stats dikhti rahein
        self._update_sidebar_stats()

    def _refresh_stats(self, action="ping"):
        url = self._stats_url()
        if not url:
            self._set_stats_display(None)
            return
        self._stats_worker = StatsWorker(url, self._get_client_id(), action=action)
        self._stats_worker.got.connect(lambda t, d, o: self._set_stats_display((t, d, o)))
        self._stats_worker.got_full.connect(self._on_world_stats)
        self._stats_worker.failed.connect(self._stats_failed)
        self._stats_worker.start()

    def _report_scan_stat(self, n):
        url = self._stats_url()
        if not url or n <= 0:
            return
        self._scan_reporter = StatsWorker(url, self._get_client_id(), action="scan", n=n)
        self._scan_reporter.got.connect(lambda t, d, o: self._set_stats_display((t, d, o)))
        self._scan_reporter.got_full.connect(self._on_world_stats)
        self._scan_reporter.start()

    def set_stats_url(self):
        cur = self._stats_url()
        url, ok = QtWidgets.QInputDialog.getText(
            self, "Stats server URL",
            "Google Apps Script Web app URL (…/exec):\n(khaali karke stats band ho jayenge)",
            text=cur)
        if not ok:
            return
        self._config["stats_url"] = url.strip()
        try:
            save_config(self._config)
        except Exception:
            pass
        self._set_stats_display(None)
        self._refresh_stats()

    def set_scanner_ip(self):
        ip, ok = QtWidgets.QInputDialog.getText(self, "Scanner IP", "Scanner IP:", text=self.ip_field.text())
        if ok:
            self.ip_field.setText(ip.strip()); self.check_connection()

    def undo_delete(self):
        if not self._undo_stack:
            self.status.showMessage("Undo ke liye kuch nahi", 3000); return
        path, row = self._undo_stack.pop()
        if not os.path.exists(path):
            return
        icon = QtGui.QIcon(self._make_thumb(path))
        item = QtWidgets.QListWidgetItem(icon, "Page")
        item.setData(QtCore.Qt.UserRole, path); item.setTextAlignment(QtCore.Qt.AlignHCenter)
        self.list.insertItem(min(row, self.list.count()), item)
        self.list.setCurrentItem(item); self.list.clearSelection()
        self._renumber_pages(); self._dirty = True
        self._update_status(); self._update_empty_state()

    def _renumber_pages(self):
        show = self._opts.get("show_page_numbers", True)
        for i in range(self.list.count()):
            it = self.list.item(i)
            title = it.data(TITLE_ROLE)
            if title:
                it.setText(title)
            else:
                it.setText(("Page %d" % (i + 1)) if show else "")

    def _update_empty_state(self):
        empty = self.list.count() == 0
        use_dash = empty and bool(self._opts.get("ui_dashboard", True)) and hasattr(self, "_dash")
        if hasattr(self, "_dash"):
            self._dash.setVisible(use_dash)
        self._empty_lbl.setVisible(empty and not use_dash)
        if empty:
            r = self.list.viewport().rect()
            self._empty_lbl.setGeometry(r)
            if hasattr(self, "_dash"):
                self._dash.setGeometry(r)
                try:
                    names = [os.path.basename(p) for p in (self._recent or [])[:3]]
                    self._dash_recent.setText(("Recent: " + "  ·  ".join(names)) if names else "")
                except Exception:
                    pass

    def _apply_thumb_zoom(self, w):
        w = max(80, min(340, int(w)))
        h = int(w * 4 / 3)
        self._thumb_w, self._thumb_h = w, h
        self.list.setIconSize(QtCore.QSize(w, h))
        self.list.setGridSize(QtCore.QSize(w + 24, h + 34))

    def _zoom_thumbs(self, factor):
        self._apply_thumb_zoom(self._thumb_w * factor)

    def eventFilter(self, obj, ev):
        if obj is self.list.viewport():
            if ev.type() == QtCore.QEvent.Resize:
                self._empty_lbl.setGeometry(self.list.viewport().rect())
                if hasattr(self, "_dash"):
                    self._dash.setGeometry(self.list.viewport().rect())
            elif ev.type() == QtCore.QEvent.Wheel and (ev.modifiers() & QtCore.Qt.ControlModifier):
                # Ctrl + mouse scroll -> zoom thumbnails in/out (like NAPS2)
                self._apply_thumb_zoom(self._thumb_w * (1.15 if ev.angleDelta().y() > 0 else 0.87))
                return True
        return super().eventFilter(obj, ev)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        exts = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".pdf")
        files = [url.toLocalFile() for url in e.mimeData().urls()
                 if url.toLocalFile().lower().endswith(exts)]
        if files:
            self._start_import(files, "normal")

    def _after_save_action(self, out):
        act = self._opts.get("after_save", "nothing")
        try:
            if act == "open":
                os.startfile(out)
            elif act == "folder":
                os.startfile(os.path.dirname(out))
        except Exception:
            pass

    # ---- Update check ----
    def check_updates(self, silent=False):
        self._upd_worker = UpdateChecker()
        self._upd_worker.result.connect(
            lambda tag, url, s=silent: self._on_update_result(tag, url, s))
        self._upd_worker.start()

    def _on_update_result(self, tag, url, silent):
        def _n(s):
            m = re.findall(r"\d+", s or "")
            return int(m[0]) if m else -1
        if not tag:
            if not silent:
                self._warn("Update check nahi ho paya. Internet chal raha hai?")
            return
        if _n(tag) > _n(VERSION):
            # Sidebar me banner dikhao — wahi se EK click me update
            self._show_update_banner(tag)
            if not silent:
                r = QtWidgets.QMessageBox.question(
                    self, "Naya version aa gaya",
                    "ApneScan ka naya version %s aa gaya hai (aapke paas v%s hai).\n\n"
                    "Abhi download karke update kar dein? (App khud band hokar nayi khul jayegi)"
                    % (tag, VERSION),
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
                if r == QtWidgets.QMessageBox.Yes:
                    self._start_self_update()
        elif not silent:
            QtWidgets.QMessageBox.information(
                self, "Update", "Aap latest version (v%s) par hi hain." % VERSION)

    def _show_update_banner(self, tag):
        self._latest_tag = tag
        try:
            self.update_box.setText(self.L(
                "🔔 Naya update %s aa gaya!\n⬇ Ek click me update karo" % tag,
                "🔔 Update %s is available!\n⬇ One-click update" % tag))
            self.update_box.setEnabled(True)
            self.update_box.show()
        except Exception:
            pass

    def _sidebar_update_clicked(self):
        try:
            self.update_box.setEnabled(False)
            self.update_box.setText("⬇ Download ho raha hai…\n(app istemal karte rahiye)")
        except Exception:
            pass
        self._start_self_update()

    def _reset_update_banner(self):
        """Download fail hone par banner wapas clickable karo."""
        try:
            if getattr(self, "_latest_tag", ""):
                self._show_update_banner(self._latest_tag)
        except Exception:
            pass

    def _start_self_update(self):
        """Naya version website se download karke KHUD install karo (feature 36)."""
        url = "https://apnescan.apnesoft.com/ApneScan.exe"

        def job():
            import urllib.request as U
            fd, tmp = tempfile.mkstemp(suffix=".exe")
            os.close(fd)
            req = U.Request(url, headers={"User-Agent": "ApneScan"})
            with U.urlopen(req, timeout=180) as r, open(tmp, "wb") as fh:
                shutil.copyfileobj(r, fh)
            if os.path.getsize(tmp) < 5_000_000:
                raise RuntimeError("Download adhoora laga (%d bytes)" % os.path.getsize(tmp))
            return tmp

        def done(res):
            if isinstance(res, Exception):
                self._reset_update_banner()
                self._warn("Update download nahi ho paya:\n%s\n\n"
                           "Website se khud download kar lein: apnescan.apnesoft.com" % res)
                try:
                    import webbrowser
                    webbrowser.open(DOWNLOAD_PAGE)
                except Exception:
                    pass
                return
            self._apply_downloaded_update(res)
        self._run_bg(job, done, "Naya version download ho raha hai… (app chalti rahegi)")

    def _apply_downloaded_update(self, tmp_exe):
        if not getattr(sys, "frozen", False):
            self._reveal_in_explorer(tmp_exe)
            QtWidgets.QMessageBox.information(
                self, "Download ho gaya", "Nayi exe yahan hai:\n%s" % tmp_exe)
            return
        cur = os.path.abspath(sys.executable)
        bat = os.path.join(tempfile.gettempdir(), "apnescan_update.bat")
        try:
            with open(bat, "w") as fh:
                fh.write('@echo off\r\n'
                         'ping 127.0.0.1 -n 4 > nul\r\n'
                         'copy /y "%s" "%s" > nul\r\n'
                         'start "" "%s"\r\n'
                         'del "%%~f0"\r\n' % (tmp_exe, cur, cur))
        except Exception as exc:
            self._warn("Update script fail: %s" % exc)
            return
        r = QtWidgets.QMessageBox.question(
            self, "Update taiyar",
            "Naya version download ho gaya.\n\nApp ab band hogi aur nayi version "
            "khud khul jayegi. Theek?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if r == QtWidgets.QMessageBox.Yes:
            try:
                os.startfile(bat)
            except Exception as exc:
                self._warn("Update start nahi hua: %s" % exc)
                return
            QtWidgets.QApplication.quit()

    def export_settings(self):
        """Saari settings ek file me — naye PC par le jaane ke liye."""
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Settings export karein",
            os.path.join(os.path.expanduser("~"), "apnescan_settings.json"),
            "Settings (*.json)")
        if not out:
            return
        try:
            shutil.copy2(CONFIG_PATH, out)
        except Exception as exc:
            self._warn("Export fail: %s" % exc); return
        QtWidgets.QMessageBox.information(
            self, "Ho gaya", "Settings export ho gayi:\n%s\n\nNaye PC par: Settings → "
            "Import settings… se wapas le aana." % out)

    def import_settings(self):
        f, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Settings file chuno", os.path.expanduser("~"), "Settings (*.json)")
        if not f:
            return
        try:
            with open(f, "r", encoding="utf-8") as fh:
                json.load(fh)          # valid json hai?
            shutil.copy2(f, CONFIG_PATH)
        except Exception as exc:
            self._warn("Import fail (file sahi nahi lagti): %s" % exc); return
        QtWidgets.QMessageBox.information(
            self, "Ho gaya", "Settings aa gayi. App band karke dobara kholein.")

    def toggle_touch_mode(self):
        self._opts["touch_mode"] = bool(self.act_touch.isChecked())
        self._save_opts()
        self._apply_style()
        self.status.showMessage(
            "Touch mode %s — poora asar app dobara kholne par." %
            ("ON" if self._opts["touch_mode"] else "OFF"), 5000)

    def open_crash_report(self):
        if os.path.exists(CRASH_PATH):
            self._open_path(CRASH_PATH)
        else:
            QtWidgets.QMessageBox.information(
                self, "Report", "Koi error report nahi hai — sab theek chal raha hai! 🙂")

    # ---- Share (WhatsApp / Email) ----
    def _pick_share_pdf(self):
        """Sabse aakhri saved PDF; koi na ho to user se file chunwao."""
        for p in (self._recent or []):
            if p and os.path.exists(p):
                return p
        start = self._opts.get("save_folder", os.path.expanduser("~"))
        f, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Kaunsi PDF bhejni hai?", start, "PDF (*.pdf)")
        return f or None

    def _reveal_in_explorer(self, path):
        try:
            import subprocess
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        except Exception:
            try:
                os.startfile(os.path.dirname(path))
            except Exception:
                pass

    def _copy_file_to_clipboard(self, path):
        """File ko clipboard par daalo taaki WhatsApp/Email me Ctrl+V se attach ho."""
        try:
            md = QtCore.QMimeData()
            md.setUrls([QtCore.QUrl.fromLocalFile(os.path.abspath(path))])
            QtWidgets.QApplication.clipboard().setMimeData(md)
            return True
        except Exception:
            return False

    def share_whatsapp(self, pdf=None):
        if not isinstance(pdf, str) or not pdf:
            pdf = self._pick_share_pdf()
        if not pdf:
            return
        copied = self._copy_file_to_clipboard(pdf)
        self._reveal_in_explorer(pdf)
        try:
            import webbrowser
            # WhatsApp Desktop installed ho to wahi khulega, warna web
            webbrowser.open("https://wa.me/")
        except Exception:
            pass
        self._pstats_bump(shared=1)
        steps = ("WhatsApp khul raha hai. Chat chuno, phir:\n\n"
                 "1) Chat me Ctrl+V dabao (file copy ho chuki hai)%s\n"
                 "2) Ya Explorer se PDF ko chat par kheench (drag) kar chhod do.\n\n"
                 "File: %s") % ("" if copied else " (copy nahi ho payi)", pdf)
        QtWidgets.QMessageBox.information(self, "WhatsApp par bhejo", steps)

    def share_email(self, pdf=None):
        if not isinstance(pdf, str) or not pdf:
            pdf = self._pick_share_pdf()
        if not pdf:
            return
        self._pstats_bump(shared=1)
        # Pehle Outlook try karo (attachment ke saath draft khulta hai)
        if HAS_W32:
            try:
                _pythoncom.CoInitialize()
                try:
                    ol = _w32.Dispatch("Outlook.Application")
                    mail = ol.CreateItem(0)
                    mail.Subject = os.path.splitext(os.path.basename(pdf))[0]
                    mail.Attachments.Add(os.path.abspath(pdf))
                    mail.Display(True)
                    return
                finally:
                    try:
                        _pythoncom.CoUninitialize()
                    except Exception:
                        pass
            except Exception:
                pass
        # Fallback: default mail app (mailto attachment support nahi karta,
        # isliye file clipboard par + Explorer me dikha do)
        self._copy_file_to_clipboard(pdf)
        self._reveal_in_explorer(pdf)
        try:
            import webbrowser
            import urllib.parse as _up
            webbrowser.open("mailto:?subject=%s"
                            % _up.quote(os.path.splitext(os.path.basename(pdf))[0]))
        except Exception:
            pass
        QtWidgets.QMessageBox.information(
            self, "Email se bhejo",
            "Email draft khul raha hai. PDF attach karne ke liye:\n\n"
            "1) Email me Ctrl+V dabao (file copy ho chuki hai)\n"
            "2) Ya Explorer se PDF ko email par kheench kar chhod do.\n\n"
            "File: %s" % pdf)

    # ---- PDF compress tool ----
    def compress_pdf_tool(self, src_pdf=None):
        """Abhi ke pages (ya koi bhi purani PDF) ko chhota karke alag PDF banao —
        portal ki 200KB/500KB/1MB/2MB limit ke hisaab se."""
        if not isinstance(src_pdf, str) or not src_pdf:
            src_pdf = None
        paths = self._ordered_paths()
        if src_pdf is None and not paths:
            start = self._opts.get("save_folder", os.path.expanduser("~"))
            src_pdf, _ = QtWidgets.QFileDialog.getOpenFileName(
                self, "Kaunsi PDF chhoti karni hai?", start, "PDF (*.pdf)")
            if not src_pdf:
                return
        targets = ["200 KB (portal upload)", "500 KB", "1 MB", "2 MB",
                   "Sirf quality kam karo (size ki limit nahi)"]
        choice, ok = QtWidgets.QInputDialog.getItem(
            self, "PDF chhota karo", "Kitni chhoti file chahiye?", targets, 2, False)
        if not ok:
            return
        idx = targets.index(choice)
        limit = {0: 200 * 1024, 1: 500 * 1024, 2: 1024 * 1024, 3: 2 * 1024 * 1024}.get(idx)
        if src_pdf:
            pages = pdf_to_images(src_pdf, self._tmpdir)
            if not pages:
                self._warn("Is PDF se pages nahi nikle.\nBehtar result ke liye 'PyMuPDF' install karein:\n  py -3.12-32 -m pip install PyMuPDF")
                return
            default = os.path.splitext(src_pdf)[0] + "_small.pdf"
        else:
            pages = paths
            base = self._build_filename(".pdf", paths=paths)
            default = (base[:-4] if base.lower().endswith(".pdf") else base) + "_small.pdf"
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Chhoti PDF save karein", default, "PDF (*.pdf)")
        if not out:
            return
        if not out.lower().endswith(".pdf"):
            out += ".pdf"

        def job():
            size = self._write_compressed_pdf(pages, out, limit)
            saved = 0
            if src_pdf:
                try:
                    saved = os.path.getsize(src_pdf) - size
                except Exception:
                    saved = 0
            return (size, saved)

        def done(res):
            if isinstance(res, Exception):
                self._warn("Compress fail:\n%s" % res); return
            size, saved = res
            self._remember_save_dir(out)
            if saved > 0:
                self._pstats_bump(saved_bytes=saved)
            kb = size / 1024.0
            pretty = ("%.0f KB" % kb) if kb < 1024 else ("%.1f MB" % (kb / 1024.0))
            note = ""
            if limit and size > limit:
                note = "\n\n(Itni chhoti nahi ho payi — pages bahut hain. Kam pages chunein ya badi limit lein.)"
            QtWidgets.QMessageBox.information(
                self, "Ho gaya", "Chhoti PDF save ho gayi (%s):\n%s%s" % (pretty, out, note))
        self._run_bg(job, done, "PDF chhoti ho rahi hai (compress)…")

    def _write_compressed_pdf(self, img_paths, out, limit_bytes=None):
        """JPEG quality/size ghata-ghata kar PDF banao jab tak limit ke andar na aaye.
        Returns final size in bytes."""
        combos = [(75, 1.0), (60, 1.0), (45, 1.0), (60, 0.8), (45, 0.8),
                  (35, 0.8), (45, 0.65), (35, 0.65), (30, 0.5), (25, 0.4)]
        if limit_bytes is None:
            combos = [(60, 1.0)]
        data = None
        for q, s in combos:
            buf = io.BytesIO()
            imgs = []
            try:
                for p in img_paths:
                    im = Image.open(p).convert("RGB")
                    if s < 1.0:
                        im = im.resize((max(1, int(im.width * s)),
                                        max(1, int(im.height * s))), Image.LANCZOS)
                    b2 = io.BytesIO()
                    im.save(b2, "JPEG", quality=q)
                    im.close()
                    b2.seek(0)
                    im2 = Image.open(b2)
                    im2.load()
                    imgs.append(im2)
                imgs[0].save(buf, "PDF", save_all=True, append_images=imgs[1:], resolution=200.0)
            finally:
                for im in imgs:
                    try:
                        im.close()
                    except Exception:
                        pass
            data = buf.getvalue()
            if limit_bytes is None or len(data) <= limit_bytes:
                break
        with open(out, "wb") as fh:
            fh.write(data)
        return len(data)

    def _run_wizard(self):
        wiz = SetupWizard(self, self._opts.get("language", "hi"))
        if wiz.exec_() == QtWidgets.QDialog.Accepted:
            self._opts["scanner_method"] = wiz.result_method
            if wiz.result_method == "wia":
                self._opts["wia_device_id"] = wiz.result_wia_id
            self._save_opts()
            prof = self._selected_profile()
            if prof is None:
                prof = {"name": "Default", "source_name": wiz.result_device_name,
                        "dpi": 200, "color": "gray", "duplex": False}
                self._profiles.append(prof)
            elif wiz.result_method == "twain" and wiz.result_device_name:
                prof["source_name"] = wiz.result_device_name
            self._save_profiles(); self._refresh_profile_combo()
            i = self.cmb_profile.findText(prof.get("name"))
            if i >= 0:
                self.cmb_profile.setCurrentIndex(i)
        self._config["setup_done"] = True
        save_config(self._config)
        self._update_status()

    def choose_scan_method(self):
        cur = self._opts.get("scanner_method", "twain")
        labels = ["escl (network duplex - no extra software)", "twain (USB duplex)", "wia", "naps2 (needs NAPS2)"]
        keys = ["escl", "twain", "wia", "naps2"]
        idx = keys.index(cur) if cur in keys else 0
        choice, ok = QtWidgets.QInputDialog.getItem(
            self, "Scan method", "Method:", labels, idx, False)
        if not ok:
            return
        method = keys[labels.index(choice)]
        self._opts["scanner_method"] = method
        if method == "escl":
            ip = self.ip_field.text().strip()
            if not ip:
                ip, oki = QtWidgets.QInputDialog.getText(
                    self, "Scanner IP", "Scanner ka network IP daalo (jaise 192.168.1.8):")
                if oki:
                    ip = ip.strip(); self.ip_field.setText(ip)
            self._opts["scanner_ip"] = ip
        elif method == "naps2":
            path = self._opts.get("naps2_path") or find_naps2()
            if not path or not os.path.exists(path):
                path, okp = QtWidgets.QInputDialog.getText(
                    self, "NAPS2",
                    "NAPS2.Console.exe ka poora path daalein\n"
                    "(jaise C:\\Program Files\\NAPS2\\NAPS2.Console.exe):")
                path = path.strip() if okp else ""
            self._opts["naps2_path"] = path
            pname, okn = QtWidgets.QInputDialog.getText(
                self, "NAPS2 profile",
                "NAPS2 ka profile naam (bilkul waisa hi jaisa NAPS2 me hai,\n"
                "jaise: 150dpi Double Side):",
                text=self._opts.get("naps2_profile", ""))
            if okn:
                self._opts["naps2_profile"] = pname.strip()
            if not path:
                self._warn("NAPS2 ka path nahi mila. Pehle NAPS2 install karo.")
        elif method == "wia":
            try:
                devs = list_wia_sources()
            except Exception as exc:
                self._warn(friendly_error(exc, self._opts.get("language", "hi"))); devs = []
            if devs:
                names = [n for _i, n in devs]
                name, ok2 = QtWidgets.QInputDialog.getItem(self, "WIA scanner", "Scanner:", names, 0, False)
                if ok2 and name:
                    for _id, n in devs:
                        if n == name:
                            self._opts["wia_device_id"] = _id; break
        else:
            # TWAIN: let the user pick the real driver source (e.g. "HP TWAIN USB"),
            # which supports duplex — unlike the WIA-bridge sources.
            try:
                names = list_sources(int(self.winId()))
            except Exception as exc:
                self._warn(friendly_error(exc, self._opts.get("language", "hi"))); names = []
            if names:
                name, ok2 = QtWidgets.QInputDialog.getItem(
                    self, "TWAIN scanner",
                    "Scanner chuno (duplex ke liye 'HP TWAIN USB' jaisa asli driver chuno,\nWIA-... wala nahi):",
                    names, 0, False)
                if ok2 and name:
                    prof = self._selected_profile()
                    if prof is not None:
                        prof["source_name"] = name
                        self._save_profiles()
                        self._load_profile_to_panel(prof)
        self._save_opts(); self._update_status(); self._refresh_method_label()

    def choose_language(self):
        cur = self._opts.get("language", "hi")
        LANGS = [("Hindi", "hi"), ("English", "en"),
                 ("मराठी Marathi (adhuri)", "mr"), ("ગુજરાતી Gujarati (adhuri)", "gu"),
                 ("ਪੰਜਾਬੀ Punjabi (adhuri)", "pa"), ("தமிழ் Tamil (adhuri)", "ta")]
        idx = next((i for i, (_n, c) in enumerate(LANGS) if c == cur), 0)
        lang, ok = QtWidgets.QInputDialog.getItem(
            self, "Language", "Bhasha / Language:", [n for n, _c in LANGS], idx, False)
        if not ok:
            return
        self._opts["language"] = dict((n, c) for n, c in LANGS).get(lang, "hi")
        self._save_opts()
        QtWidgets.QMessageBox.information(
            self, "Language",
            "Language badal gayi. App band karke dobara kholein taaki poora asar dikhe."
            if self._opts["language"] == "hi"
            else "Language changed. Close and reopen the app to fully apply.")

    def show_help(self):
        HelpDialog(self, self._opts.get("language", "hi")).exec_()

    def run_duplex_test(self):
        """Scan with Both-side ON, then report whether the scanner really
        returned 2 sides — with the duplex value used and any error."""
        if self.list.count() > 0:
            if QtWidgets.QMessageBox.question(
                    self, "Duplex Test",
                    "Test se pehle maujuda pages hat jayenge. Aage badhein?"
                    if self._lang == "hi" else "This clears current pages. Continue?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                    QtWidgets.QMessageBox.No) != QtWidgets.QMessageBox.Yes:
                return
            for p in self._ordered_paths():
                try:
                    os.remove(p)
                except Exception:
                    pass
            self.list.clear(); self._update_empty_state()

        QtWidgets.QMessageBox.information(
            self, "Duplex Test",
            "Feeder me 1 kaagaz (dono taraf likha hua) rakho, phir OK dabao.\n"
            "Test us kaagaz ko scan karke batayega ki dono taraf aaya ya nahi."
            if self._lang == "hi" else
            "Put ONE double-sided sheet in the feeder, then press OK.")

        prof = self._selected_profile()
        if prof is None:
            self._warn("Pehle profile/scanner chuno."); return
        method = self._opts.get("scanner_method", "twain")
        dpi = 200
        pixel = "gray"
        got = {"pages": 0, "err": None, "trace": []}
        WIA_FMT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"

        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            if method == "wia":
                try:
                    _pythoncom.CoInitialize()
                except Exception:
                    pass
                try:
                    tr = got["trace"]
                    wid = self._opts.get("wia_device_id")

                    def _connect():
                        dm = _w32.Dispatch("WIA.DeviceManager")
                        for i in range(1, dm.DeviceInfos.Count + 1):
                            info = dm.DeviceInfos.Item(i)
                            if wid is None or info.DeviceID == wid:
                                return info.Connect()
                        return None

                    def _setp(obj, pid, val):
                        try:
                            for p in obj.Properties:
                                if int(p.PropertyID) == pid:
                                    p.Value = val
                                    return True
                        except Exception:
                            pass
                        return False

                    def _valid_dpi(it, want):
                        try:
                            for p in it.Properties:
                                if int(p.PropertyID) == 6147:
                                    vals = [int(x) for x in p.SubType.Values]
                                    if vals:
                                        return min(vals, key=lambda v: abs(v - want))
                        except Exception:
                            pass
                        return want

                    def _try_transfers(item, n=3):
                        pages = 0; err = None
                        for k in range(n):
                            try:
                                try:
                                    item.Transfer(WIA_FMT_BMP)
                                except Exception:
                                    item.Transfer()
                                pages += 1
                            except Exception as e:
                                err = str(e); break
                        return pages, err

                    # Each strategy: (label, set-3088?, set-3096?, set item props?, plain-only?)
                    strategies = [
                        ("A duplex=5, no pages, no item-props", 5, False, False),
                        ("B duplex=5 + pages=0", 5, True, False),
                        ("C duplex=5 + item props(res/depth)", 5, False, True),
                        ("D duplex=7", 7, False, False),
                        ("E duplex=0x8005(32773)", 0x8005, False, False),
                    ]
                    best = 0
                    for label, dupval, setpages, setitem in strategies:
                        dev = None
                        try:
                            dev = _connect()
                            if dev is None:
                                tr.append("%s -> connect fail" % label); continue
                            ok88 = _setp(dev, 3088, dupval)
                            if setpages:
                                _setp(dev, 3096, 0)
                            it = dev.Items[1]
                            if setitem:
                                _setp(it, 4104, 2)  # grayscale
                                dv = _valid_dpi(it, 200)
                                _setp(it, 6147, dv); _setp(it, 6148, dv)
                            pages, err = _try_transfers(it, 3)
                            tr.append("%s -> set88=%s pages=%d err=%s"
                                      % (label, ok88, pages, err or "none"))
                            if pages > best:
                                best = pages
                            if pages >= 1:
                                tr.append(">>> Ye tarika chala! (pages=%d). Baaki skip." % pages)
                                break
                        except Exception as e:
                            tr.append("%s -> EXC %s" % (label, e))
                        finally:
                            try:
                                if dev is not None:
                                    dev = None
                            except Exception:
                                pass
                    got["pages"] = best
                finally:
                    try:
                        _pythoncom.CoUninitialize()
                    except Exception:
                        pass
            else:
                def _count(img):
                    got["pages"] += 1
                try:
                    scan_pages(int(self.winId()), prof.get("source_name"), dpi, pixel, True, _count)
                except Exception as exc:
                    got["err"] = str(exc)
        except Exception as exc:
            got["err"] = str(exc)
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()

        lines = []
        lines.append("===== ApneScan Duplex Test =====")
        lines.append("Time: %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        lines.append("Scan method: %s" % method)
        lines.append("Device: %s" % (prof.get("source_name") or "(none)"))
        lines.append("Test settings: 1 sheet, Both-side ON, 200 dpi")
        lines.append("")
        lines.append("Pages returned: %d" % got["pages"])
        lines.append("Error: %s" % (got["err"] or "(none)"))
        if got["trace"]:
            lines.append("")
            lines.append("--- Detail (step by step) ---")
            for t in got["trace"]:
                lines.append("  " + t)
        lines.append("")
        if got["pages"] >= 2:
            lines.append("RESULT: DUPLEX WORKING \u2705  (scanner ne dono taraf diye)")
        elif got["pages"] == 1:
            lines.append("RESULT: SINGLE SIDE ONLY \u274c  (scanner ne sirf ek taraf diya)")
            lines.append("Matlab: is method/driver par duplex nahi mil raha.")
        else:
            lines.append("RESULT: koi page nahi aaya (scanner/feeder check karo)")
        lines.append("================================")
        report = "\n".join(lines)

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Duplex Test Result"); dlg.resize(560, 420)
        lay = QtWidgets.QVBoxLayout(dlg)
        lay.addWidget(QtWidgets.QLabel(
            "Ye result copy karke share karein:" if self._lang == "hi" else "Copy & share this result:"))
        box = QtWidgets.QPlainTextEdit(); box.setPlainText(report); box.setReadOnly(True)
        box.setStyleSheet("font-family: Consolas, monospace; font-size: 12px;")
        lay.addWidget(box, 1)
        row = QtWidgets.QHBoxLayout()
        b_copy = QtWidgets.QPushButton("Copy karo" if self._lang == "hi" else "Copy")
        b_copy.clicked.connect(lambda: (QtWidgets.QApplication.clipboard().setText(report),
                                        self.status.showMessage("Copy ho gaya", 3000)))
        b_close = QtWidgets.QPushButton("Band karo" if self._lang == "hi" else "Close")
        b_close.clicked.connect(dlg.accept)
        row.addWidget(b_copy); row.addStretch(1); row.addWidget(b_close)
        lay.addLayout(row)
        dlg.exec_()

    def _show_report(self, title, report):
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle(title); dlg.resize(640, 560)
        lay = QtWidgets.QVBoxLayout(dlg)
        lay.addWidget(QtWidgets.QLabel("Ye report copy karke share karein / Copy and share this report:"))
        box = QtWidgets.QPlainTextEdit(); box.setPlainText(report); box.setReadOnly(True)
        box.setStyleSheet("font-family: Consolas, monospace; font-size: 12px;")
        lay.addWidget(box, 1)
        row = QtWidgets.QHBoxLayout()
        b_copy = QtWidgets.QPushButton("Copy")
        def _copy():
            QtWidgets.QApplication.clipboard().setText(report)
            self.status.showMessage("Report copy ho gaya", 3000)
        b_copy.clicked.connect(_copy)
        b_save = QtWidgets.QPushButton("Save to file")
        def _save():
            p, _ = QtWidgets.QFileDialog.getSaveFileName(
                dlg, "Save report", os.path.join(os.path.expanduser("~"), "ApneScan_escl_test.txt"), "Text (*.txt)")
            if p:
                try:
                    with open(p, "w", encoding="utf-8") as fh:
                        fh.write(report)
                    self.status.showMessage("Saved: %s" % p, 5000)
                except Exception as exc:
                    self._warn("Save fail: %s" % exc)
        b_save.clicked.connect(_save)
        b_close = QtWidgets.QPushButton("Close"); b_close.clicked.connect(dlg.accept)
        row.addWidget(b_copy); row.addWidget(b_save); row.addStretch(1); row.addWidget(b_close)
        lay.addLayout(row)
        dlg.exec_()

    def run_escl_test(self):
        ip = self.ip_field.text().strip() or self._config.get("scanner_ip", "")
        if not ip:
            self._warn("Scanner IP set nahi hai. Settings \u2192 Scanner IP me daalein (jaise 192.168.1.8).")
            return
        try:
            self._state_timer.stop()
        except Exception:
            pass
        self.status.showMessage("eSCL test chal raha hai... (kuch second, scanner se baat kar rahe hain)", 0)
        self._escl_tester = EsclTestWorker(ip)
        self._escl_tester.done.connect(self._on_escl_test_done)
        self._escl_tester.start()

    def _on_escl_test_done(self, report):
        try:
            self._state_timer.start()
        except Exception:
            pass
        self.status.clearMessage()
        self._show_report("eSCL Test (network scan jaanch)", report)

    def run_diagnostics(self):
        """Collect scanner + app info and any errors into one report the user
        can copy and share. Runs a safe non-scanning probe of TWAIN and WIA."""
        import platform
        lines = []
        def add(t=""):
            lines.append(t)

        add("===== ApneScan Diagnostics =====")
        add("Version: %s" % VERSION)
        add("Time: %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        try:
            add("Windows: %s" % platform.platform())
            add("Python: %s (%d-bit)" % (platform.python_version(),
                                         64 if sys.maxsize > 2**32 else 32))
        except Exception:
            pass
        add("Libs: TWAIN=%s  WIA(pywin32)=%s  OCR=%s  Excel=%s  Barcode=%s"
            % (HAS_TWAIN, HAS_W32, HAS_OCR_LIBS, HAS_XLSX, HAS_ZBAR))
        # Tesseract (needed for document-name OCR) — show path + version so we can
        # tell "not installed" apart from "installed but no heading found".
        try:
            tcmd = ""
            try:
                tcmd = pytesseract.pytesseract.tesseract_cmd
            except Exception:
                tcmd = "(default: tesseract)"
            add("Tesseract cmd: %s" % tcmd)
            add("Tesseract cmd exists: %s" % (os.path.exists(tcmd) if (tcmd and tcmd != "tesseract") else "unknown"))
            try:
                add("Tesseract version: %s  -> WORKS" % pytesseract.get_tesseract_version())
            except Exception as _te:
                add("Tesseract version: NOT FOUND (%s)" % type(_te).__name__)
        except Exception:
            add("Tesseract: (pytesseract missing)")
        add("")
        add("Scan method: %s" % self._opts.get("scanner_method", "twain"))
        prof = self._selected_profile()
        add("Profile: %s" % (prof.get("name") if prof else "(none)"))
        add("Device (stored): %s" % (prof.get("source_name") if prof else "(none)"))
        try:
            add("Panel -> Source: %s | Sides: %s | DPI: %s | Depth: %s"
                % (self.cmb_source.currentText(), self.cmb_sides.currentText(),
                   self.cmb_dpi.currentText(), self.cmb_depth.currentText()))
            add("Fast mode: %s" % self.chk_fast.isChecked())
        except Exception:
            pass
        add("Scanner IP: %s" % self.ip_field.text())
        add("")

        # ---- TWAIN probe ----
        add("----- TWAIN sources -----")
        if HAS_TWAIN:
            try:
                names = list_sources(int(self.winId()))
                if names:
                    for n in names:
                        add("  * %s" % n)
                else:
                    add("  (koi TWAIN source nahi mila)")
            except Exception as exc:
                add("  TWAIN error: %s" % exc)
        else:
            add("  TWAIN (pytwain) install nahi hai")
        add("")

        # ---- WIA probe + duplex capability ----
        add("----- WIA devices -----")
        if HAS_W32:
            try:
                try:
                    _pythoncom.CoInitialize()
                except Exception:
                    pass
                devs = list_wia_sources()
                if not devs:
                    add("  (koi WIA device nahi mila)")
                for dev_id, dev_name in devs:
                    add("  * %s" % dev_name)
                    try:
                        dm = _w32.Dispatch("WIA.DeviceManager")
                        device = None
                        for i in range(1, dm.DeviceInfos.Count + 1):
                            info = dm.DeviceInfos.Item(i)
                            if info.DeviceID == dev_id:
                                device = info.Connect(); break
                        if device is not None:
                            for p in device.Properties:
                                try:
                                    pid = int(p.PropertyID)
                                except Exception:
                                    continue
                                if pid == 3086:
                                    add("      Duplex capable (handling caps value): %s" % p.Value)
                                if pid == 3088:
                                    add("      Current handling select: %s" % p.Value)
                    except Exception as exc:
                        add("      (device probe error: %s)" % exc)
            except Exception as exc:
                add("  WIA error: %s" % exc)
            finally:
                try:
                    _pythoncom.CoUninitialize()
                except Exception:
                    pass
        else:
            add("  WIA (pywin32) install nahi hai")
        add("")
        add("Last scan error: %s" % getattr(self, "_last_error", "(none)"))
        add("================================")

        report = "\n".join(lines)
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Test / Diagnostics"); dlg.resize(640, 560)
        lay = QtWidgets.QVBoxLayout(dlg)
        lay.addWidget(QtWidgets.QLabel(
            "Ye report copy karke share karein (scanner + error info):"
            if self._lang == "hi" else "Copy and share this report:"))
        box = QtWidgets.QPlainTextEdit(); box.setPlainText(report); box.setReadOnly(True)
        box.setStyleSheet("font-family: Consolas, monospace; font-size: 12px;")
        lay.addWidget(box, 1)
        row = QtWidgets.QHBoxLayout()
        b_copy = QtWidgets.QPushButton("Copy karo" if self._lang == "hi" else "Copy")
        def _copy():
            QtWidgets.QApplication.clipboard().setText(report)
            self.status.showMessage("Report copy ho gaya", 3000)
        b_copy.clicked.connect(_copy)
        b_save = QtWidgets.QPushButton("File me save karo" if self._lang == "hi" else "Save to file")
        def _save():
            p, _ = QtWidgets.QFileDialog.getSaveFileName(
                dlg, "Save report", os.path.join(os.path.expanduser("~"), "ApneScan_diagnostics.txt"),
                "Text (*.txt)")
            if p:
                try:
                    with open(p, "w", encoding="utf-8") as fh:
                        fh.write(report)
                    self.status.showMessage("Saved: %s" % p, 5000)
                except Exception as exc:
                    self._warn("Save fail: %s" % exc)
        b_save.clicked.connect(_save)
        b_close = QtWidgets.QPushButton("Band karo" if self._lang == "hi" else "Close")
        b_close.clicked.connect(dlg.accept)
        row.addWidget(b_copy); row.addWidget(b_save); row.addStretch(1); row.addWidget(b_close)
        lay.addLayout(row)
        dlg.exec_()

    def show_about(self):
        QtWidgets.QMessageBox.information(
            self, "About ApneScan",
            "ApneScan\nFree document scanning + PDF tool.\n"
            "TWAIN + WIA scanners support.\nVersion " + VERSION)

    def toggle_simple_mode(self):
        self._opts["simple_mode"] = bool(self.act_simple.isChecked())
        self._save_opts()
        self._apply_simple_mode()

    def _apply_simple_mode(self):
        simple = bool(self._opts.get("simple_mode"))
        for w in getattr(self, "_adv_btns", []):
            try:
                w.setVisible(not simple)
            except Exception:
                pass


    def show_whatsnew(self):
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle(tr("whatsnew", self._lang)); dlg.resize(560, 460)
        lay = QtWidgets.QVBoxLayout(dlg)
        br = QtWidgets.QTextBrowser(); br.setHtml(CHANGELOG_HTML); lay.addWidget(br)
        b = QtWidgets.QPushButton("Close" if self._lang == "en" else "Band karo")
        b.clicked.connect(dlg.accept); lay.addWidget(b)
        dlg.exec_()

    def send_feedback(self):
        title = tr("feedback", self._lang)
        text, ok = QtWidgets.QInputDialog.getMultiLineText(
            self, title,
            ("Aapka sujhav ya problem likhein:" if self._lang == "hi"
             else "Write your feedback or problem:"))
        if not ok or not text.strip():
            return
        email = self._opts.get("feedback_email", "").strip()
        if not email:
            email, ok2 = QtWidgets.QInputDialog.getText(
                self, title,
                ("Feedback kis email par bheju? (ek baar set ho jayega)"
                 if self._lang == "hi" else "Which email should feedback go to?"))
            if not ok2 or not email.strip():
                QtWidgets.QMessageBox.information(self, title, text)
                return
            self._opts["feedback_email"] = email.strip(); self._save_opts()
            email = email.strip()
        try:
            from PyQt5.QtGui import QDesktopServices
            body = QtCore.QUrl.toPercentEncoding(text).data().decode()
            url = QtCore.QUrl("mailto:%s?subject=ApneScan%%20Feedback&body=%s" % (email, body))
            QDesktopServices.openUrl(url)
        except Exception:
            QtWidgets.QMessageBox.information(self, title,
                                              "Email: %s\n\n%s" % (email, text))

    # ---- UI ----
    def _build_ui(self):
        central = QtWidgets.QWidget(); self.setCentralWidget(central)
        outer = QtWidgets.QVBoxLayout(central); outer.setContentsMargins(0, 0, 0, 0); outer.setSpacing(0)

        # ---------- Top toolbar ----------
        tbwrap = QtWidgets.QWidget(); tbwrap.setObjectName("toolbar")
        tb = QtWidgets.QHBoxLayout(tbwrap); tb.setContentsMargins(8, 4, 8, 4); tb.setSpacing(2)
        self._adv_btns = []

        tips = {
            "scan": "Scan shuru karo (F5)", "profiles": "Scan profiles banao/badlo",
            "ocr": "OCR: searchable PDF (dabakar on karo)", "fast": "Fast: 200 dpi + B&W, sabse tez",
            "import": "Computer se image daalo", "savepdf": "PDF save karo (tir se: sab/selected)",
            "images": "Pages ko JPG/PNG me save karo", "print": "Print karo (Ctrl+P)",
            "rotate": "Selected page ghumao", "up": "Page upar karo", "down": "Page neeche karo",
            "delete": "Selected page hatao (Ctrl+Z se wapas)", "clear": "Sab pages hatao",
            "language": "Bhasha badlo (Hindi/English)", "about": "App ke baare me",
        }
        def tbtn(kind, label, fn, advanced=False, checkable=False):
            b = QtWidgets.QToolButton()
            b.setToolTip(tips.get(kind, label))
            b.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
            b.setIcon(_make_icon(kind)); b.setIconSize(QtCore.QSize(28, 28))
            b.setText(label); b.setAutoRaise(True); b.setMinimumWidth(62)
            if checkable:
                b.setCheckable(True)
                b.toggled.connect(fn) if fn else None
            else:
                b.clicked.connect(fn)
            tb.addWidget(b)
            if advanced:
                self._adv_btns.append(b)
            return b

        self.btn_scan_top = tbtn("scan", tr("scan", self._lang), self.do_scan)
        self.btn_profiles = tbtn("profiles", tr("profiles", self._lang).replace("…", ""), self.open_profiles, advanced=True)
        self.chk_ocr = tbtn("ocr", "OCR", None, checkable=True)
        if not HAS_OCR_LIBS:
            self.chk_ocr.setEnabled(False)
        self.chk_fast = tbtn("fast", tr("fast", self._lang), None, checkable=True)
        self.chk_fast.setToolTip("Fast: 200 dpi + Black & White, bina extra processing")
        self.chk_fast.setChecked(bool(self._opts.get("fast_mode")))
        self.chk_fast.toggled.connect(self._on_fast_toggled)
        tb.addWidget(self._vsep())
        self.btn_import = tbtn("import", tr("import", self._lang), self.import_images, advanced=True)
        tbtn("import", self.L("Camera", "Camera"), self.scan_from_camera, advanced=True)
        self.btn_save_pdf = tbtn("savepdf", tr("save_pdf", self._lang), self.save_pdf)
        tbtn("images", "Save Images", self.save_images, advanced=True)
        self.btn_print = tbtn("print", "Print", self.print_all, advanced=True)
        tb.addWidget(self._vsep())
        tbtn("rotate", "Rotate", self.rotate_right, advanced=True)
        tbtn("up", tr("up", self._lang), self.move_up, advanced=True)
        tbtn("down", tr("down", self._lang), self.move_down, advanced=True)
        tbtn("delete", tr("delete", self._lang), self.delete_page, advanced=True)
        tbtn("clear", tr("clear", self._lang), self.clear_all, advanced=True)
        tb.addStretch(1)
        tbtn("language", "Language", self.choose_language)
        tbtn("about", tr("about", self._lang), self.show_about)
        pdfmenu = QtWidgets.QMenu(self.btn_save_pdf); pdfmenu.setToolTipsVisible(True)
        self._ma(pdfmenu, tr("save_all", self._lang), self.save_pdf_all,
                 "हिन्दी: Sabhi pages ki ek PDF banao.\nEnglish: Save all pages as one PDF.")
        self._ma(pdfmenu, tr("save_sel", self._lang), self.save_pdf_selected,
                 "हिन्दी: Sirf chune hue pages ki PDF.\nEnglish: PDF of only the selected pages.")
        self.btn_save_pdf.setMenu(pdfmenu)
        self.btn_save_pdf.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
        printmenu = QtWidgets.QMenu(self.btn_print); printmenu.setToolTipsVisible(True)
        self._ma(printmenu, "All print (sabhi pages)", self.print_all,
                 "हिन्दी: Saare pages print karo.\nEnglish: Print all pages.")
        self._ma(printmenu, "Selected print (chune hue)", self.print_selected,
                 "हिन्दी: Sirf chune hue pages print karo.\nEnglish: Print only the selected pages.")
        self._ma(printmenu, "ID print (2 ID ek page par)", self.print_ids,
                 "हिन्दी: ID cards 2-per-A4 print (kaagaz bachega).\nEnglish: Print IDs two per A4 sheet.")
        self._ma(printmenu, "ID print - sirf selected", self.print_ids_selected,
                 "हिन्दी: Sirf chune hue IDs 2-per-page.\nEnglish: Selected IDs, two per sheet.")
        self.btn_print.setMenu(printmenu)
        self.btn_print.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
        self._classic_toolbar = tbwrap
        outer.addWidget(tbwrap)
        # ---- UI #1: Ribbon toolbar (MS-Office jaisa — tabs me buttons) ----
        self.ribbon = QtWidgets.QTabWidget()
        self.ribbon.setObjectName("ribbon")
        self.ribbon.setMaximumHeight(96)

        def _ribbon_tab(pairs):
            w = QtWidgets.QWidget()
            h = QtWidgets.QHBoxLayout(w)
            h.setContentsMargins(8, 4, 8, 4); h.setSpacing(4)
            for icon, text, fn in pairs:
                b = QtWidgets.QToolButton()
                b.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
                b.setText("%s\n%s" % (icon, text))
                b.setAutoRaise(True); b.setMinimumWidth(66); b.setMinimumHeight(56)
                b.clicked.connect(fn)
                h.addWidget(b)
            h.addStretch(1)
            return w
        L = self.L
        self.ribbon.addTab(_ribbon_tab([
            ("🖨", L("Scan", "Scan"), self.do_scan),
            ("📷", "Camera", self.scan_from_camera),
            ("📥", "Import", self.import_images),
            ("🖼", "Photo", self.import_photos),
            ("💾", "Save PDF", self.save_pdf_all),
            ("🖨️", "Print", self.print_pages)]), L("🏠 Home", "🏠 Home"))
        self.ribbon.addTab(_ribbon_tab([
            ("↺", L("Ghumao", "Rotate"), self.rotate_left),
            ("✂", "Crop", self.autocrop_current),
            ("✒", "Sign", self.place_sign),
            ("🔤", "OCR text", self.copy_page_text),
            ("🗑", "Delete", self.delete_page),
            ("↩", "Undo", self.undo_delete)]), L("✏ Edit", "✏ Edit"))
        self.ribbon.addTab(_ribbon_tab([
            ("🗜", "Compress", self.compress_pdf_tool),
            ("🧩", "Merge", self.merge_pdfs),
            ("✂", "Split", self.split_pdfs),
            ("🔍", "Search", self.search_pdfs),
            ("📊", "Stats", self.show_stats_dashboard),
            ("📷", "Cards", self.business_cards)]), L("🧰 Tools", "🧰 Tools"))
        self.ribbon.addTab(_ribbon_tab([
            ("🟢", "WhatsApp", self.share_whatsapp),
            ("✉", "Email", self.share_email)]), L("📤 Share", "📤 Share"))
        self.ribbon.setVisible(bool(self._opts.get("ui_ribbon", False)))
        if self._opts.get("ui_ribbon"):
            tbwrap.hide()
        outer.addWidget(self.ribbon)
        # ---- UI #7: Status-header card (toolbar ke neeche patli smart patti) ----
        self.ui_header = QtWidgets.QWidget()
        self.ui_header.setObjectName("uiheader")
        self.ui_header.setStyleSheet(
            "#uiheader{background:#f0fdfa;border-bottom:1px solid #ccfbf1;}"
            "#uiheader QLabel{font-size:12px;color:#115e59;}")
        _hb = QtWidgets.QHBoxLayout(self.ui_header)
        _hb.setContentsMargins(10, 3, 10, 3)
        self.hdr_scanner = QtWidgets.QLabel("●")
        self.hdr_profile = QtWidgets.QLabel("")
        self.hdr_today = QtWidgets.QLabel("")
        _hb.addWidget(self.hdr_scanner)
        _hb.addWidget(self.hdr_profile)
        _hb.addStretch(1)
        _hb.addWidget(self.hdr_today)
        self.ui_header.setVisible(bool(self._opts.get("ui_header", True)))
        outer.addWidget(self.ui_header)
        # ---- UI #10: Job-chips bar (ek click me profile+folder+naming set) ----
        self.jobs_bar = QtWidgets.QWidget()
        self.jobs_bar.setObjectName("jobsbar")
        self.jobs_bar.setStyleSheet("#jobsbar{background:#fffbeb;border-bottom:1px solid #fde68a;}")
        self._jobs_lay = QtWidgets.QHBoxLayout(self.jobs_bar)
        self._jobs_lay.setContentsMargins(10, 4, 10, 4)
        self._jobs_lay.setSpacing(6)
        self.jobs_bar.setVisible(bool(self._opts.get("ui_jobs", False)))
        outer.addWidget(self.jobs_bar)
        self._rebuild_jobs_bar()
        hr = QtWidgets.QFrame(); hr.setObjectName("hr"); hr.setFrameShape(QtWidgets.QFrame.HLine); outer.addWidget(hr)

        # ---------- Body: left settings panel | thumbnails ----------
        body = QtWidgets.QHBoxLayout(); body.setContentsMargins(0, 0, 0, 0); body.setSpacing(0)
        panel = QtWidgets.QWidget(); panel.setObjectName("panel"); panel.setFixedWidth(252)
        self.left_panel = panel
        if not self._opts.get("show_left_panel", True):
            panel.hide()
        pl = QtWidgets.QVBoxLayout(panel); pl.setContentsMargins(14, 14, 14, 14); pl.setSpacing(5)

        pl.addWidget(QtWidgets.QLabel(tr("profile", self._lang)))
        prow = QtWidgets.QHBoxLayout(); prow.setSpacing(4)
        self.cmb_profile = QtWidgets.QComboBox(); self.cmb_profile.currentTextChanged.connect(self._on_profile_changed)
        prow.addWidget(self.cmb_profile, 1)
        bnew = QtWidgets.QPushButton("+"); bnew.setFixedWidth(30); bnew.clicked.connect(self._quick_new_profile); prow.addWidget(bnew)
        bedit = QtWidgets.QPushButton("✎"); bedit.setFixedWidth(30); bedit.clicked.connect(self._quick_edit_profile); prow.addWidget(bedit)
        pw = QtWidgets.QWidget(); pw.setLayout(prow); pl.addWidget(pw)

        pl.addSpacing(4); pl.addWidget(QtWidgets.QLabel("Device:"))
        self.dev_lbl = QtWidgets.QLabel("(koi device nahi)"); self.dev_lbl.setObjectName("dev"); self.dev_lbl.setWordWrap(True); pl.addWidget(self.dev_lbl)
        self.method_lbl = QtWidgets.QLabel(""); self.method_lbl.setObjectName("dev")
        pl.addWidget(self.method_lbl)

        pl.addSpacing(4); pl.addWidget(QtWidgets.QLabel("Claim No.:"))
        self.claim_edit = QtWidgets.QLineEdit(); self.claim_edit.setPlaceholderText("optional"); pl.addWidget(self.claim_edit)

        pl.addWidget(QtWidgets.QLabel("Paper source:"))
        self.cmb_source = QtWidgets.QComboBox(); self.cmb_source.addItems(["Feeder (ADF)", "Glass (Flatbed)"]); pl.addWidget(self.cmb_source)
        pl.addWidget(QtWidgets.QLabel("Scan sides:"))
        self.cmb_sides = QtWidgets.QComboBox()
        self.cmb_sides.addItems(["Single side (ek taraf)", "Both side (dono taraf)"])
        self.cmb_sides.setToolTip("Both side = kaagaz ke dono taraf scan (duplex)")
        pl.addWidget(self.cmb_sides)
        pl.addWidget(QtWidgets.QLabel("Page size:"))
        self.cmb_pagesize = QtWidgets.QComboBox(); self.cmb_pagesize.addItems(["Auto (alag-alag size khud pakde)", "A4 (210x297 mm)", "Letter", "Legal", "A5"]); pl.addWidget(self.cmb_pagesize)
        self.cmb_pagesize.setToolTip("Auto = har page ki asli size khud detect (mixed size / ID card / aadha page bhi poora). A4/Letter/Legal = fixed size.")
        pl.addWidget(QtWidgets.QLabel("Resolution:"))
        self.cmb_dpi = QtWidgets.QComboBox(); self.cmb_dpi.addItems([d + " dpi" for d in RESOLUTIONS]); self.cmb_dpi.setCurrentText("200 dpi"); pl.addWidget(self.cmb_dpi)
        pl.addWidget(QtWidgets.QLabel("Bit depth:"))
        self.cmb_depth = QtWidgets.QComboBox(); self.cmb_depth.addItems(["24-bit Colour", "Grayscale", "Black & White"]); self.cmb_depth.setCurrentText("Grayscale"); pl.addWidget(self.cmb_depth)

        pl.addStretch(1)
        # UPDATE BANNER: naya version website par aate hi yahan dikhta hai —
        # ek click me download + install + restart (chhupa rehta hai warna).
        self.update_box = QtWidgets.QPushButton()
        self.update_box.setObjectName("updatebox")
        self.update_box.setStyleSheet(
            "#updatebox{background:#f59e0b; color:#1f2937; font-weight:700;"
            " border:none; border-radius:8px; padding:10px; font-size:12px;}"
            "#updatebox:hover{background:#d97706; color:#fff;}"
            "#updatebox:disabled{background:#fbbf24; color:#6b7280;}")
        self.update_box.clicked.connect(self._sidebar_update_clicked)
        self.update_box.hide()
        pl.addWidget(self.update_box)
        self.stats_box = QtWidgets.QLabel()
        self.stats_box.setTextFormat(QtCore.Qt.RichText)
        self.stats_box.setObjectName("statsbox")
        self.stats_box.setStyleSheet(
            "#statsbox{border:1px solid #cbd5e1; border-radius:8px; padding:8px;"
            " color:#334155; font-size:12px; background:#f8fafc;}")
        self.stats_box.setToolTip("Click karo \u2014 poora Stats Dashboard khulega.\n"
                                  "Kya-kya dikhe: Settings \u2192 Sidebar stats chuno")
        self.stats_box.setCursor(QtCore.Qt.PointingHandCursor)
        self.stats_box.mousePressEvent = lambda _e: self.show_stats_dashboard()
        self._set_stats_display(None)
        pl.addWidget(self.stats_box)
        # ---- UI #9: sidebar me 7-din ka chhota graph ----
        self.side_graph = SparkBars([0] * 7)
        self.side_graph.setMinimumHeight(56)
        self.side_graph.setMaximumHeight(56)
        self.side_graph.setToolTip(self.L("Pichhle 7 din ke pages", "Pages in the last 7 days"))
        self.side_graph.setVisible(bool(self._opts.get("ui_graph", True)))
        pl.addWidget(self.side_graph)
        self.btn_scan = QtWidgets.QPushButton("▶  " + tr("scan", self._lang)); self.btn_scan.setObjectName("primary")
        self.btn_scan.setMinimumHeight(38); self.btn_scan.clicked.connect(self.do_scan); pl.addWidget(self.btn_scan)
        self.btn_scan.setToolTip("Scan shuru karo (F5)")
        self.claim_edit.setToolTip("Claim/Patient number (file ke naam me aayega)")
        self.cmb_dpi.setToolTip("Resolution: kam dpi = tez scan")
        self.cmb_depth.setToolTip("Black & White sabse tez, Colour dhima")
        self.setAcceptDrops(True)
        body.addWidget(panel)

        vline = QtWidgets.QFrame(); vline.setObjectName("hr"); vline.setFrameShape(QtWidgets.QFrame.VLine); body.addWidget(vline)

        self.list = QtWidgets.QListWidget()
        self.list.setViewMode(QtWidgets.QListView.IconMode)
        self.list.setIconSize(QtCore.QSize(self.THUMB_W, self.THUMB_H))
        self.list.setResizeMode(QtWidgets.QListView.Adjust)
        self.list.setMovement(QtWidgets.QListView.Snap)
        self.list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.list.setSpacing(10); self.list.setUniformItemSizes(True)
        self.list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self._thumb_w, self._thumb_h = self.THUMB_W, self.THUMB_H   # zoomable display size
        self._apply_thumb_zoom(self.THUMB_W)
        # empty-state hint (shown when no pages)
        self._empty_lbl = QtWidgets.QLabel(
            "\u2b07  Scan dabao ya Import se image daalo\n\n(ya kisi image ko yahan drag karo)"
            if self._lang == "hi" else
            "\u2b07  Press Scan, or Import an image\n\n(or drag an image here)",
            self.list.viewport())
        self._empty_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._empty_lbl.setStyleSheet("color:#94a3b8; font-size:15px;")
        # ---- UI #2: Start-dashboard (khaali screen par bade action-cards) ----
        self._dash = QtWidgets.QWidget(self.list.viewport())
        _dv = QtWidgets.QVBoxLayout(self._dash)
        _dv.addStretch(1)
        _dt = QtWidgets.QLabel(self.L("Kya karna hai?", "What would you like to do?"))
        _dt.setAlignment(QtCore.Qt.AlignCenter)
        _dt.setStyleSheet("color:#64748b;font-size:17px;font-weight:600;")
        _dv.addWidget(_dt)
        _dr = QtWidgets.QHBoxLayout()
        _dr.addStretch(1)

        def _dbtn(icon, text, slot):
            b = QtWidgets.QPushButton("%s\n%s" % (icon, text))
            b.setMinimumSize(118, 84)
            b.setStyleSheet("QPushButton{font-size:13px;font-weight:700;border:1px solid "
                            "#cbd5e1;border-radius:12px;background:#fff;padding:8px;}"
                            "QPushButton:hover{border-color:#0f766e;color:#0f766e;}")
            b.clicked.connect(slot)
            _dr.addWidget(b)
        _dbtn("🖨", self.L("Scan karo", "Scan"), self.do_scan)
        _dbtn("📥", "Import", self.import_images)
        _dbtn("📷", "Photo→PDF", self.import_photos)
        _dbtn("🕘", "History", self.show_history)
        _dr.addStretch(1)
        _dv.addLayout(_dr)
        self._dash_recent = QtWidgets.QLabel("")
        self._dash_recent.setAlignment(QtCore.Qt.AlignCenter)
        self._dash_recent.setStyleSheet("color:#94a3b8;font-size:11px;")
        _dv.addWidget(self._dash_recent)
        _dv.addStretch(1)
        self._dash.hide()
        # ---- UI #6: Floating Scan button (FAB) ----
        self.fab = QtWidgets.QPushButton("🖨", self)
        self.fab.setFixedSize(56, 56)
        self.fab.setToolTip("Scan (F5)")
        self.fab.setStyleSheet(
            "QPushButton{background:#0f766e;color:#fff;border:none;border-radius:28px;"
            "font-size:22px;}QPushButton:hover{background:#115e59;}")
        self.fab.clicked.connect(self.do_scan)
        self.fab.setVisible(bool(self._opts.get("ui_fab", False)))
        self.list.viewport().installEventFilter(self)
        self.list.itemDoubleClicked.connect(self._open_preview_dialog)
        self.list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.list.customContextMenuRequested.connect(self._list_context_menu)
        body.addWidget(self.list, 1)

        # ---------- Right sidebar: "Meri Files" (folder list + save-here) ----------
        self.files_panel = QtWidgets.QWidget()
        self.files_panel.setObjectName("panel")
        fp = QtWidgets.QVBoxLayout(self.files_panel)
        fp.setContentsMargins(8, 8, 8, 8)
        _hdr = QtWidgets.QLabel(self.L("📁 <b>Meri Files</b>", "📁 <b>My Files</b>"))
        _hdr.setTextFormat(QtCore.Qt.RichText)
        _hdr.setToolTip("Save folder ke folders aur documents — yahin se naya folder "
                        "banao aur scan ki PDF seedha usme save karo.")
        fp.addWidget(_hdr)
        # Search: kisi bhi folder ke andar naam se dhoondo (folder chuna ho to
        # usi ke andar, warna poore save-folder me)
        self.files_search = QtWidgets.QLineEdit()
        self.files_search.setPlaceholderText(
            self.L("🔍 Dhoondo… (chune folder ke andar)", "🔍 Search… (inside selected folder)"))
        self.files_search.setClearButtonEnabled(True)
        fp.addWidget(self.files_search)
        self.fav_bar = QtWidgets.QHBoxLayout()
        self.fav_bar.setSpacing(4)
        fp.addLayout(self.fav_bar)
        _root = self._files_root()
        self.files_model = LibraryModel(self)
        self.files_model.setRootPath(_root)
        self.files_model.setNameFilters(["*.pdf", "*.jpg", "*.jpeg", "*.png",
                                         "*.tif", "*.tiff", "*.docx", "*.xlsx"])
        self.files_model.setNameFilterDisables(False)
        self.files_tree = FilesTree(self._on_pages_dropped)
        self.files_tree.setModel(self.files_model)
        self.files_tree.setRootIndex(self.files_model.index(_root))
        for _c in (1, 2, 3):
            self.files_tree.hideColumn(_c)
        self.files_tree.setHeaderHidden(True)
        self.files_tree.setAnimated(True)
        self.files_tree.doubleClicked.connect(self._files_tree_open)
        self.files_tree.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.files_tree.customContextMenuRequested.connect(self._files_tree_menu)
        self.files_tree.selectionModel().currentChanged.connect(self._files_sel_changed)
        self.files_tree.setToolTip("Folder chuno → 'Yahan save' dabao, ya pages ko\n"
                                   "kheench kar folder par chhod do.\n"
                                   "Right-click = share/rename/delete/merge.\n"
                                   "File par double-click = kholo. Hari = aaj ki file.")
        fp.addWidget(self.files_tree, 1)
        # search ke results (search karte hi tree ki jagah dikhte hain)
        self.files_results = QtWidgets.QListWidget()
        self.files_results.itemDoubleClicked.connect(
            lambda it: it.data(QtCore.Qt.UserRole) and self._open_path(it.data(QtCore.Qt.UserRole)))
        self.files_results.hide()
        fp.addWidget(self.files_results, 1)
        self._files_search_timer = QtCore.QTimer(self)
        self._files_search_timer.setSingleShot(True)
        self._files_search_timer.setInterval(350)
        self._files_search_timer.timeout.connect(self._run_files_search)
        self.files_search.textChanged.connect(lambda _t: self._files_search_timer.start())
        self._rebuild_fav_bar()
        _row = QtWidgets.QHBoxLayout()
        _bnew = QtWidgets.QPushButton(self.L("➕ Naya folder", "➕ New folder"))
        _bnew.clicked.connect(self.new_library_folder)
        _row.addWidget(_bnew)
        fp.addLayout(_row)
        self.btn_save_here = QtWidgets.QPushButton(self.L("💾 Yahan save karo", "💾 Save here"))
        self.btn_save_here.setObjectName("primary")
        self.btn_save_here.setMinimumHeight(34)
        self.btn_save_here.clicked.connect(self.save_into_selected_folder)
        fp.addWidget(self.btn_save_here)
        self.files_panel.setFixedWidth(235)
        body.addWidget(self.files_panel)
        if not self._opts.get("show_files_panel", True):
            self.files_panel.hide()

        # ---- UI #3: Preview panel (page click → badi jhalak + quick-edit) ----
        self.preview_panel = QtWidgets.QWidget()
        self.preview_panel.setObjectName("panel")
        self.preview_panel.setFixedWidth(300)
        pv = QtWidgets.QVBoxLayout(self.preview_panel)
        pv.setContentsMargins(8, 8, 8, 8)
        self.pv_title = QtWidgets.QLabel(self.L("👁 Preview", "👁 Preview"))
        self.pv_title.setStyleSheet("font-weight:700;")
        pv.addWidget(self.pv_title)
        self.pv_img = QtWidgets.QLabel()
        self.pv_img.setAlignment(QtCore.Qt.AlignCenter)
        self.pv_img.setMinimumHeight(320)
        self.pv_img.setStyleSheet("border:1px solid #cbd5e1;border-radius:8px;background:#fff;")
        pv.addWidget(self.pv_img, 1)
        _qe = QtWidgets.QHBoxLayout()
        for _t, _tip, _fn in (("↺", "Rotate left", self.rotate_left),
                              ("↻", "Rotate right", self.rotate_right),
                              ("✂", "Auto-crop", self.autocrop_current),
                              ("✒", "Sign/Stamp", self.place_sign),
                              ("🗑", "Delete", self.delete_page)):
            _b = QtWidgets.QPushButton(_t); _b.setToolTip(_tip)
            _b.setFixedHeight(30); _b.clicked.connect(_fn)
            _qe.addWidget(_b)
        pv.addLayout(_qe)
        self.pv_info = QtWidgets.QLabel("")
        self.pv_info.setStyleSheet("color:#64748b;font-size:11px;")
        self.pv_info.setWordWrap(True)
        pv.addWidget(self.pv_info)
        self.preview_panel.setVisible(bool(self._opts.get("ui_preview", False)))
        body.addWidget(self.preview_panel)
        self.list.currentItemChanged.connect(lambda cur, prev: self._update_preview_panel())

        outer.addLayout(body, 1)

        # ---- UI #8: Kiosk overlay (bade buttons — dukaan/touch) ----
        self.kiosk = QtWidgets.QWidget(self)
        self.kiosk.setStyleSheet("background:#f4f6f8;")
        _kg = QtWidgets.QGridLayout(self.kiosk)
        _kg.setContentsMargins(30, 30, 30, 30)
        _kg.setSpacing(16)
        _kbtns = [("🖨", self.L("SCAN", "SCAN"), self.do_scan),
                  ("💾", self.L("SAVE PDF", "SAVE PDF"), self.save_pdf_all),
                  ("🟢", "WHATSAPP", self.share_whatsapp),
                  ("📷", self.L("PHOTO→PDF", "PHOTO→PDF"), self.import_photos),
                  ("🗜", "COMPRESS", self.compress_pdf_tool),
                  ("🖨️", "PRINT", self.print_pages),
                  ("🕘", "HISTORY", self.show_history),
                  ("🚪", self.L("BAND KARO", "EXIT KIOSK"), self.toggle_kiosk)]
        for i, (ic, tx, fn) in enumerate(_kbtns):
            b = QtWidgets.QPushButton("%s\n%s" % (ic, tx))
            b.setMinimumHeight(120)
            b.setStyleSheet("QPushButton{font-size:20px;font-weight:800;border:2px solid "
                            "#0f766e;border-radius:18px;background:#fff;color:#0f766e;}"
                            "QPushButton:hover{background:#0f766e;color:#fff;}")
            b.clicked.connect(fn)
            _kg.addWidget(b, i // 3, i % 3)
        self.kiosk.hide()

        # ---------- Status bar: connection ----------
        self.status = self.statusBar()
        self.lbl_conn = QtWidgets.QLabel(); self.lbl_conn.setTextFormat(QtCore.Qt.RichText)
        self._set_conn_display(None, tr("checking", self._lang))
        self.status.addWidget(self.lbl_conn)
        self.lbl_busy = QtWidgets.QLabel(); self.lbl_busy.setTextFormat(QtCore.Qt.RichText)
        self.status.addPermanentWidget(self.lbl_busy)
        self._set_busy_display("unknown")
        self._state_timer = QtCore.QTimer(self)
        self._state_timer.setInterval(6000)
        self._state_timer.timeout.connect(self._tick_scanner_state)
        self._state_timer.start()
        QtCore.QTimer.singleShot(800, self._tick_scanner_state)
        self._stats_timer = QtCore.QTimer(self)
        self._stats_timer.setInterval(90000)
        self._stats_timer.timeout.connect(lambda: self._refresh_stats("ping"))
        self._stats_timer.start()
        QtCore.QTimer.singleShot(1500, lambda: self._refresh_stats("ping"))
        # Naya version aaya ho to sidebar me banner dikhao — startup par aur
        # phir har 6 ghante (lambi chalti app bhi update dekh legi)
        QtCore.QTimer.singleShot(4000, lambda: self.check_updates(True))
        self._upd_timer = QtCore.QTimer(self)
        self._upd_timer.setInterval(6 * 3600 * 1000)
        self._upd_timer.timeout.connect(lambda: self.check_updates(True))
        self._upd_timer.start()
        # Mahine ki pehli baar kholne par "Aapka Mahina" summary
        QtCore.QTimer.singleShot(6000, self._maybe_month_wrap)
        # Kiosk mode band karke exit kiya tha to wapas usi me kholo
        if self._opts.get("ui_kiosk"):
            QtCore.QTimer.singleShot(300, lambda: (self.kiosk.setGeometry(
                self.centralWidget().rect()), self.kiosk.show(), self.kiosk.raise_()))
        # kept (hidden) for the connection feature; IP set via Settings menu
        self.ip_field = QtWidgets.QLineEdit(self._config.get("scanner_ip", "")); self.ip_field.hide()
        self.btn_check = QtWidgets.QPushButton(); self.btn_check.hide()
        self._pb_holder = None


    def _apply_style(self):
        self._apply_base_style()
        # Touch / buzurg mode: sab kuch bada — buttons, text, thumbnails
        if self._opts.get("touch_mode"):
            self.setStyleSheet(self.styleSheet() + """
                QToolButton { font-size:14px; padding:10px 12px; }
                QPushButton { font-size:15px; padding:12px 18px; }
                QMenuBar, QMenu, QLabel, QLineEdit, QComboBox, QSpinBox { font-size:14px; }
                QListWidget { font-size:14px; }
            """)
            try:
                self.list.setIconSize(QtCore.QSize(190, 250))
            except Exception:
                pass

    def _apply_base_style(self):
        if self._opts.get("theme") == "darkpro":
            # UI #4: Dark Pro — gehra premium look, teal glow
            self.setStyleSheet("""
                QMainWindow, QWidget { background:#0b1220; color:#e2e8f0; }
                QMenuBar { background:#0b1220; color:#e2e8f0; }
                QMenuBar::item:selected { background:#132036; }
                QMenu { background:#111c30; color:#e2e8f0; border:1px solid #1e3a5f; }
                QMenu::item:selected { background:#134e4a; color:#5eead4; }
                #toolbar { background:#0d1526; border-bottom:1px solid #1e3a5f; }
                QToolButton { border:1px solid transparent; border-radius:9px; padding:5px 7px; color:#94a3b8; font-size:11px; }
                QToolButton:hover { background:#132036; border-color:#2dd4bf; color:#5eead4; }
                QToolButton:checked { background:#134e4a; border-color:#2dd4bf; color:#5eead4; }
                #panel { background:#0d1526; }
                #uiheader { background:#0d1f1d; border-bottom:1px solid #134e4a; }
                #uiheader QLabel { color:#5eead4; }
                #dev { color:#64748b; font-size:12px; }
                #hr { color:#1e3a5f; }
                QLabel { color:#cbd5e1; }
                QLineEdit, QComboBox, QSpinBox, QPlainTextEdit { background:#111c30; border:1px solid #1e3a5f; border-radius:8px; padding:5px 9px; color:#e2e8f0; }
                QLineEdit:focus, QComboBox:focus { border-color:#2dd4bf; }
                QPushButton { background:#111c30; border:1px solid #1e3a5f; border-radius:8px; padding:6px 12px; color:#e2e8f0; }
                QPushButton:hover { background:#132036; border-color:#2dd4bf; }
                QPushButton#primary { background:qlineargradient(x1:0,y1:0,x2:0,y2:1,stop:0 #14b8a6,stop:1 #0f766e); border:1px solid #2dd4bf; color:#fff; font-weight:700; }
                QPushButton#primary:hover { background:#0d9488; }
                QListWidget { background:#080e1a; border:none; }
                QListWidget::item:selected { background:#134e4a; color:#e2e8f0; }
                QTreeView { background:#0d1526; color:#cbd5e1; border:none; }
                QTreeView::item:selected { background:#134e4a; color:#5eead4; }
                QScrollArea { background:#0d1526; }
                #statsbox { border:1px solid #1e3a5f; border-radius:10px; padding:8px; color:#94a3b8; background:#0d1526; }
            """)
            return
        if self._opts.get("theme") == "dark":
            self.setStyleSheet("""
                QMainWindow, QWidget { background:#0f172a; color:#e2e8f0; }
                QMenuBar { background:#0f172a; color:#e2e8f0; }
                QMenuBar::item:selected { background:#1e293b; }
                QMenu { background:#1e293b; color:#e2e8f0; border:1px solid #334155; }
                QMenu::item:selected { background:#334155; }
                #toolbar { background:#111827; }
                QToolButton { border:1px solid transparent; border-radius:8px; padding:4px 6px; color:#cbd5e1; font-size:11px; }
                QToolButton:hover { background:#1e293b; border-color:#334155; }
                QToolButton:checked { background:#134e4a; border-color:#2dd4bf; color:#5eead4; }
                #panel { background:#111827; }
                #dev { color:#94a3b8; font-size:12px; }
                #hr { color:#334155; }
                QLabel { color:#cbd5e1; }
                QLineEdit, QComboBox, QSpinBox, QPlainTextEdit { background:#1e293b; border:1px solid #475569; border-radius:6px; padding:4px 8px; color:#e2e8f0; }
                QPushButton { background:#1e293b; border:1px solid #475569; border-radius:6px; padding:6px 12px; color:#e2e8f0; }
                QPushButton:hover { background:#334155; }
                QPushButton#primary { background:#0f766e; border:1px solid #0f766e; color:#fff; font-weight:700; }
                QPushButton#primary:hover { background:#115e59; }
                QListWidget { background:#0b1220; border:none; }
                QListWidget::item:selected { background:#134e4a; color:#e2e8f0; }
                QScrollArea { background:#111827; }
            """)
        else:
            self.setStyleSheet("""
                QMainWindow, QWidget { background:#f4f6f8; color:#1f2933; }
                QMenuBar { background:#f4f6f8; }
                #toolbar { background:#ffffff; }
                QToolButton { border:1px solid transparent; border-radius:8px; padding:4px 6px; color:#334155; font-size:11px; }
                QToolButton:hover { background:#eef2f6; border-color:#e2e8f0; }
                QToolButton:checked { background:#d1eae7; border-color:#0f766e; color:#0f766e; }
                #panel { background:#ffffff; }
                #dev { color:#475569; font-size:12px; }
                #hr { color:#e2e8f0; }
                QLabel { color:#334155; }
                QLineEdit, QComboBox, QSpinBox { background:#fff; border:1px solid #cbd5e1; border-radius:6px; padding:4px 8px; }
                QPushButton { background:#fff; border:1px solid #cbd5e1; border-radius:6px; padding:6px 12px; color:#1f2933; }
                QPushButton:hover { background:#eef2f6; }
                QPushButton#primary { background:#0f766e; border:1px solid #0f766e; color:#fff; font-weight:700; }
                QPushButton#primary:hover { background:#115e59; }
                QListWidget { background:#fbfcfd; border:none; }
                QListWidget::item:selected { background:#d1eae7; color:#0f172a; }
            """)


    def _set_conn_display(self, state, message):
        colour = {True: "#16a34a", False: "#dc2626", None: "#9ca3af"}[state]
        dot = '<span style="color:%s; font-size:16px;">&#9679;</span>' % colour
        self.lbl_conn.setText("%s&nbsp; %s" % (dot, message))

    def check_connection(self):
        ip = self.ip_field.text().strip()
        self._config["scanner_ip"] = ip; save_config(self._config)
        if not ip:
            self._set_conn_display(False, "Scanner IP daalein"); return
        self._set_conn_display(None, "Check ho raha hai... (%s)" % ip)
        self.btn_check.setEnabled(False)
        self._checker = ConnectionChecker(ip)
        self._checker.result.connect(self._on_conn_result)
        self._checker.finished.connect(lambda: self.btn_check.setEnabled(True))
        self._checker.start()

    def _on_conn_result(self, ok, message):
        self._set_conn_display(bool(ok), message)

    def _set_busy_display(self, kind):
        if kind == "free":
            txt, col = "Scanner FREE (taiyar)", "#16a34a"
        elif kind == "busy":
            txt, col = "Scanner BUSY (vyast)", "#dc2626"
        else:
            txt, col = "Scanner: --", "#9ca3af"
        self.lbl_busy.setText(
            '<span style="color:%s; font-size:15px;">&#9679;</span>&nbsp;<b>%s</b>' % (col, txt))
        # UI header ki scanner-pill bhi yahi dikhaye
        try:
            self.hdr_scanner.setText(
                '<span style="color:%s;">●</span> <b>%s</b>' % (col, txt))
        except Exception:
            pass

    def _tick_scanner_state(self):
        # IMPORTANT: do NOT poll the scanner over the network here. This HP scanner
        # allows only one eSCL connection at a time, so a background status poll
        # collides with the real scan and makes it fail with HTTP 503. The indicator
        # simply reflects OUR own activity: busy only while we are scanning.
        if getattr(self, "_worker", None) is not None and self._worker.isRunning():
            self._set_busy_display("busy")
        elif getattr(self, "_scanning", False):
            self._set_busy_display("busy")
        else:
            self._set_busy_display("free")

    # ---- list ----
    def _add_item_for_path(self, path, qimg=None):
        # qimg: background worker se pehle se bana-banaya thumbnail (QImage) —
        # UI thread par bhaari decode/scale nahi karna padta.
        if qimg is not None and not qimg.isNull():
            icon = QtGui.QIcon(QtGui.QPixmap.fromImage(qimg))
        else:
            icon = QtGui.QIcon(self._make_thumb(path))
        _lbl = ("Page %d" % (self.list.count() + 1)) if self._opts.get("show_page_numbers", True) else ""
        item = QtWidgets.QListWidgetItem(icon, _lbl)
        item.setData(QtCore.Qt.UserRole, path); item.setTextAlignment(QtCore.Qt.AlignHCenter)
        self.list.addItem(item); self.list.setCurrentItem(item)
        self.list.clearSelection()  # nothing "selected" by default; user picks with Ctrl/Shift
        self._dirty = True
        self._update_status(); self._update_empty_state()

    THUMB_HI_W = 360
    THUMB_HI_H = 480

    def _make_thumb(self, path):
        # QImageReader ka scaled decode (khaaskar JPEG me) poori image decode
        # karke SmoothTransformation se sikodne se KAI GUNA tez hai — isi se
        # scan/import ke time UI ka atakna band hota hai.
        r = QtGui.QImageReader(path)
        r.setAutoTransform(True)
        sz = r.size()
        if sz.isValid() and sz.width() > 0 and sz.height() > 0:
            s = min(float(self.THUMB_HI_W) / sz.width(),
                    float(self.THUMB_HI_H) / sz.height(), 1.0)
            r.setScaledSize(QtCore.QSize(max(1, int(sz.width() * s)),
                                         max(1, int(sz.height() * s))))
        img = r.read()
        if img.isNull():
            return QtGui.QPixmap(self.THUMB_HI_W, self.THUMB_HI_H)
        return QtGui.QPixmap.fromImage(img)

    def _refresh_item(self, item):
        path = item.data(QtCore.Qt.UserRole)
        item.setIcon(QtGui.QIcon(self._make_thumb(path)))
        if item is self.list.currentItem():
            self._show_preview(item, None)

    def _ordered_paths(self):
        return [self.list.item(i).data(QtCore.Qt.UserRole) for i in range(self.list.count())]

    def _selected_paths(self):
        # selected items, kept in the on-screen (list) order
        sel = set(self.list.row(it) for it in self.list.selectedItems())
        return [self.list.item(i).data(QtCore.Qt.UserRole) for i in range(self.list.count()) if i in sel]

    def _show_preview(self, current, _prev=None):
        # NAPS2-style: no side preview pane; double-click opens a viewer.
        return


    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "fab"):
            self.fab.move(self.width() - 78, self.height() - 100)
            self.fab.raise_()


    def do_scan(self):
        method = self._opts.get("scanner_method", "twain")
        prof = self._selected_profile()
        if method == "escl":
            ip = self.ip_field.text().strip()
            self._opts["scanner_ip"] = ip
            if not ip:
                self._warn("Scanner IP set nahi hai. Settings \u2192 Scanner IP me IP daalo (jaise 192.168.1.8)."); return
        elif method == "naps2":
            if not (self._opts.get("naps2_path") or find_naps2()):
                self._warn("NAPS2 nahi mila. Settings \u2192 Scan method \u2192 naps2 me path set karo."); return
            if not self._opts.get("naps2_profile"):
                self._warn("Settings \u2192 Scan method \u2192 naps2 me NAPS2 profile ka naam set karo."); return
        else:
            if not HAS_TWAIN and method == "twain":
                self._warn("TWAIN (pytwain) install nahi hai."); return
            if prof is None:
                QtWidgets.QMessageBox.information(self, APP_NAME, "Pehli baar ek profile banayein.\n'Profiles…' → New.")
                self.open_profiles(); prof = self._selected_profile()
                if prof is None:
                    return
            if method == "twain" and not prof.get("source_name"):
                self._warn("Is profile me scanner set nahi hai. 'Profiles…' → Edit → Choose device."); return
        self._barcode_tried = False
        self._scan_count = 0
        self._progress = ScanProgressDialog(self, prof.get("source_name") or APP_NAME, self._lang)
        self._progress.cancelled.connect(self._cancel_scan)
        dpi, color, duplex = self._panel_scan_params(prof or {})
        opts = self._opts
        if self.chk_fast.isChecked():
            dpi, color = 200, "bw"     # keep the user's duplex (both-side) choice
            opts = dict(self._opts)    # lean copy: skip heavy per-page processing
            for k in ("remove_blank", "auto_crop", "deskew", "quality_enhance"):
                opts[k] = False
        opts = dict(opts)
        opts["page_size"] = self.cmb_pagesize.currentText().strip().lower()
        self._worker = ScanWorker(int(self.winId()), (prof or {}).get("source_name"),
                                  dpi, color, duplex, self._tmpdir, opts)
        self._worker.page_done.connect(self._on_page_scanned)
        self._worker.done.connect(self._on_scan_done)
        self._worker.failed.connect(self._on_scan_failed)
        self.btn_scan.setEnabled(False)
        try:
            self._conn_timer.stop()
        except Exception:
            pass
        self._scanning = True
        try:
            self._state_timer.stop()
        except Exception:
            pass
        self._worker.start(); self._progress.show()
        self._set_busy_display("busy")

    def _on_page_scanned(self, path):
        self._add_item_for_path(path)
        self._scan_count += 1
        if self._progress:
            self._progress.set_page(self._scan_count)
        if (self._opts.get("barcode_autofill") and not self._barcode_tried
                and not self.claim_edit.text().strip()):
            self._barcode_tried = True
            code = read_barcode(path)
            if code:
                self.claim_edit.setText(code)
                self.status.showMessage("Barcode se claim number mila: %s" % code, 5000)

    def _on_scan_done(self, kept, skipped):
        if kept:
            self._pstats_bump(scan_ok=1)
        if self._progress:
            self._progress.close(); self._progress = None
        self._scanning = False
        self.btn_scan.setEnabled(True)
        try:
            self._state_timer.start()
        except Exception:
            pass
        QtCore.QTimer.singleShot(2500, self._tick_scanner_state)
        try:
            self._conn_timer.start()
        except Exception:
            pass
        msg = "%d page scan ho gaye." % kept
        if skipped:
            msg += " (%d blank hataye)" % skipped
        self.status.showMessage(msg, 5000)
        self._report_scan_stat(kept)
        if self._opts.get("auto_name") and kept:
            self.auto_name_pages()
        if self._opts.get("auto_save") and kept:
            saved = self._auto_save_pdf()
            if saved and self._opts.get("batch_mode"):
                self._start_next_batch()

    def _on_scan_failed(self, msg):
        self._pstats_bump(scan_fail=1)
        self._last_error = msg
        if self._progress:
            self._progress.close(); self._progress = None
        self._scanning = False
        self.btn_scan.setEnabled(True)
        try:
            self._state_timer.start()
        except Exception:
            pass
        QtCore.QTimer.singleShot(2500, self._tick_scanner_state)
        try:
            self._conn_timer.start()
        except Exception:
            pass
        self._warn(friendly_error(msg, self._opts.get("language", "hi")))

    def _cancel_scan(self):
        try:
            if getattr(self, "_worker", None):
                self._worker.requestInterruption()
        except Exception:
            pass
        if self._progress:
            self._progress.close(); self._progress = None
        self.btn_scan.setEnabled(True)
        try:
            self._conn_timer.start()
        except Exception:
            pass

    def _start_next_batch(self):
        for p in self._ordered_paths():
            try:
                os.remove(p)
            except Exception:
                pass
        self.list.clear(); self._show_preview(None)
        text, ok = QtWidgets.QInputDialog.getText(self, "Batch mode", "Agle document ka Claim No.:")
        if ok:
            self.claim_edit.setText(text.strip())
        self._update_status()

    def import_images(self):
        start = self._config.get("last_import_dir") or self._config.get("last_save_dir") or ""
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Images / PDF chuno", start,
            "Images & PDF (*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.pdf);;PDF (*.pdf);;Images (*.png *.jpg *.jpeg *.bmp *.tif *.tiff)")
        if not files:
            return
        try:
            self._config["last_import_dir"] = os.path.dirname(files[0])
            save_config(self._config)
        except Exception:
            pass
        self._start_import(files, "normal")

    def _start_import(self, files, mode):
        """Import BACKGROUND me chalta hai — app kabhi nahi rukti/jam hoti."""
        if getattr(self, "_importer", None) is not None and self._importer.isRunning():
            self._warn("Pehla import abhi chal raha hai — thoda ruk kar dobara try karein.")
            return
        self.status.showMessage("Import ho raha hai… (app istemal karte rahiye)", 0)
        self._importer = ImportWorker(files, self._tmpdir, mode,
                                      self.THUMB_HI_W, self.THUMB_HI_H)
        self._importer.page_ready.connect(self._on_import_page)
        self._importer.done.connect(self._on_import_done)
        self._importer.start()

    def _on_import_page(self, path, qimg):
        self._add_item_for_path(path, qimg)

    def _on_import_done(self, count):
        if count:
            self.status.showMessage("%d page import ho gaye." % count, 4000)
            if tesseract_available():
                self.auto_name_pages()
        else:
            self.status.clearMessage()
            self._warn("Import nahi ho paya. PDF ho to behtar result ke liye "
                       "'PyMuPDF' install karein:\n  py -3.12-32 -m pip install PyMuPDF")

    def scan_from_camera(self):
        """Webcam/USB-camera se seedha document capture — scanner na ho to bhi."""
        if not HAS_CAMERA:
            self._warn("Camera support is build me nahi hai (PyQt5 QtMultimedia).\n"
                       "Filhaal 'Phone-photo se PDF' ya Import istemal karein.")
            return
        if not QCameraInfo.availableCameras():
            self._warn("Koi camera nahi mila. Webcam juda hai aur on hai?")
            return
        dlg = CameraDialog(self, self._tmpdir)
        if dlg.exec_() == QtWidgets.QDialog.Accepted and dlg.captured:
            for p in dlg.captured:
                self._add_item_for_path(p)
            self.status.showMessage("%d page camera se add ho gaye." % len(dlg.captured), 4000)
            if tesseract_available():
                self.auto_name_pages()

    def import_photos(self):
        """Phone se kheenchi photos ko saaf karke pages banao (shadow hatana,
        seedha karna, chhota karna) — background me, app nahi rukti."""
        start = self._config.get("last_import_dir") or ""
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Phone-photos chuno", start,
            "Photos (*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp)")
        if not files:
            return
        try:
            self._config["last_import_dir"] = os.path.dirname(files[0])
            save_config(self._config)
        except Exception:
            pass
        self._start_import(files, "photo")

    def split_id_cards(self):
        """Selected page par rakhe 2-3 ID cards ko alag-alag pages me kaato."""
        item = self._current_item_or_warn()
        if not item:
            return
        if not HAS_NUMPY:
            self._warn("Is feature ke liye numpy chahiye (pip install numpy).")
            return
        path = item.data(QtCore.Qt.UserRole)
        try:
            with Image.open(path) as im:
                img = im.convert("RGB").copy()
        except Exception as exc:
            self._warn("Page nahi khula:\n%s" % exc)
            return
        boxes = detect_content_boxes(img)
        area = img.width * img.height
        keep = []
        for (l, t, r, b) in boxes:
            if (r - l) * (b - t) >= area * 0.01 and (r - l) > 60 and (b - t) > 60:
                keep.append((max(0, l - 10), max(0, t - 10),
                             min(img.width, r + 10), min(img.height, b + 10)))
        keep.sort(key=lambda x: (x[1], x[0]))
        if not keep:
            self._warn("Is page par alag-alag cards nahi mile.")
            return
        added = 0
        for box in keep:
            try:
                card = img.crop(box)
                fd, png = tempfile.mkstemp(suffix=".png", dir=self._tmpdir)
                os.close(fd)
                card.save(png, "PNG")
                self._add_item_for_path(png)
                added += 1
            except Exception:
                pass
        if added:
            self._dirty = True
            QtWidgets.QMessageBox.information(
                self, "Ho gaya",
                "%d card/hissa alag karke aage add kar diye.\n(Original page waise hi hai — "
                "chahiye na ho to Delete kar dein.)" % added)

    def find_scanners(self):
        """Network par eSCL scanners khud dhoondo aur chun kar IP set karo."""
        self.status.showMessage("Network par scanner dhoondh rahe hain… (10-20 sec)", 0)
        self._finder = ScannerFinder()
        self._finder.found.connect(self._on_scanners_found)
        self._finder.start()

    def _on_scanners_found(self, results):
        self.status.clearMessage()
        if not results:
            self._warn("Network par koi eSCL scanner nahi mila.\n\n"
                       "Scanner ON hai aur isi WiFi/LAN par hai? IP hath se bhi "
                       "daal sakte ho: Settings → Scanner IP.")
            return
        items = ["%s  —  %s" % (ip, model) for ip, model in results]
        pick, ok = QtWidgets.QInputDialog.getItem(
            self, "Scanner mil gaye", "%d scanner mile — kaunsa use karein?" % len(results),
            items, 0, False)
        if not ok or not pick:
            return
        ip = pick.split()[0]
        self._config["scanner_ip"] = ip
        self._opts["scanner_ip"] = ip
        self._opts["scanner_method"] = "escl"
        self.ip_field.setText(ip)
        self._save_opts(); save_config(self._config)
        QtWidgets.QMessageBox.information(
            self, "Ho gaya", "Scanner set ho gaya: %s\n(Method: eSCL network scan)" % ip)

    def bundle_split_save(self):
        """Bundle scan → BLANK page ko separator maan kar alag-alag PDFs banao.
        (Scan karte waqt 'blank hatao' OFF rakhein aur documents ke beech ek
        khaali page daalein.)"""
        paths = self._ordered_paths()
        if not paths:
            self._warn(tr("scan_first", self._lang)); return
        groups, cur = [], []
        for p in paths:
            blank = False
            try:
                with Image.open(p) as im:
                    blank = is_blank_page(im.convert("RGB"))
            except Exception:
                pass
            if blank:
                if cur:
                    groups.append(cur); cur = []
            else:
                cur.append(p)
        if cur:
            groups.append(cur)
        if len(groups) <= 1:
            self._warn("Koi blank separator page nahi mila — sab ek hi document lag "
                       "raha hai.\n(Documents ke beech ek khaali page rakh kar scan "
                       "karein, aur Settings me 'blank hatao' OFF ho.)")
            return
        folder = QtWidgets.QFileDialog.getExistingDirectory(
            self, "PDFs kis folder me banayein?", self._opts.get("save_folder", ""))
        if not folder:
            return
        os.makedirs(folder, exist_ok=True)
        made = []
        for i, grp in enumerate(groups, 1):
            it0 = None
            for r in range(self.list.count()):
                if self.list.item(r).data(QtCore.Qt.UserRole) == grp[0]:
                    it0 = self.list.item(r); break
            name = (it0.data(TITLE_ROLE) if it0 else "") or ("Document_%d" % i)
            out = os.path.join(folder, sanitize(underscore_name(name)) + ".pdf")
            n = 2
            while os.path.exists(out):
                out = os.path.join(folder, "%s_%d.pdf" % (sanitize(underscore_name(name)), n)); n += 1
            try:
                self._pages_as_pdf(grp, out)
                made.append(os.path.basename(out))
            except Exception:
                pass
        QtWidgets.QMessageBox.information(
            self, "Ho gaya", "%d alag PDF ban gayi:\n\n%s" % (len(made), "\n".join(made[:15])))

    def split_book_page(self):
        """Khuli kitab ka scan → do alag pages (baayan + daayan)."""
        items = self.list.selectedItems() or ([self.list.currentItem()] if self.list.currentItem() else [])
        if not items:
            self._warn("Pehle koi page select karein."); return
        added = 0
        for it in items:
            path = it.data(QtCore.Qt.UserRole)
            try:
                with Image.open(path) as im:
                    img = im.convert("RGB").copy()
                mid = img.width // 2
                for half in (img.crop((0, 0, mid, img.height)),
                             img.crop((mid, 0, img.width, img.height))):
                    fd, png = tempfile.mkstemp(suffix=".png", dir=self._tmpdir)
                    os.close(fd); half.save(png, "PNG")
                    self._add_item_for_path(png); added += 1
            except Exception:
                pass
        if added:
            self._dirty = True
            self.status.showMessage("%d pages ban gaye (original ko delete kar sakte ho)." % added, 5000)

    def business_cards(self):
        """Visiting cards ke scan se contacts nikaalo: har card alag + naam/phone/email
        padh kar .vcf (contact file) aur contacts.xlsx me save."""
        item = self._current_item_or_warn()
        if not item:
            return
        if not tesseract_available():
            self._warn("Iske liye Tesseract OCR chahiye."); return
        path = item.data(QtCore.Qt.UserRole)
        with Image.open(path) as im:
            img = im.convert("RGB").copy()
        boxes = detect_content_boxes(img)
        area = img.width * img.height
        cards = [b for b in boxes if (b[2]-b[0])*(b[3]-b[1]) >= area*0.01
                 and (b[2]-b[0]) > 80 and (b[3]-b[1]) > 50]
        if not cards:
            cards = [(0, 0, img.width, img.height)]
        folder = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Contacts kahan save karein?", self._opts.get("save_folder", ""))
        if not folder:
            return
        os.makedirs(folder, exist_ok=True)
        got = 0
        rows = []
        for i, b in enumerate(sorted(cards, key=lambda x: (x[1], x[0])), 1):
            card = img.crop(b)
            try:
                text = pytesseract.image_to_string(card, lang="eng+hin")
            except Exception:
                text = ""
            lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
            name = lines[0] if lines else ("Card_%d" % i)
            phones = re.findall(r"(?:\+?\d[\d\s\-]{8,14}\d)", text or "")
            emails = re.findall(r"[\w.+-]+@[\w-]+\.[\w.]+", text or "")
            phone = (phones[0].replace(" ", "").replace("-", "") if phones else "")
            email = emails[0] if emails else ""
            vcf = os.path.join(folder, sanitize(underscore_name(name))[:40] + ".vcf")
            try:
                with open(vcf, "w", encoding="utf-8") as fh:
                    fh.write("BEGIN:VCARD\nVERSION:3.0\nFN:%s\n" % name)
                    if phone:
                        fh.write("TEL;TYPE=CELL:%s\n" % phone)
                    if email:
                        fh.write("EMAIL:%s\n" % email)
                    fh.write("END:VCARD\n")
                got += 1
            except Exception:
                pass
            rows.append((name, phone, email))
            fd, png = tempfile.mkstemp(suffix=".png", dir=self._tmpdir)
            os.close(fd); card.save(png, "PNG")
            self._add_item_for_path(png)
        if HAS_XLSX and rows:
            try:
                xp = os.path.join(folder, "contacts.xlsx")
                wb = openpyxl.load_workbook(xp) if os.path.exists(xp) else openpyxl.Workbook()
                ws = wb.active
                if ws.max_row == 1 and not ws.cell(1, 1).value:
                    ws.append(["Naam", "Phone", "Email", "Date"])
                for nm, ph, em in rows:
                    ws.append([nm, ph, em, datetime.datetime.now().strftime("%Y-%m-%d")])
                wb.save(xp)
            except Exception:
                pass
        QtWidgets.QMessageBox.information(
            self, "Ho gaya",
            "%d card mile.\n%d contact (.vcf) files bani + contacts.xlsx me entry.\n\n"
            "(.vcf ko phone me bhejo — kholte hi contact save ho jata hai.)" % (len(rows), got))

    def restore_photo_current(self):
        """Selected page/photo ko sudharo (purani feeki photo wala mode)."""
        items = self.list.selectedItems() or ([self.list.currentItem()] if self.list.currentItem() else [])
        if not items:
            self._warn("Pehle koi page select karein."); return
        for it in items:
            path = it.data(QtCore.Qt.UserRole)
            try:
                with Image.open(path) as im:
                    out = restore_photo(im)
                out.save(path, "PNG")
                self._refresh_item(it)
            except Exception:
                pass
        self._dirty = True
        self.status.showMessage("Photo sudhar di gayi.", 4000)

    def show_history(self):
        """Save folder ki saari PDFs — nayi se purani, search ke saath."""
        root = self._opts.get("save_folder", os.path.expanduser("~"))
        files = []
        try:
            for dp, _dn, fn in os.walk(root):
                for f in fn:
                    if f.lower().endswith(".pdf"):
                        p = os.path.join(dp, f)
                        try:
                            st = os.stat(p)
                            files.append((st.st_mtime, st.st_size, p))
                        except Exception:
                            pass
        except Exception:
            pass
        files.sort(reverse=True)
        files = files[:300]
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Scan History (%d files)" % len(files))
        dlg.resize(720, 480)
        v = QtWidgets.QVBoxLayout(dlg)
        ed = QtWidgets.QLineEdit(); ed.setPlaceholderText("Naam se filter karein…")
        v.addWidget(ed)
        tree = QtWidgets.QTreeWidget()
        tree.setHeaderLabels(["File", "Kab", "Size"])
        tree.setRootIsDecorated(False)
        tree.setColumnWidth(0, 380); tree.setColumnWidth(1, 150)
        for mt, sz, p in files:
            when = datetime.datetime.fromtimestamp(mt).strftime("%d-%m-%Y %H:%M")
            kb = sz / 1024.0
            pretty = ("%.0f KB" % kb) if kb < 1024 else ("%.1f MB" % (kb / 1024))
            it = QtWidgets.QTreeWidgetItem([os.path.basename(p), when, pretty])
            it.setData(0, QtCore.Qt.UserRole, p)
            tree.addTopLevelItem(it)
        v.addWidget(tree, 1)

        def _filter(txt):
            t = txt.strip().lower()
            for i in range(tree.topLevelItemCount()):
                it = tree.topLevelItem(i)
                it.setHidden(bool(t) and t not in it.text(0).lower())
        ed.textChanged.connect(_filter)
        tree.itemDoubleClicked.connect(
            lambda it, col: self._open_path(it.data(0, QtCore.Qt.UserRole)))
        row = QtWidgets.QHBoxLayout()
        b1 = QtWidgets.QPushButton("Kholo")
        b1.clicked.connect(lambda: tree.currentItem() and self._open_path(
            tree.currentItem().data(0, QtCore.Qt.UserRole)))
        b2 = QtWidgets.QPushButton("Folder kholo")
        b2.clicked.connect(lambda: tree.currentItem() and self._reveal_in_explorer(
            tree.currentItem().data(0, QtCore.Qt.UserRole)))
        bc = QtWidgets.QPushButton("Band karo"); bc.clicked.connect(dlg.accept)
        row.addWidget(b1); row.addWidget(b2); row.addStretch(1); row.addWidget(bc)
        v.addLayout(row)
        dlg.exec_()

    # ---- Personal stats (sab kuch aapke PC par — kahin nahi jaata) ----
    MILESTONES = [10, 50, 100, 250, 500, 1000, 2500, 5000, 10000]

    def _pstats(self):
        if not hasattr(self, "_pstats_cache"):
            try:
                with open(PSTATS_PATH, "r", encoding="utf-8") as fh:
                    self._pstats_cache = json.load(fh)
            except Exception:
                self._pstats_cache = {}
        return self._pstats_cache

    def _pstats_save(self):
        try:
            with open(PSTATS_PATH, "w", encoding="utf-8") as fh:
                json.dump(self._pstats(), fh)
        except Exception:
            pass

    def _pstats_bump(self, pages=0, pdfs=0, shared=0, ocr_named=0,
                     scan_ok=0, scan_fail=0, saved_bytes=0, doc_type=None):
        st = self._pstats()
        day = datetime.datetime.now().strftime("%Y-%m-%d")
        d = st.setdefault("days", {}).setdefault(day, {})
        for k, v in (("pages", pages), ("pdfs", pdfs), ("shared", shared)):
            if v:
                d[k] = d.get(k, 0) + v
        t = st.setdefault("totals", {})
        for k, v in (("pages", pages), ("pdfs", pdfs), ("shared", shared),
                     ("ocr_named", ocr_named), ("scan_ok", scan_ok),
                     ("scan_fail", scan_fail), ("saved_bytes", saved_bytes)):
            if v:
                t[k] = t.get(k, 0) + v
        if pages:
            hh = st.setdefault("hours", {})
            h = datetime.datetime.now().strftime("%H")
            hh[h] = hh.get(h, 0) + pages
        if doc_type:
            ty = st.setdefault("types", {})
            key = str(doc_type).strip()[:24] or "Anya"
            ty[key] = ty.get(key, 0) + 1
        self._pstats_save()
        if pdfs:
            self._check_milestones()
        self._update_sidebar_stats()

    def _pstats_sum(self, days=7, key="pages"):
        st = self._pstats().get("days", {})
        now = datetime.datetime.now()
        total = 0
        for i in range(days):
            k = (now - datetime.timedelta(days=i)).strftime("%Y-%m-%d")
            total += (st.get(k) or {}).get(key, 0)
        return total

    def _pstats_streak(self):
        st = self._pstats().get("days", {})
        now = datetime.datetime.now()
        streak = 0
        i = 0
        # aaj kaam na hua ho to kal se ginti shuru (streak abhi tuta nahi)
        if not (st.get(now.strftime("%Y-%m-%d")) or {}).get("pages", 0):
            i = 1
        while True:
            k = (now - datetime.timedelta(days=i)).strftime("%Y-%m-%d")
            if (st.get(k) or {}).get("pages", 0) > 0:
                streak += 1
                i += 1
            else:
                break
        return streak

    def _pstats_best_day(self):
        days = self._pstats().get("days", {})
        if not days:
            return ("—", 0)
        k, v = max(days.items(), key=lambda kv: kv[1].get("pages", 0))
        return (k, v.get("pages", 0))

    def _check_milestones(self):
        st = self._pstats()
        t = st.get("totals", {})
        shown = st.setdefault("milestones_shown", [])
        for kind, label in (("pdfs", "PDFs"), ("pages", "pages")):
            val = t.get(kind, 0)
            for m in self.MILESTONES:
                key = "%s_%d" % (kind, m)
                if val >= m and key not in shown:
                    shown.append(key)
                    self._pstats_save()
                    self.status.showMessage(
                        "🏆 Badhai ho! Aapke %s %s poore ho gaye!" % ("{:,}".format(m), label), 9000)

    def _maybe_month_wrap(self):
        """Mahine ki 1 tarikh ko pichhle mahine ka 'Aapka Mahina' summary."""
        st = self._pstats()
        now = datetime.datetime.now()
        this_month = now.strftime("%Y-%m")
        if st.get("month_shown") == this_month:
            return
        prev = (now.replace(day=1) - datetime.timedelta(days=1)).strftime("%Y-%m")
        days = {k: v for k, v in (st.get("days") or {}).items() if k.startswith(prev)}
        st["month_shown"] = this_month
        self._pstats_save()
        if not days:
            return
        pages = sum(d.get("pages", 0) for d in days.values())
        pdfs = sum(d.get("pdfs", 0) for d in days.values())
        busy_k, busy_v = max(days.items(), key=lambda kv: kv[1].get("pages", 0))
        QtWidgets.QMessageBox.information(
            self, "📊 Aapka mahina (%s)" % prev,
            "Pichhle mahine aapne:\n\n📄 %s pages scan kiye\n🗂 %s PDFs banayi\n"
            "🔥 Sabse busy din: %s (%d pages)\n\nShaandaar! 🎉"
            % ("{:,}".format(pages), "{:,}".format(pdfs), busy_k, busy_v))

    # ---- Sidebar stats (aap chunte ho kya-kya dikhe) ----
    def _sidebar_stat_items(self):
        """(key, label, value) — sidebar aur chooser dono iski list se chalte hain."""
        w = getattr(self, "_world_stats", {}) or {}
        st = self._pstats()
        t = st.get("totals", {})
        day = (st.get("days") or {}).get(datetime.datetime.now().strftime("%Y-%m-%d"), {})
        L = self.L
        return [
            ("world_total", L("🌍 Total scans (world)", "🌍 Total scans (world)"), w.get("total")),
            ("world_today", L("📅 Aaj (world)", "📅 Today (world)"), w.get("today")),
            ("world_online", L("🟢 Abhi online", "🟢 Online now"), w.get("online")),
            ("world_users", L("👥 Kul users (world)", "👥 Total users (world)"), w.get("users")),
            ("my_today", L("📄 Mere aaj ke pages", "📄 My pages today"), day.get("pages", 0)),
            ("my_today_pdfs", L("🗂 Meri aaj ki PDFs", "🗂 My PDFs today"), day.get("pdfs", 0)),
            ("my_week", L("🗓 Is hafte (pages)", "🗓 This week (pages)"), self._pstats_sum(7, "pages")),
            ("my_total", L("📚 Mere kul pages", "📚 My total pages"), t.get("pages", 0)),
            ("my_pdfs", L("🗂 Meri kul PDFs", "🗂 My total PDFs"), t.get("pdfs", 0)),
            ("streak", L("🔥 Streak (din)", "🔥 Streak (days)"), self._pstats_streak()),
            ("shared", L("📤 Share ki hui", "📤 Shared"), t.get("shared", 0)),
            ("saved_mb", L("🗜 Compress se bachaya (MB)", "🗜 Saved by compress (MB)"),
             int(t.get("saved_bytes", 0) / 1048576)),
        ]

    DEFAULT_SIDEBAR_STATS = ["world_total", "world_today", "world_online",
                             "my_today", "streak"]

    def _update_sidebar_stats(self):
        try:
            sel = self._opts.get("sidebar_stats") or self.DEFAULT_SIDEBAR_STATS
            items = {k: (lbl, val) for k, lbl, val in self._sidebar_stat_items()}
            lines = ['<b>📊 ApneScan</b> <span style="color:#94a3b8;font-size:10px;">'
                     '(click = dashboard)</span>']
            for k in sel:
                if k in items:
                    lbl, val = items[k]
                    v = "…" if val is None else ("{:,}".format(val) if isinstance(val, int) else str(val))
                    lines.append("%s: <b>%s</b>" % (lbl, v))
            self.stats_box.setText("<br>".join(lines))
            # sidebar graph + header ke numbers bhi taaza karo
            if hasattr(self, "side_graph"):
                days = self._pstats().get("days", {})
                now = datetime.datetime.now()
                vals = [(days.get((now - datetime.timedelta(days=i)).strftime("%Y-%m-%d")) or {}).get("pages", 0)
                        for i in range(6, -1, -1)]
                self.side_graph.set_values(vals)
            if hasattr(self, "hdr_today"):
                d = (self._pstats().get("days") or {}).get(now.strftime("%Y-%m-%d"), {}) if hasattr(self, "side_graph") else {}
                self.hdr_today.setText(self.L(
                    "Aaj: <b>%d pages · %d PDFs</b> 🔥%d" %
                    (d.get("pages", 0), d.get("pdfs", 0), self._pstats_streak()),
                    "Today: <b>%d pages · %d PDFs</b> 🔥%d" %
                    (d.get("pages", 0), d.get("pdfs", 0), self._pstats_streak())))
            if hasattr(self, "hdr_profile"):
                try:
                    prof = self._selected_profile()
                    self.hdr_profile.setText("〔 %s 〕" % (prof.get("name") if prof else "—"))
                except Exception:
                    pass
        except Exception:
            pass

    def choose_sidebar_stats(self):
        """Aap khud chuno sidebar me kaun-kaun si stats dikhein."""
        items = self._sidebar_stat_items()
        sel = set(self._opts.get("sidebar_stats") or self.DEFAULT_SIDEBAR_STATS)
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Sidebar stats chuno")
        v = QtWidgets.QVBoxLayout(dlg)
        v.addWidget(QtWidgets.QLabel("Tick karo jo sidebar me dikhana hai:"))
        lw = QtWidgets.QListWidget()
        for k, lbl, _val in items:
            it = QtWidgets.QListWidgetItem(lbl)
            it.setData(QtCore.Qt.UserRole, k)
            it.setFlags(it.flags() | QtCore.Qt.ItemIsUserCheckable)
            it.setCheckState(QtCore.Qt.Checked if k in sel else QtCore.Qt.Unchecked)
            lw.addItem(it)
        v.addWidget(lw, 1)
        bb = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return
        chosen = [lw.item(i).data(QtCore.Qt.UserRole) for i in range(lw.count())
                  if lw.item(i).checkState() == QtCore.Qt.Checked]
        self._opts["sidebar_stats"] = chosen
        self._save_opts()
        self._update_sidebar_stats()

    def _on_world_stats(self, data):
        try:
            self._world_stats = dict(data or {})
        except Exception:
            self._world_stats = {}
        self._update_sidebar_stats()

    # ---- Stats Dashboard ----
    def show_stats_dashboard(self):
        st = self._pstats()
        t = st.get("totals", {})
        w = getattr(self, "_world_stats", {}) or {}
        now = datetime.datetime.now()
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("📊 Stats Dashboard")
        dlg.resize(560, 560)
        v = QtWidgets.QVBoxLayout(dlg)
        tabs = QtWidgets.QTabWidget()
        v.addWidget(tabs, 1)

        # --- Meri Stats tab ---
        p1 = QtWidgets.QWidget()
        l1 = QtWidgets.QVBoxLayout(p1)
        days = st.get("days", {})
        last7, labels = [], []
        for i in range(6, -1, -1):
            dt = now - datetime.timedelta(days=i)
            last7.append((days.get(dt.strftime("%Y-%m-%d")) or {}).get("pages", 0))
            labels.append(["So", "Mo", "Tu", "We", "Th", "Fr", "Sa"][int(dt.strftime("%w"))])
        l1.addWidget(QtWidgets.QLabel("<b>Pichhle 7 din (pages):</b>"))
        l1.addWidget(SparkBars(last7, labels))
        best_k, best_v = self._pstats_best_day()
        ok, fail = t.get("scan_ok", 0), t.get("scan_fail", 0)
        rate = ("%.0f%%" % (100.0 * ok / (ok + fail))) if (ok + fail) else "—"
        types = sorted((st.get("types") or {}).items(), key=lambda kv: -kv[1])[:5]
        tys = "<br>".join("&nbsp;&nbsp;%s: <b>%d</b>" % (k, n) for k, n in types) or "&nbsp;&nbsp;(abhi nahi)"
        hours = st.get("hours") or {}
        peak_h = max(hours.items(), key=lambda kv: kv[1])[0] + ":00 baje" if hours else "—"
        info = QtWidgets.QLabel(
            "📄 Aaj: <b>%d</b> pages &nbsp;|&nbsp; 🗓 Hafta: <b>%d</b> &nbsp;|&nbsp; "
            "Mahina: <b>%d</b><br>"
            "📚 Kul: <b>%s pages</b>, <b>%s PDFs</b><br>"
            "🔥 Streak: <b>%d din</b> &nbsp;|&nbsp; 🏅 Best din: <b>%s (%d)</b><br>"
            "📤 Share: <b>%d</b> &nbsp;|&nbsp; 🗜 Bachaya: <b>%.1f MB</b><br>"
            "🩺 Scan success: <b>%s</b> &nbsp;|&nbsp; 🔤 OCR-naam mile: <b>%d</b><br>"
            "⏰ Sabse busy waqt: <b>%s</b><br><br><b>Document types (top 5):</b><br>%s"
            % (self._pstats_sum(1), self._pstats_sum(7), self._pstats_sum(30),
               "{:,}".format(t.get("pages", 0)), "{:,}".format(t.get("pdfs", 0)),
               self._pstats_streak(), best_k, best_v,
               t.get("shared", 0), t.get("saved_bytes", 0) / 1048576.0,
               rate, t.get("ocr_named", 0), peak_h, tys))
        info.setTextFormat(QtCore.Qt.RichText)
        info.setWordWrap(True)
        l1.addWidget(info)
        l1.addStretch(1)
        bexp = QtWidgets.QPushButton("📥 Excel me export karo")
        bexp.clicked.connect(self._export_pstats_excel)
        l1.addWidget(bexp)
        tabs.addTab(p1, "🙋 Meri Stats")

        # --- Worldwide tab ---
        p2 = QtWidgets.QWidget()
        l2 = QtWidgets.QVBoxLayout(p2)
        week = w.get("week") or []
        try:
            wvals = [int(x[1]) for x in week][-7:]
            wlabs = [str(x[0])[-2:] for x in week][-7:]
        except Exception:
            wvals, wlabs = [], []
        if wvals:
            l2.addWidget(QtWidgets.QLabel("<b>Duniya bhar me — pichhle 7 din:</b>"))
            l2.addWidget(SparkBars(wvals, wlabs, "#2563eb"))
        def _fmt(x):
            return "…" if x is None else ("{:,}".format(int(x)) if str(x).isdigit() or isinstance(x, int) else str(x))
        vers = w.get("versions") or {}
        vtxt = ", ".join("v%s: %s" % (k, n) for k, n in
                         sorted(vers.items(), key=lambda kv: -int(kv[1]))[:4]) or "—"
        ctry = w.get("countries") or {}
        ctxt = ", ".join("%s: %s" % (k, n) for k, n in
                         sorted(ctry.items(), key=lambda kv: -int(kv[1]))[:5]) or "—"
        info2 = QtWidgets.QLabel(
            "🌍 Total scans: <b>%s</b><br>📅 Aaj: <b>%s</b> &nbsp;|&nbsp; "
            "🟢 Abhi online: <b>%s</b> &nbsp;|&nbsp; 📈 Aaj ka peak: <b>%s</b><br>"
            "👥 Kul users (ab tak): <b>%s</b><br>⏱ Is ghante ke scans: <b>%s</b><br><br>"
            "<b>Versions:</b> %s<br><b>Desh:</b> %s<br><br>"
            "<span style='color:#64748b;'>Privacy: server par sirf GINTI jaati hai — "
            "kabhi koi document, naam ya file nahi.</span>"
            % (_fmt(w.get("total")), _fmt(w.get("today")), _fmt(w.get("online")),
               _fmt(w.get("peak")), _fmt(w.get("users")), _fmt(w.get("hour")),
               vtxt, ctxt))
        info2.setTextFormat(QtCore.Qt.RichText)
        info2.setWordWrap(True)
        l2.addWidget(info2)
        l2.addStretch(1)
        tabs.addTab(p2, "🌍 Worldwide")

        bcl = QtWidgets.QPushButton("Band karo")
        bcl.clicked.connect(dlg.accept)
        v.addWidget(bcl)
        self._refresh_stats("ping")   # taaza worldwide numbers
        dlg.exec_()

    def _export_pstats_excel(self):
        if not HAS_XLSX:
            self._warn("openpyxl install nahi hai."); return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Stats export", os.path.join(os.path.expanduser("~"), "apnescan_stats.xlsx"),
            "Excel (*.xlsx)")
        if not out:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Date", "Pages", "PDFs", "Shared"])
            for k in sorted(self._pstats().get("days", {})):
                d = self._pstats()["days"][k]
                ws.append([k, d.get("pages", 0), d.get("pdfs", 0), d.get("shared", 0)])
            wb.save(out)
        except Exception as exc:
            self._warn("Export fail: %s" % exc); return
        QtWidgets.QMessageBox.information(self, "Ho gaya", "Stats export:\n%s" % out)

    # ---- "Meri Files" right panel ----
    def _files_root(self):
        root = (self._opts.get("save_folder")
                or os.path.join(os.path.expanduser("~"), "Documents", "NobleScans"))
        try:
            os.makedirs(root, exist_ok=True)
        except Exception:
            pass
        return root

    def _refresh_files_root(self):
        try:
            root = self._files_root()
            self.files_model.setRootPath(root)
            self.files_tree.setRootIndex(self.files_model.index(root))
        except Exception:
            pass

    def _files_tree_open(self, index):
        try:
            path = self.files_model.filePath(index)
        except Exception:
            return
        if path and os.path.isfile(path):
            self._open_path(path)

    def _selected_library_folder(self):
        try:
            idx = self.files_tree.currentIndex()
            if idx.isValid():
                p = self.files_model.filePath(idx)
                return p if os.path.isdir(p) else os.path.dirname(p)
        except Exception:
            pass
        return self._files_root()

    def new_library_folder(self):
        base = self._selected_library_folder()
        name, ok = QtWidgets.QInputDialog.getText(
            self, "Naya folder",
            "Folder ka naam:\n(banega: '%s' ke andar)" % (os.path.basename(base) or "Meri Files"))
        if not ok or not name.strip():
            return
        p = os.path.join(base, sanitize(underscore_name(name.strip())))
        try:
            os.makedirs(p, exist_ok=True)
        except Exception as exc:
            self._warn("Folder nahi bana:\n%s" % exc)
            return
        try:
            self.files_tree.setCurrentIndex(self.files_model.index(p))
            self.files_tree.expand(self.files_model.index(base))
        except Exception:
            pass
        self.status.showMessage("Folder ban gaya: %s" % os.path.basename(p), 4000)

    def save_into_selected_folder(self):
        """Panel me chune folder me SEEDHA save — bas naam confirm karo, Enter."""
        paths = self._ordered_paths()
        if not paths:
            self._warn(tr("scan_first", self._lang))
            return
        self._save_pages_to_folder(self._selected_library_folder(), paths, ask_name=True)

    def _save_pages_to_folder(self, folder, paths, ask_name=True):
        if not paths:
            self._warn(tr("scan_first", self._lang))
            return
        default = os.path.basename(self._build_filename(".pdf", paths=paths))
        if default.lower().endswith(".pdf"):
            default = default[:-4]
        if ask_name:
            name, ok = QtWidgets.QInputDialog.getText(
                self, "Save karein",
                "File ka naam:   (folder: %s)" % (os.path.basename(folder) or folder),
                text=default)
            if not ok or not name.strip():
                return
            default = name.strip()
        base = sanitize(underscore_name(default))
        out = os.path.join(folder, base + ".pdf")
        n = 2
        while os.path.exists(out):          # duplicate par khud numbering
            out = os.path.join(folder, "%s_%d.pdf" % (base, n))
            n += 1
        if not self._validate_claim_ok() or not self._duplicate_ok():
            return
        try:
            self._pages_as_pdf(paths, out)
        except Exception:
            self._warn("Save fail:\n%s" % traceback.format_exc())
            return
        self._remember_save_dir(out)
        self._remember_doc_name(out)
        self._record_save(out, len(paths))
        self._dirty = False
        self._after_save_action(out)
        try:
            self.files_tree.setCurrentIndex(self.files_model.index(out))
        except Exception:
            pass
        self.status.showMessage("✔ Save ho gayi: %s" % out, 7000)

    def _on_pages_dropped(self, idx):
        """Pages ko folder par DROP karo = turant wahan save (auto naam se).
        Jo pages select hain wahi; kuch select na ho to saare."""
        try:
            p = self.files_model.filePath(idx) if idx.isValid() else self._files_root()
        except Exception:
            p = self._files_root()
        folder = p if os.path.isdir(p) else os.path.dirname(p)
        paths = self._selected_paths() or self._ordered_paths()
        self._save_pages_to_folder(folder, paths, ask_name=False)

    def _files_sel_changed(self, cur, _prev):
        # Folder chunte hi status me uska hisaab: kitni files, kitni jagah
        try:
            p = self.files_model.filePath(cur)
            if p and os.path.isdir(p):
                files = [f for f in os.listdir(p)
                         if os.path.isfile(os.path.join(p, f))]
                sz = sum(os.path.getsize(os.path.join(p, f)) for f in files)
                self.status.showMessage(
                    "📁 %s — %d files, %.1f MB" %
                    (os.path.basename(p) or p, len(files), sz / 1048576.0), 4000)
        except Exception:
            pass

    def _files_tree_menu(self, pos):
        idx = self.files_tree.indexAt(pos)
        menu = QtWidgets.QMenu(self)
        if idx.isValid():
            path = self.files_model.filePath(idx)
            if os.path.isdir(path):
                menu.addAction("💾 Yahan save karo",
                               lambda: self._save_pages_to_folder(path, self._ordered_paths()))
                menu.addAction("➕ Naya folder isme…", lambda: self._new_folder_in(path))
                menu.addAction("🧩 Is folder ki saari PDFs → ek PDF…",
                               lambda: self._merge_folder_pdfs(path))
                menu.addAction("📂 Explorer me kholo", lambda: self._open_path(path))
                favs = self._opts.get("fav_folders") or []
                menu.addAction("⭐ Favourite hatao" if path in favs else "⭐ Favourite banao",
                               lambda: self._toggle_fav(path))
            else:
                menu.addAction("📖 Kholo", lambda: self._open_path(path))
                if path.lower().endswith(".pdf"):
                    menu.addAction("🟢 WhatsApp par bhejo", lambda: self.share_whatsapp(path))
                    menu.addAction("✉ Email se bhejo", lambda: self.share_email(path))
                    menu.addAction("🗜 Chhota karo (compress)…",
                                   lambda: self.compress_pdf_tool(path))
                    menu.addAction("🏷 Tag lagao…", lambda: self.tag_pdf(path))
                menu.addAction("✏ Naam badlo…", lambda: self._rename_library_file(path))
                menu.addAction("🗑 Delete…", lambda: self._delete_library_file(path))
        else:
            menu.addAction("➕ Naya folder", self.new_library_folder)
        menu.exec_(self.files_tree.viewport().mapToGlobal(pos))

    def _new_folder_in(self, base):
        name, ok = QtWidgets.QInputDialog.getText(
            self, "Naya folder",
            "Folder ka naam:\n(banega: '%s' ke andar)" % (os.path.basename(base) or "Meri Files"))
        if not ok or not name.strip():
            return
        p = os.path.join(base, sanitize(underscore_name(name.strip())))
        try:
            os.makedirs(p, exist_ok=True)
        except Exception as exc:
            self._warn("Folder nahi bana:\n%s" % exc)
            return
        try:
            self.files_tree.expand(self.files_model.index(base))
            self.files_tree.setCurrentIndex(self.files_model.index(p))
        except Exception:
            pass
        self.status.showMessage("Folder ban gaya: %s" % os.path.basename(p), 4000)

    def _toggle_fav(self, path):
        favs = self._opts.setdefault("fav_folders", [])
        if path in favs:
            favs.remove(path)
        else:
            favs.append(path)
            while len(favs) > 4:
                favs.pop(0)
        self._save_opts()
        self._rebuild_fav_bar()

    def _rebuild_fav_bar(self):
        while self.fav_bar.count():
            it = self.fav_bar.takeAt(0)
            w = it.widget()
            if w is not None:
                w.deleteLater()
        for p in (self._opts.get("fav_folders") or []):
            if not os.path.isdir(p):
                continue
            b = QtWidgets.QPushButton("⭐ " + (os.path.basename(p) or p)[:12])
            b.setToolTip(p)
            b.setStyleSheet("padding:3px 8px; font-size:11px;")
            b.clicked.connect(lambda _c=False, pp=p: self._jump_to_folder(pp))
            self.fav_bar.addWidget(b)
        self.fav_bar.addStretch(1)

    def _jump_to_folder(self, p):
        try:
            self.files_search.clear()
            idx = self.files_model.index(p)
            self.files_tree.setCurrentIndex(idx)
            self.files_tree.expand(idx)
            self.files_tree.scrollTo(idx)
        except Exception:
            pass

    def _run_files_search(self):
        """Panel ki search: chune folder ke andar (ya poore save-folder me)
        naam se file dhoondo — background me, turant results."""
        q = self.files_search.text().strip().lower()
        if len(q) < 2:
            self.files_results.hide()
            self.files_tree.show()
            return
        scope = self._selected_library_folder()

        def job():
            hits = []
            for dp, _dn, fn in os.walk(scope):
                for f in fn:
                    if q in f.lower():
                        hits.append(os.path.join(dp, f))
                        if len(hits) >= 400:
                            return hits
            return hits

        def done(res):
            if isinstance(res, Exception):
                return
            if self.files_search.text().strip().lower() != q:
                return                      # tab tak nayi search shuru ho gayi
            self.files_results.clear()
            for p in res:
                rel = os.path.dirname(os.path.relpath(p, scope))
                label = "📄 " + os.path.basename(p)
                if rel and rel != ".":
                    label += "   (%s)" % rel
                it = QtWidgets.QListWidgetItem(label)
                it.setToolTip(p)
                it.setData(QtCore.Qt.UserRole, p)
                self.files_results.addItem(it)
            if not res:
                it = QtWidgets.QListWidgetItem("(kuch nahi mila)")
                it.setFlags(QtCore.Qt.NoItemFlags)
                self.files_results.addItem(it)
            self.files_tree.hide()
            self.files_results.show()
        self._run_bg(job, done, "Dhoondh rahe hain…")

    def _merge_folder_pdfs(self, folder):
        if not HAS_OCR_LIBS:
            self._warn("pypdf install nahi hai.")
            return
        pdfs = sorted(os.path.join(folder, f) for f in os.listdir(folder)
                      if f.lower().endswith(".pdf"))
        if len(pdfs) < 2:
            self._warn("Is folder me 2 se kam PDF hain.")
            return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Merged PDF save karein",
            os.path.join(folder, (os.path.basename(folder) or "sab") + "_sab.pdf"),
            "PDF (*.pdf)")
        if not out:
            return
        try:
            writer = PdfWriter()
            used = 0
            for p in pdfs:
                if os.path.abspath(p) == os.path.abspath(out):
                    continue
                try:
                    for pg in PdfReader(p).pages:
                        writer.add_page(pg)
                    used += 1
                except Exception:
                    pass
            with open(out, "wb") as fh:
                writer.write(fh)
        except Exception:
            self._warn("Merge fail:\n%s" % traceback.format_exc())
            return
        QtWidgets.QMessageBox.information(
            self, "Ho gaya", "%d PDFs jud kar ek ban gayi:\n%s" % (used, out))

    def _rename_library_file(self, path):
        stem, ext = os.path.splitext(os.path.basename(path))
        name, ok = QtWidgets.QInputDialog.getText(self, "Naam badlo", "Naya naam:", text=stem)
        if not ok or not name.strip():
            return
        new = os.path.join(os.path.dirname(path),
                           sanitize(underscore_name(name.strip())) + ext)
        try:
            os.rename(path, new)
        except Exception as exc:
            self._warn("Rename fail: %s" % exc)

    def _delete_library_file(self, path):
        if QtWidgets.QMessageBox.question(
                self, "Delete",
                "'%s' ko delete kar dein?\n(Ye wapas nahi aayegi)" % os.path.basename(path),
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.No) != QtWidgets.QMessageBox.Yes:
            return
        try:
            os.remove(path)
        except Exception as exc:
            self._warn("Delete fail: %s" % exc)

    def toggle_files_panel(self):
        vis = not self.files_panel.isVisible()
        self.files_panel.setVisible(vis)
        self._opts["show_files_panel"] = vis
        self._save_opts()

    def _update_preview_panel(self):
        if not getattr(self, "preview_panel", None) or not self.preview_panel.isVisible():
            return
        it = self.list.currentItem()
        if it is None:
            self.pv_img.clear(); self.pv_info.setText(""); return
        path = it.data(QtCore.Qt.UserRole)
        try:
            pm = QtGui.QPixmap(path)
            if not pm.isNull():
                self.pv_img.setPixmap(pm.scaled(
                    self.pv_img.width() - 8, self.pv_img.height() - 8,
                    QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
            name = it.data(TITLE_ROLE) or it.text() or "-"
            kb = os.path.getsize(path) / 1024.0
            self.pv_info.setText("%s\n%s · %dx%d" %
                                 (name, ("%.0f KB" % kb) if kb < 1024 else ("%.1f MB" % (kb / 1024)),
                                  pm.width(), pm.height()))
        except Exception:
            pass

    def _rebuild_jobs_bar(self):
        while self._jobs_lay.count():
            it = self._jobs_lay.takeAt(0)
            w = it.widget()
            if w:
                w.deleteLater()
        lbl = QtWidgets.QLabel(self.L("JOBS:", "JOBS:"))
        lbl.setStyleSheet("color:#92400e;font-weight:700;")
        self._jobs_lay.addWidget(lbl)
        for job in (self._opts.get("jobs") or []):
            b = QtWidgets.QPushButton(job.get("icon", "📋") + " " + job.get("name", "Job"))
            b.setStyleSheet("padding:3px 10px;border:1px solid #fbbf24;border-radius:99px;"
                            "background:#fff;")
            b.clicked.connect(lambda _c=False, j=job: self._apply_job(j))
            self._jobs_lay.addWidget(b)
        add = QtWidgets.QPushButton("➕")
        add.setToolTip(self.L("Naya job banao", "Create a job"))
        add.setStyleSheet("padding:3px 8px;border:1px dashed #fbbf24;border-radius:99px;background:#fff;")
        add.clicked.connect(self._new_job)
        self._jobs_lay.addWidget(add)
        self._jobs_lay.addStretch(1)

    def _apply_job(self, job):
        """Job-chip: ek click me profile + folder + naming set."""
        prof = job.get("profile")
        if prof:
            i = self.cmb_profile.findText(prof)
            if i >= 0:
                self.cmb_profile.setCurrentIndex(i)
        if job.get("folder"):
            self._opts["save_folder"] = job["folder"]
        if job.get("template"):
            self._opts["filename_template"] = job["template"]
        self._opts["auto_name"] = bool(job.get("auto_name", True))
        self._save_opts()
        self._refresh_files_root()
        self.status.showMessage(
            self.L("Job '%s' set ho gaya — ab scan karo." % job.get("name"),
                   "Job '%s' applied — scan now." % job.get("name")), 5000)

    def _new_job(self):
        name, ok = QtWidgets.QInputDialog.getText(
            self, self.L("Naya job", "New job"),
            self.L("Job ka naam (jaise Claim, ID-Cards):", "Job name (e.g. Claim, IDs):"))
        if not ok or not name.strip():
            return
        folder = QtWidgets.QFileDialog.getExistingDirectory(
            self, self.L("Is job ki files kis folder me jaayen?", "Save this job's files to which folder?"),
            self._opts.get("save_folder", ""))
        icon, _ = QtWidgets.QInputDialog.getText(
            self, "Icon", self.L("Ek emoji (optional):", "One emoji (optional):"), text="📋")
        prof = self.cmb_profile.currentText()
        jobs = self._opts.setdefault("jobs", [])
        jobs.append({"name": name.strip(), "icon": (icon.strip() or "📋")[:2],
                     "profile": prof, "folder": folder or "",
                     "template": "{name}_{date}", "auto_name": True})
        self._save_opts()
        self._rebuild_jobs_bar()

    def toggle_kiosk(self):
        on = not self.kiosk.isVisible()
        if on:
            self.kiosk.setGeometry(self.centralWidget().rect())
            self.kiosk.show(); self.kiosk.raise_()
        else:
            self.kiosk.hide()
        self._opts["ui_kiosk"] = on
        self._save_opts()

    def customize_ui(self):
        """UI ke design-elements — user apne hisaab se on/off kare (v24)."""
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle(self.L("🎨 UI customize", "🎨 Customize UI"))
        form = QtWidgets.QFormLayout(dlg)
        cmb = QtWidgets.QComboBox()
        cmb.addItems(["Light", "Dark", "Dark Pro"])
        cmb.setCurrentIndex({"light": 0, "dark": 1, "darkpro": 2}.get(self._opts.get("theme"), 0))
        form.addRow("Theme:", cmb)
        checks = {}
        default_on = {"ui_fab", "ui_preview", "ui_jobs", "ui_kiosk", "ui_ribbon"}
        for key, hi, en in (
                ("ui_ribbon", "Ribbon toolbar (MS-Office jaisa — tabs me buttons)",
                 "Ribbon toolbar (Office-style tabs)"),
                ("ui_dashboard", "Start dashboard (khaali screen par bade buttons)",
                 "Start dashboard (big buttons on empty screen)"),
                ("ui_header", "Status-patti (scanner/profile/aaj ka kaam)",
                 "Status header (scanner/profile/today)"),
                ("ui_preview", "Preview panel (page ki badi jhalak + quick-edit)",
                 "Preview panel (big page preview + quick edit)"),
                ("ui_jobs", "Job-chips patti (1 click me profile+folder set)",
                 "Job chips bar (1-click profile+folder)"),
                ("ui_fab", "Floating gol Scan button (neeche-daayein)",
                 "Floating round Scan button (bottom-right)"),
                ("ui_graph", "Sidebar me 7-din ka graph",
                 "7-day graph in the sidebar")):
            c = QtWidgets.QCheckBox(self.L(hi, en))
            c.setChecked(bool(self._opts.get(key, key not in default_on)))
            checks[key] = c
            form.addRow(c)
        bkiosk = QtWidgets.QPushButton(self.L("🖥 Kiosk mode ab chalu karo (bade buttons)",
                                             "🖥 Enter Kiosk mode now (big buttons)"))
        bkiosk.clicked.connect(lambda: (dlg.accept(), self.toggle_kiosk()))
        form.addRow(bkiosk)
        note = QtWidgets.QLabel(self.L(
            "<span style='color:#64748b;font-size:11px;'>Ribbon toolbar aur Icon-rail "
            "sidebars agle update me.</span>",
            "<span style='color:#64748b;font-size:11px;'>Ribbon toolbar and icon-rail "
            "sidebars coming in the next update.</span>"))
        note.setTextFormat(QtCore.Qt.RichText)
        form.addRow(note)
        bb = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject)
        form.addRow(bb)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return
        self._opts["theme"] = ["light", "dark", "darkpro"][cmb.currentIndex()]
        for key, c in checks.items():
            self._opts[key] = c.isChecked()
        self._save_opts()
        # turant lagao — restart ki zaroorat nahi
        self._apply_style()
        try:
            self.fab.setVisible(bool(self._opts.get("ui_fab")))
            self.ui_header.setVisible(bool(self._opts.get("ui_header")))
            self.side_graph.setVisible(bool(self._opts.get("ui_graph")))
            self.preview_panel.setVisible(bool(self._opts.get("ui_preview")))
            self.jobs_bar.setVisible(bool(self._opts.get("ui_jobs")))
            rib = bool(self._opts.get("ui_ribbon"))
            self.ribbon.setVisible(rib)
            self._classic_toolbar.setVisible(not rib)
            self._update_preview_panel()
            self._update_empty_state()
            self._update_sidebar_stats()
        except Exception:
            pass

    def toggle_left_panel(self):
        vis = not self.left_panel.isVisible()
        self.left_panel.setVisible(vis)
        self._opts["show_left_panel"] = vis
        self._save_opts()
        self.status.showMessage(
            self.L("Left sidebar %s (F9 se wapas)" % ("ON" if vis else "OFF"),
                   "Left sidebar %s (press F9 to toggle)" % ("ON" if vis else "OFF")), 4000)

    def L(self, hi, en):
        """Chhota helper: language ke hisaab se Hindi/English text."""
        return en if self._lang == "en" else hi

    # ---- Tags + OCR search index ----
    INDEX_PATH = os.path.join(os.path.expanduser("~"), ".apnescan_index.json")

    def _load_index(self):
        try:
            with open(self.INDEX_PATH, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return {}

    def _save_index(self, idx):
        try:
            with open(self.INDEX_PATH, "w", encoding="utf-8") as fh:
                json.dump(idx, fh)
        except Exception:
            pass

    def build_search_index(self):
        """Saari saved PDFs ka text EK BAAR padh kar index bana lo — uske baad
        'andar ke text' wali search turant hoti hai."""
        if not HAS_OCR_LIBS:
            self._warn("pypdf install nahi hai."); return
        root = self._opts.get("save_folder", os.path.expanduser("~"))

        def job():
            idx = self._load_index()
            pdfs = []
            for dp, _dn, fn in os.walk(root):
                for f in fn:
                    if f.lower().endswith(".pdf"):
                        pdfs.append(os.path.join(dp, f))
            changed = 0
            for p in pdfs:
                try:
                    m = os.path.getmtime(p)
                except Exception:
                    continue
                e = idx.get(p)
                if e and e.get("m") == m:
                    continue
                text = ""
                try:
                    reader = PdfReader(p)
                    for pg in reader.pages[:10]:
                        text += (pg.extract_text() or "")
                except Exception:
                    pass
                idx[p] = {"m": m, "t": (text or "").lower()[:20000]}
                changed += 1
            for p in list(idx):
                if not os.path.exists(p):
                    idx.pop(p, None)
            self._save_index(idx)
            return (len(pdfs), changed)

        def done(res):
            if isinstance(res, Exception):
                self._warn("Index fail:\n%s" % res); return
            total, changed = res
            QtWidgets.QMessageBox.information(
                self, "Index taiyar",
                "%d PDFs index me hain (%d nayi/badli padhi).\n\n"
                "Ab Ctrl+F wali search me 'andar ka text' turant milega.\n"
                "(Jo PDF 'OCR searchable' tick karke bani hain unhi ke andar text hota hai.)"
                % (total, changed))
        self._run_bg(job, done, "Search index ban raha hai… (app chalti rahegi)")

    def tag_pdf(self, src=None):
        """Kisi bhi saved PDF par tags lagao (jaise: Aadhaar, School, Bijli-bill)."""
        if not isinstance(src, str) or not src:
            src = self._pick_pdf("Kis PDF par tag lagana hai?")
        if not src:
            return
        tags = self._opts.setdefault("tags", {})
        cur = ", ".join(tags.get(src, []))
        text, ok = QtWidgets.QInputDialog.getText(
            self, "Tags", "Tags likhein (comma se alag, jaise: Aadhaar, Ghar):", text=cur)
        if not ok:
            return
        lst = [t.strip() for t in text.split(",") if t.strip()]
        if lst:
            tags[src] = lst
        else:
            tags.pop(src, None)
        self._save_opts()
        self.status.showMessage("Tags save ho gaye: %s" % (", ".join(lst) or "(hata diye)"), 5000)

    def search_by_tag(self):
        tags = self._opts.get("tags", {}) or {}
        all_tags = sorted({t for lst in tags.values() for t in lst})
        if not all_tags:
            self._warn("Abhi kisi PDF par tag nahi laga.\n(Tools → Tag lagao… se lagayein.)")
            return
        pick, ok = QtWidgets.QInputDialog.getItem(
            self, "Tag se dhoondo", "Tag chuno:", all_tags, 0, False)
        if not ok or not pick:
            return
        matches = [p for p, lst in tags.items() if pick in lst and os.path.exists(p)]
        if not matches:
            self._warn("Is tag ki koi file ab maujood nahi."); return
        item, ok = QtWidgets.QInputDialog.getItem(
            self, "'%s' wali files" % pick, "%d file mili — kholne ke liye chuno:" % len(matches),
            matches, 0, False)
        if ok and item:
            self._open_path(item)

    # ---- PDF tools (naye) ----
    def _pick_pdf(self, title):
        start = self._opts.get("save_folder", os.path.expanduser("~"))
        f, _ = QtWidgets.QFileDialog.getOpenFileName(self, title, start, "PDF (*.pdf)")
        return f or None

    def _run_bg(self, fn, on_done, busy_msg):
        """Bhaari kaam BACKGROUND me chalao; UI 1 second ke liye bhi nahi rukti.
        done hone par on_done(result) (main thread me). Ek non-blocking
        '⏳ kaam ho raha hai' indicator dikhta hai; kai kaam ek saath bhi
        chal sakte hain (counter se)."""
        self._bg_workers = getattr(self, "_bg_workers", [])
        self._bg_count = getattr(self, "_bg_count", 0) + 1
        self._show_busy_indicator(busy_msg)
        w = FuncWorker(fn)

        def _fin(res, w=w):
            self._bg_count = max(0, getattr(self, "_bg_count", 1) - 1)
            self._show_busy_indicator(None)
            try:
                self._bg_workers.remove(w)
            except ValueError:
                pass
            on_done(res)
        w.done.connect(_fin)
        self._bg_workers.append(w)
        w.start()

    def _show_busy_indicator(self, msg):
        """Status bar me ghoomta hua '⏳' — kaam chal raha hai par app chalu."""
        if not hasattr(self, "_busy_lbl"):
            self._busy_lbl = QtWidgets.QLabel("")
            self._busy_lbl.setStyleSheet(
                "QLabel{background:#0f766e;color:#fff;border-radius:9px;"
                "padding:1px 10px;font-weight:700;}")
            try:
                self.status.addPermanentWidget(self._busy_lbl)
            except Exception:
                pass
            self._busy_spin = ["⏳", "⌛"]
            self._busy_i = 0
            self._busy_timer = QtCore.QTimer(self)
            self._busy_timer.setInterval(500)

            def _tick():
                self._busy_i = (self._busy_i + 1) % len(self._busy_spin)
                base = getattr(self, "_busy_msg", "Kaam ho raha hai…")
                n = getattr(self, "_bg_count", 0)
                extra = (" (%d)" % n) if n > 1 else ""
                self._busy_lbl.setText("%s %s%s" %
                                       (self._busy_spin[self._busy_i], base, extra))
            self._busy_timer.timeout.connect(_tick)
        if msg:
            self._busy_msg = msg
        if getattr(self, "_bg_count", 0) > 0:
            self._busy_lbl.setText("⏳ " + getattr(self, "_busy_msg", "Kaam ho raha hai…"))
            self._busy_lbl.show()
            self._busy_timer.start()
        else:
            self._busy_timer.stop()
            self._busy_lbl.hide()

    def pdf_page_editor(self):
        """Kisi bhi PDF ke pages ka kram badlo / ghumao / hatao — bina quality
        kharaab kiye (lossless, pypdf se)."""
        if not HAS_OCR_LIBS:
            self._warn("pypdf install nahi hai."); return
        src = self._pick_pdf("Kaunsi PDF edit karni hai?")
        if not src:
            return
        try:
            reader = PdfReader(src)
            if reader.is_encrypted:
                pw, ok = QtWidgets.QInputDialog.getText(
                    self, "Password", "Is PDF ka password:", QtWidgets.QLineEdit.Password)
                if not ok or not reader.decrypt(pw or ""):
                    self._warn("Password galat hai."); return
            n = len(reader.pages)
        except Exception as exc:
            self._warn("PDF nahi khuli:\n%s" % exc); return
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("PDF page editor — %s (%d pages)" % (os.path.basename(src), n))
        dlg.resize(420, 460)
        v = QtWidgets.QVBoxLayout(dlg)
        lw = QtWidgets.QListWidget()
        for i in range(n):
            it = QtWidgets.QListWidgetItem("Page %d" % (i + 1))
            it.setData(QtCore.Qt.UserRole, [i, 0])   # [original index, extra rotation]
            lw.addItem(it)
        v.addWidget(lw, 1)
        row = QtWidgets.QHBoxLayout()

        def _move(d):
            r = lw.currentRow()
            if r < 0 or not (0 <= r + d < lw.count()):
                return
            it = lw.takeItem(r); lw.insertItem(r + d, it); lw.setCurrentItem(it)

        def _rot():
            it = lw.currentItem()
            if it:
                data = it.data(QtCore.Qt.UserRole)
                data[1] = (data[1] + 90) % 360
                it.setData(QtCore.Qt.UserRole, data)
                it.setText("Page %d  (ghuma: %d°)" % (data[0] + 1, data[1]))

        def _del():
            r = lw.currentRow()
            if r >= 0:
                lw.takeItem(r)
        for t, s in [("⬆ Upar", lambda: _move(-1)), ("⬇ Neeche", lambda: _move(1)),
                     ("↻ Ghumao", _rot), ("🗑 Hatao", _del)]:
            b = QtWidgets.QPushButton(t); b.clicked.connect(s); row.addWidget(b)
        v.addLayout(row)
        bb = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        if dlg.exec_() != QtWidgets.QDialog.Accepted or lw.count() == 0:
            return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Nayi PDF save karein", src[:-4] + "_edit.pdf", "PDF (*.pdf)")
        if not out:
            return
        try:
            writer = PdfWriter()
            for i in range(lw.count()):
                idx, rot = lw.item(i).data(QtCore.Qt.UserRole)
                pg = reader.pages[idx]
                if rot:
                    pg.rotate(rot)
                writer.add_page(pg)
            with open(out, "wb") as fh:
                writer.write(fh)
        except Exception:
            self._warn("Save fail:\n%s" % traceback.format_exc()); return
        self._remember_save_dir(out)
        QtWidgets.QMessageBox.information(self, "Ho gaya", "PDF ban gayi:\n%s" % out)

    def place_sign(self):
        """Apne sign/mohar ki image current page par lagao (white background
        apne aap transparent ho jata hai)."""
        item = self._current_item_or_warn()
        if not item:
            return
        sp = self._opts.get("sign_image") or ""
        if not sp or not os.path.exists(sp):
            f, _ = QtWidgets.QFileDialog.getOpenFileName(
                self, "Apne sign/stamp ki image chuno (white background wali)",
                "", "Images (*.png *.jpg *.jpeg *.bmp)")
            if not f:
                return
            self._opts["sign_image"] = f
            self._save_opts()
            sp = f
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Sign/Stamp lagao")
        form = QtWidgets.QFormLayout(dlg)
        POS = ["Neeche-daayein", "Neeche-beech", "Neeche-baayein",
               "Upar-daayein", "Upar-beech", "Upar-baayein", "Beech me"]
        cmb = QtWidgets.QComboBox(); cmb.addItems(POS)
        spn = QtWidgets.QSpinBox(); spn.setRange(8, 60); spn.setValue(22); spn.setSuffix(" % chaudai")
        btn = QtWidgets.QPushButton("Sign image badlo…")

        def _chg():
            f, _ = QtWidgets.QFileDialog.getOpenFileName(
                self, "Sign/stamp image", "", "Images (*.png *.jpg *.jpeg *.bmp)")
            if f:
                self._opts["sign_image"] = f; self._save_opts()
        btn.clicked.connect(_chg)
        form.addRow("Jagah:", cmb)
        form.addRow("Size:", spn)
        form.addRow(btn)
        bb = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject)
        form.addRow(bb)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return
        path = item.data(QtCore.Qt.UserRole)
        try:
            with Image.open(path) as im:
                page = im.convert("RGB").copy()
            with Image.open(self._opts["sign_image"]) as s:
                sign = s.convert("RGBA")
            g = sign.convert("L")
            sign.putalpha(g.point(lambda v: 0 if v > 230 else 255))
            w = max(30, int(page.width * spn.value() / 100.0))
            h = max(1, int(sign.height * w / sign.width))
            sign = sign.resize((w, h), Image.LANCZOS)
            m = max(20, page.width // 40)
            x_r, x_c, x_l = page.width - w - m, (page.width - w) // 2, m
            y_b, y_c, y_t = page.height - h - m, (page.height - h) // 2, m
            pos = [(x_r, y_b), (x_c, y_b), (x_l, y_b),
                   (x_r, y_t), (x_c, y_t), (x_l, y_t), (x_c, y_c)][cmb.currentIndex()]
            page.paste(sign, pos, sign)
            save_image_keep_ext(page, path)
            self._refresh_item(item)
            self._dirty = True
        except Exception:
            self._warn("Sign nahi laga:\n%s" % traceback.format_exc())

    def add_page_numbers(self):
        """Sab pages par 'Page X / N' (aur chaaho to upar apna text) chhapo."""
        paths = self._ordered_paths()
        if not paths:
            self._warn(tr("scan_first", self._lang)); return
        header, ok = QtWidgets.QInputDialog.getText(
            self, "Header (optional)", "Upar kya likhna hai? (khaali chhod sakte ho):")
        if not ok:
            return
        n = len(paths)
        for i, p in enumerate(paths, 1):
            try:
                with Image.open(p) as im:
                    img = im.convert("RGB").copy()
                d = ImageDraw.Draw(img)
                size = max(16, img.width // 50)
                try:
                    font = ImageFont.truetype("arial.ttf", size)
                except Exception:
                    font = ImageFont.load_default()
                foot = "Page %d / %d" % (i, n)
                try:
                    tw = d.textbbox((0, 0), foot, font=font)[2]
                except Exception:
                    tw = len(foot) * size // 2
                d.text(((img.width - tw) // 2, img.height - size - 12), foot,
                       fill=(70, 70, 70), font=font)
                if header.strip():
                    d.text((16, 10), header.strip()[:80], fill=(70, 70, 70), font=font)
                save_image_keep_ext(img, p)
            except Exception:
                pass
        for r in range(self.list.count()):
            self._refresh_item(self.list.item(r))
        self._dirty = True
        self.status.showMessage("Page numbers lag gaye.", 4000)

    def watermark_pdf_tool(self):
        """Kisi bhi purani PDF par watermark/stamp chhapo."""
        src = self._pick_pdf("Kis PDF par watermark lagana hai?")
        if not src:
            return
        text, ok = QtWidgets.QInputDialog.getText(
            self, "Watermark", "Kya likhna hai?",
            text=self._opts.get("watermark_text", "Noble Care Hospital"))
        if not ok or not text.strip():
            return
        pages = pdf_to_images(src, self._tmpdir)
        if not pages:
            self._warn("PDF se pages nahi nikle (PyMuPDF install karein)."); return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save karein", src[:-4] + "_wm.pdf", "PDF (*.pdf)")
        if not out:
            return
        try:
            imgs = [apply_watermark(Image.open(p).convert("RGB"), text.strip()) for p in pages]
            imgs[0].save(out, "PDF", save_all=True, append_images=imgs[1:], resolution=200.0)
            for im in imgs:
                im.close()
        except Exception:
            self._warn("Fail:\n%s" % traceback.format_exc()); return
        QtWidgets.QMessageBox.information(self, "Ho gaya", "Watermark wali PDF:\n%s" % out)

    def remove_pdf_password(self):
        """Password pata ho to PDF ki bina-password copy banao."""
        if not HAS_OCR_LIBS:
            self._warn("pypdf install nahi hai."); return
        src = self._pick_pdf("Password wali PDF chuno")
        if not src:
            return
        pw, ok = QtWidgets.QInputDialog.getText(
            self, "Password", "Is PDF ka password:", QtWidgets.QLineEdit.Password)
        if not ok:
            return
        try:
            reader = PdfReader(src)
            if reader.is_encrypted and not reader.decrypt(pw or ""):
                self._warn("Password galat hai."); return
            writer = PdfWriter()
            for pg in reader.pages:
                writer.add_page(pg)
            out = src[:-4] + "_unlocked.pdf"
            with open(out, "wb") as fh:
                writer.write(fh)
        except Exception:
            self._warn("Fail:\n%s" % traceback.format_exc()); return
        QtWidgets.QMessageBox.information(self, "Ho gaya", "Bina password wali copy:\n%s" % out)

    def pdf_to_jpgs(self):
        """Kisi bhi PDF ke pages JPG images me nikaalo."""
        src = self._pick_pdf("Kaunsi PDF ki images chahiye?")
        if not src:
            return
        folder = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Images kahan save karein?", os.path.dirname(src))
        if not folder:
            return
        pages = pdf_to_images(src, self._tmpdir)
        if not pages:
            self._warn("PDF se pages nahi nikle (PyMuPDF install karein)."); return
        base = os.path.splitext(os.path.basename(src))[0]
        cnt = 0
        for i, p in enumerate(pages, 1):
            try:
                with Image.open(p) as im:
                    im.convert("RGB").save(
                        os.path.join(folder, "%s_p%02d.jpg" % (base, i)), "JPEG", quality=90)
                cnt += 1
            except Exception:
                pass
        QtWidgets.QMessageBox.information(self, "Ho gaya", "%d images ban gayi:\n%s" % (cnt, folder))

    def folder_to_pdf(self):
        """Ek folder ki SAARI images (naam ke kram me) ek PDF me."""
        folder = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Images wala folder chuno", self._opts.get("save_folder", ""))
        if not folder:
            return
        exts = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")
        files = sorted(os.path.join(folder, f) for f in os.listdir(folder)
                       if f.lower().endswith(exts))
        if not files:
            self._warn("Is folder me koi image nahi mili."); return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "PDF save karein",
            os.path.join(folder, os.path.basename(folder) + ".pdf"), "PDF (*.pdf)")
        if not out:
            return
        try:
            imgs = [Image.open(f).convert("RGB") for f in files]
            imgs[0].save(out, "PDF", save_all=True, append_images=imgs[1:], resolution=200.0)
            for im in imgs:
                im.close()
        except Exception:
            self._warn("Fail:\n%s" % traceback.format_exc()); return
        QtWidgets.QMessageBox.information(
            self, "Ho gaya", "%d images ki PDF ban gayi:\n%s" % (len(files), out))

    def _collect_page_texts(self, src_pdf=None):
        """OCR se har page ka text (current pages ya kisi PDF ka). Bhaari kaam —
        FuncWorker ke andar hi bulayein."""
        if src_pdf:
            pages = pdf_to_images(src_pdf, self._tmpdir) or []
        else:
            pages = self._ordered_paths()
        texts = []
        for p in pages:
            try:
                with Image.open(p) as im:
                    texts.append(pytesseract.image_to_string(im, lang="eng+hin"))
            except Exception:
                texts.append("")
        return texts

    def pdf_to_word(self):
        """Pages/PDF ka text OCR karke Word (.docx) file banao."""
        if not tesseract_available():
            self._warn("Iske liye Tesseract OCR chahiye."); return
        src = None
        if not self._ordered_paths():
            src = self._pick_pdf("Kaunsi PDF ko Word banana hai?")
            if not src:
                return
        default = (src[:-4] + ".docx") if src else self._build_filename(".docx")
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Word file save karein", default, "Word (*.docx)")
        if not out:
            return

        def job():
            texts = self._collect_page_texts(src)
            try:
                import docx
                doc = docx.Document()
                for i, t in enumerate(texts):
                    if i:
                        doc.add_page_break()
                    doc.add_paragraph(t)
                doc.save(out)
                return out
            except ImportError:
                alt = os.path.splitext(out)[0] + ".txt"
                with open(alt, "w", encoding="utf-8") as fh:
                    fh.write("\n\n----\n\n".join(texts))
                return alt

        def done(res):
            if isinstance(res, Exception):
                self._warn("Word banane me fail:\n%s" % res); return
            QtWidgets.QMessageBox.information(self, "Ho gaya", "File ban gayi:\n%s" % res)
        self._run_bg(job, done, "Word bana rahe hain… (OCR chal raha hai)")

    def pdf_to_excel(self):
        """Bill/table wale pages ko OCR karke Excel me nikaalo (best-effort)."""
        if not tesseract_available() or not HAS_XLSX:
            self._warn("Iske liye Tesseract OCR + openpyxl chahiye."); return
        src = None
        if not self._ordered_paths():
            src = self._pick_pdf("Kaunsi PDF ko Excel banana hai?")
            if not src:
                return
        default = (src[:-4] + ".xlsx") if src else self._build_filename(".xlsx")
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Excel save karein", default, "Excel (*.xlsx)")
        if not out:
            return

        def job():
            pages = (pdf_to_images(src, self._tmpdir) or []) if src else self._ordered_paths()
            wb = openpyxl.Workbook()
            ws = wb.active
            for p in pages:
                with Image.open(p) as im:
                    data = pytesseract.image_to_data(
                        im.convert("RGB"), lang="eng+hin",
                        output_type=pytesseract.Output.DICT)
                lines = {}
                for i in range(len(data["text"])):
                    w = (data["text"][i] or "").strip()
                    if not w:
                        continue
                    key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
                    lines.setdefault(key, []).append((data["left"][i], data["width"][i], w))
                for key in sorted(lines):
                    words = sorted(lines[key])
                    cells, cur, last_end = [], [], None
                    for left, width, w in words:
                        if last_end is not None and left - last_end > 45:
                            cells.append(" ".join(cur)); cur = []
                        cur.append(w)
                        last_end = left + width
                    if cur:
                        cells.append(" ".join(cur))
                    ws.append(cells)
                ws.append([])
            wb.save(out)
            return out

        def done(res):
            if isinstance(res, Exception):
                self._warn("Excel banane me fail:\n%s" % res); return
            QtWidgets.QMessageBox.information(self, "Ho gaya", "Excel ban gayi:\n%s" % res)
        self._run_bg(job, done, "Excel bana rahe hain… (OCR chal raha hai)")

    def save_archival_pdf(self):
        """High-quality PDF + poora metadata (title/date/producer) — lambe samay
        tak sambhal kar rakhne ke liye."""
        paths = self._ordered_paths()
        if not paths:
            self._warn(tr("scan_first", self._lang)); return
        if not HAS_OCR_LIBS:
            self._warn("pypdf install nahi hai."); return
        default = self._build_filename(".pdf")
        default = (default[:-4] if default.lower().endswith(".pdf") else default) + "_archive.pdf"
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Archival PDF save karein", default, "PDF (*.pdf)")
        if not out:
            return
        try:
            imgs = [Image.open(p).convert("RGB") for p in paths]
            tmp = out + ".tmp.pdf"
            imgs[0].save(tmp, "PDF", save_all=True, append_images=imgs[1:],
                         resolution=300.0, quality=95)
            for im in imgs:
                im.close()
            reader = PdfReader(tmp)
            writer = PdfWriter()
            for pg in reader.pages:
                writer.add_page(pg)
            writer.add_metadata({
                "/Title": os.path.splitext(os.path.basename(out))[0],
                "/Producer": "ApneScan v%s" % VERSION,
                "/CreationDate": datetime.datetime.now().strftime("D:%Y%m%d%H%M%S"),
            })
            with open(out, "wb") as fh:
                writer.write(fh)
            os.remove(tmp)
        except Exception:
            self._warn("Fail:\n%s" % traceback.format_exc()); return
        self._remember_save_dir(out)
        self._record_save(out, len(paths))
        QtWidgets.QMessageBox.information(
            self, "Ho gaya", "Archival PDF (300dpi, metadata ke saath):\n%s" % out)

    def copy_page_text(self):
        """Selected page ka poora text OCR karke clipboard par."""
        item = self._current_item_or_warn()
        if not item:
            return
        if not tesseract_available():
            self._warn("Iske liye Tesseract OCR chahiye."); return
        path = item.data(QtCore.Qt.UserRole)

        def job():
            with Image.open(path) as im:
                return pytesseract.image_to_string(im, lang="eng+hin")

        def done(res):
            if isinstance(res, Exception):
                self._warn("OCR fail:\n%s" % res); return
            QtWidgets.QApplication.clipboard().setText(res or "")
            self.status.showMessage("Text copy ho gaya (Ctrl+V se kahin bhi paste karo).", 5000)
        self._run_bg(job, done, "Text padh rahe hain…")

    def translate_page(self):
        """Page ka text padh kar Hindi ↔ English translate karo (internet chahiye)."""
        item = self._current_item_or_warn()
        if not item:
            return
        if not tesseract_available():
            self._warn("Iske liye Tesseract OCR chahiye."); return
        opts = ["English → Hindi", "Hindi → English"]
        pick, ok = QtWidgets.QInputDialog.getItem(
            self, "Translate", "Kis taraf translate karein?", opts, 0, False)
        if not ok:
            return
        tl = "hi" if pick.startswith("English") else "en"
        path = item.data(QtCore.Qt.UserRole)

        def job():
            with Image.open(path) as im:
                text = pytesseract.image_to_string(im, lang="eng+hin")
            text = (text or "").strip()[:4500]
            if not text:
                raise RuntimeError("Page par padhne layak text nahi mila.")
            import urllib.parse as P
            import urllib.request as U
            import json as J
            url = ("https://translate.googleapis.com/translate_a/single"
                   "?client=gtx&sl=auto&tl=%s&dt=t&q=%s" % (tl, P.quote(text)))
            req = U.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            data = J.loads(U.urlopen(req, timeout=20).read().decode("utf-8", "ignore"))
            return "".join(seg[0] for seg in data[0] if seg and seg[0])

        def done(res):
            if isinstance(res, Exception):
                self._warn("Translate nahi ho paya (internet chal raha hai?):\n%s" % res)
                return
            dlg = QtWidgets.QDialog(self)
            dlg.setWindowTitle("Translation (%s)" % pick)
            dlg.resize(560, 420)
            v = QtWidgets.QVBoxLayout(dlg)
            ed = QtWidgets.QPlainTextEdit(res)
            v.addWidget(ed, 1)
            row = QtWidgets.QHBoxLayout()
            bcopy = QtWidgets.QPushButton("Copy")
            bcopy.clicked.connect(
                lambda: QtWidgets.QApplication.clipboard().setText(ed.toPlainText()))
            bcl = QtWidgets.QPushButton("Band karo"); bcl.clicked.connect(dlg.accept)
            row.addWidget(bcopy); row.addStretch(1); row.addWidget(bcl)
            v.addLayout(row)
            dlg.exec_()
        self._run_bg(job, done, "Translate ho raha hai…")

    # ---- page edit ----
    def _current_item_or_warn(self):
        item = self.list.currentItem()
        if item is None:
            self._warn("Pehle koi page select karein.")
        return item

    def _rotate(self, angle):
        item = self._current_item_or_warn()
        if not item:
            return
        path = item.data(QtCore.Qt.UserRole)
        try:
            with Image.open(path) as im:
                im.rotate(angle, expand=True).save(path)
        except Exception as exc:
            self._warn("Rotate fail:\n%s" % exc); return
        self._dirty = True
        self._refresh_item(item)

    def rotate_left(self):
        self._rotate(90)

    def rotate_right(self):
        self._rotate(-90)

    def _enhance_current(self, brightness, contrast):
        item = self._current_item_or_warn()
        if not item:
            return
        if enhance_image(item.data(QtCore.Qt.UserRole), brightness, contrast):
            self._refresh_item(item)

    def autocrop_current(self):
        item = self._current_item_or_warn()
        if not item:
            return
        path = item.data(QtCore.Qt.UserRole)
        try:
            with Image.open(path) as im:
                autocrop(im.convert("RGB")).save(path, "PNG")
        except Exception as exc:
            self._warn("Crop fail:\n%s" % exc); return
        self._refresh_item(item)

    def _list_context_menu(self, pos):
        if self.list.count() == 0:
            return
        menu = QtWidgets.QMenu(self)
        act_rename = menu.addAction("\u270f Naam badlo / Rename")
        act_del = menu.addAction("\U0001f5d1 Delete")
        chosen = menu.exec_(self.list.viewport().mapToGlobal(pos))
        if chosen == act_rename:
            self.rename_current_page()
        elif chosen == act_del:
            self.delete_page()

    def rename_current_page(self):
        it = self.list.currentItem() or (self.list.selectedItems() or [None])[0]
        if it is None:
            self._warn("Pehle koi page select karein."); return
        cur = it.data(TITLE_ROLE) or ""
        name, ok = QtWidgets.QInputDialog.getText(self, "Naam badlo", "Is page ka naam:", text=cur)
        if not ok:
            return
        name = underscore_name(name)
        if not name:
            return
        it.setData(TITLE_ROLE, name)
        it.setText(name)
        # LEARN: remember this document's content -> this name, so next time a
        # similar page is scanned it gets named automatically.
        self._learn_name(it.data(QtCore.Qt.UserRole), name)

    def _learn_name(self, path, name):
        """Rename ka OCR-learning ab BACKGROUND me hota hai — pehle yahi OCR
        UI thread par chal kar har rename ko 3-10 second ke liye jama deta tha."""
        if not path or not tesseract_available():
            return
        self._learn_workers = getattr(self, "_learn_workers", [])
        w = LearnWorker(path, name)
        w.got.connect(self._store_learned_name)
        w.finished.connect(lambda w=w: self._learn_workers.remove(w)
                           if w in self._learn_workers else None)
        self._learn_workers.append(w)
        w.start()

    def _store_learned_name(self, name, words):
        if len(words) < 4:
            return   # too little text to identify the form reliably
        learned = self._config.setdefault("learned_names", [])
        wset = set(words)
        # if a very similar signature already exists, update its name; else add new
        for e in learned:
            if _jaccard(wset, set(e.get("words", []))) >= 0.6:
                e["words"] = words
                e["name"] = name
                break
        else:
            learned.append({"words": words, "name": name})
            # keep the list from growing without bound
            if len(learned) > 300:
                del learned[0:len(learned) - 300]
        try:
            save_config(self._config)
        except Exception:
            pass
        self.status.showMessage("Naam yaad rakh liya \u2014 agli baar aisa document apne aap '%s' ho jayega." % name, 5000)

    def _list_context_menu_end(self):
        pass

    def delete_page(self):
        items = self.list.selectedItems()
        if not items:
            item = self._current_item_or_warn()
            if not item:
                return
            items = [item]
        # delete from the highest row down so earlier indices stay valid
        rows = sorted((self.list.row(it) for it in items), reverse=True)
        for row in rows:
            it = self.list.item(row)
            if it is None:
                continue
            path = it.data(QtCore.Qt.UserRole)
            self.list.takeItem(row)
            self._undo_stack.append((path, row))     # keep file for undo
        self._undo_stack = self._undo_stack[-15:]
        self._dirty = True
        self._renumber_pages()
        self._update_status(); self._update_empty_state()

    def move_up(self):
        row = self.list.currentRow()
        if row > 0:
            it = self.list.takeItem(row); self.list.insertItem(row - 1, it); self.list.setCurrentItem(it)

    def move_down(self):
        row = self.list.currentRow()
        if 0 <= row < self.list.count() - 1:
            it = self.list.takeItem(row); self.list.insertItem(row + 1, it); self.list.setCurrentItem(it)

    def clear_all(self):
        if self.list.count() == 0:
            return
        if QtWidgets.QMessageBox.question(self, "Confirm", "Saare pages hata dein?") != QtWidgets.QMessageBox.Yes:
            return
        for p in self._ordered_paths():
            try:
                os.remove(p)
            except Exception:
                pass
        self.list.clear()
        for p, _r in self._undo_stack:
            try:
                os.remove(p)
            except Exception:
                pass
        self._undo_stack = []
        self._dirty = False
        self._update_status(); self._update_empty_state()

    # ---- filenames + saving ----
    def _remember_save_dir(self, path):
        try:
            d = os.path.dirname(path)
            if d and os.path.isdir(d):
                self._config["last_save_dir"] = d
                save_config(self._config)
        except Exception:
            pass

    def _target_folder(self):
        # If the user isn't using auto-organizing folders, reopen wherever they
        # saved last time.
        last = self._config.get("last_save_dir")
        if (last and os.path.isdir(last)
                and not self._opts.get("year_month_folders")
                and not self._opts.get("make_claim_folder")):
            return last
        folder = self._opts.get("save_folder") or os.path.join(os.path.expanduser("~"), "Documents")
        now = datetime.datetime.now()
        if self._opts.get("year_month_folders"):
            folder = os.path.join(folder, now.strftime("%Y"), now.strftime("%m"))
        claim = sanitize(self.claim_edit.text().strip(), "")
        if self._opts.get("make_claim_folder") and claim:
            folder = os.path.join(folder, claim)
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception:
            pass
        return folder

    def _suggested_save_name(self, paths=None):
        """Name to default the save filename to. If specific paths are being saved
        (e.g. Save Selected), use the FIRST of THOSE pages' names — so each document
        saves with its own name instead of always the first page's."""
        pathset = set(paths) if paths else None
        for i in range(self.list.count()):
            it = self.list.item(i)
            if pathset is not None and it.data(QtCore.Qt.UserRole) not in pathset:
                continue
            t = it.data(TITLE_ROLE)
            if t:
                return underscore_name(t)
        return None

    def _remember_doc_name(self, out):
        """After saving, remember this saved name against the document title so the
        SAME document (next scan) shows this exact name again."""
        base = os.path.splitext(os.path.basename(out))[0]
        keys = []
        for i in range(self.list.count()):
            k = self.list.item(i).data(NAMEKEY_ROLE)
            if k:
                keys.append(k)
        if not keys:
            return
        dn = self._config.setdefault("doc_names", {})
        # map the first page's title (the document's identity) to this saved name
        dn[keys[0]] = base
        try:
            save_config(self._config)
        except Exception:
            pass

    def _build_filename(self, ext=".pdf", seq=1, paths=None):
        # If pages were auto-named, default the save filename to that document name
        # (spaces already underscored). Otherwise use the claim/date template.
        suggested = self._suggested_save_name(paths)
        if suggested:
            return os.path.join(self._target_folder(), suggested + ext)
        now = datetime.datetime.now()
        claim = sanitize(self.claim_edit.text().strip(), "scan")
        # Pehle page ka OCR-naam (ho to) — {name} placeholder ke liye
        docname = ""
        try:
            it0 = self.list.item(0)
            if it0:
                docname = underscore_name(it0.data(TITLE_ROLE) or "")
        except Exception:
            pass
        name = self._opts.get("filename_template", "{claim}_{date}_{seq}")
        name = (name.replace("{claim}", claim)
                    .replace("{date}", now.strftime("%Y-%m-%d"))
                    .replace("{year}", now.strftime("%Y"))
                    .replace("{month}", now.strftime("%m"))
                    .replace("{day}", now.strftime("%d"))
                    .replace("{time}", now.strftime("%H%M%S"))
                    .replace("{name}", docname or "scan")
                    .replace("{seq}", "%03d" % seq))
        return os.path.join(self._target_folder(), underscore_name(name) + ext)

    def _validate_claim_ok(self):
        if not self._opts.get("validate_claim"):
            return True
        claim = self.claim_edit.text().strip()
        pat = self._opts.get("claim_pattern") or r"^[A-Za-z0-9\-]{4,}$"
        try:
            if re.match(pat, claim):
                return True
        except Exception:
            return True
        self._warn("Claim number sahi nahi lag raha (pattern match nahi hua).\n"
                   "Sudhaar karein ya Settings me validation band karein.")
        return False

    def _duplicate_ok(self):
        if not self._opts.get("duplicate_check"):
            return True
        claim = self.claim_edit.text().strip()
        if claim and claim in self._used_claims:
            return QtWidgets.QMessageBox.question(
                self, "Duplicate", "Ye claim number pehle bhi save ho chuka hai.\nPhir bhi save karein?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.No) == QtWidgets.QMessageBox.Yes
        return True

    def _pages_as_pdf(self, paths, out, password=None):
        compress = self._opts.get("compress")
        q = int(self._opts.get("jpeg_quality", 60))
        wm = self._opts.get("watermark")
        wt = self._opts.get("watermark_text", "")
        imgs = []
        for p in paths:
            im = Image.open(p).convert("RGB")
            if wm and wt:
                im = apply_watermark(im, wt)
            if compress:
                buf = io.BytesIO(); im.save(buf, "JPEG", quality=q); buf.seek(0)
                im2 = Image.open(buf); im2.load(); im = im2
            imgs.append(im)
        if password:
            tmp = out + ".tmp.pdf"
            imgs[0].save(tmp, "PDF", save_all=True, append_images=imgs[1:], resolution=200.0)
            reader = PdfReader(tmp); writer = PdfWriter()
            for pg in reader.pages:
                writer.add_page(pg)
            writer.encrypt(password)
            with open(out, "wb") as fh:
                writer.write(fh)
            try:
                os.remove(tmp)
            except Exception:
                pass
        else:
            imgs[0].save(out, "PDF", save_all=True, append_images=imgs[1:], resolution=200.0)
        for im in imgs:
            try:
                im.close()
            except Exception:
                pass

    def _record_save(self, out, num_pages):
        claim = self.claim_edit.text().strip()
        # personal stats (sirf aapke PC par)
        try:
            it0 = self.list.item(0)
            dt = (it0.data(TITLE_ROLE) if it0 else "") or ""
            dt = dt.split("_")[0] if dt else None
        except Exception:
            dt = None
        self._pstats_bump(pages=num_pages, pdfs=1, doc_type=dt)
        # recent
        self._add_recent(out)
        # used claims (duplicate detection)
        if claim:
            self._used_claims.add(claim)
            self._config["used_claims"] = list(self._used_claims)[-2000:]
            save_config(self._config)
        # activity log
        if self._opts.get("activity_log"):
            try:
                logp = os.path.join(self._opts.get("save_folder", "."), "activity_log.txt")
                os.makedirs(os.path.dirname(logp), exist_ok=True)
                with open(logp, "a", encoding="utf-8") as fh:
                    fh.write("%s\tclaim=%s\tpages=%d\t%s\n"
                             % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                claim, num_pages, out))
            except Exception:
                pass
        # excel register
        if self._opts.get("excel_log") and HAS_XLSX:
            self._append_excel(claim, num_pages, out)
        # backup
        if self._opts.get("backup"):
            self._backup_file(out)

    def _append_excel(self, claim, num_pages, out):
        try:
            xp = os.path.join(self._opts.get("save_folder", "."), "register.xlsx")
            os.makedirs(os.path.dirname(xp), exist_ok=True)
            if os.path.exists(xp):
                wb = openpyxl.load_workbook(xp)
                ws = wb.active
            else:
                wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Register"
                ws.append(["Date", "Time", "Claim No.", "Pages", "File"])
            now = datetime.datetime.now()
            ws.append([now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S"),
                       claim, num_pages, out])
            wb.save(xp)
        except Exception:
            pass

    def _backup_file(self, out):
        try:
            day = datetime.datetime.now().strftime("%Y-%m-%d")
            bdir = os.path.join(self._opts.get("backup_folder", "."), day)
            os.makedirs(bdir, exist_ok=True)
            shutil.copy2(out, os.path.join(bdir, os.path.basename(out)))
        except Exception:
            pass

    def _auto_save_pdf(self):
        paths = self._ordered_paths()
        if not paths:
            return False
        if not self._validate_claim_ok() or not self._duplicate_ok():
            return False
        out = self._build_filename(".pdf")
        try:
            self._pages_as_pdf(paths, out)
        except Exception:
            self._warn("Auto-save fail:\n%s" % traceback.format_exc()); return False
        self._record_save(out, len(paths)); self._dirty = False
        if self._opts.get("save_images_too"):
            base = os.path.splitext(out)[0]
            for i, p in enumerate(paths):
                try:
                    with Image.open(p) as im:
                        im.convert("RGB").save("%s_p%d.jpg" % (base, i + 1), "JPEG", quality=90)
                except Exception:
                    pass
        self._after_save_action(out)
        self.status.showMessage("Auto-save: %s" % out, 8000)
        return True

    def save_pdf_all(self):
        self.save_pdf(self._ordered_paths())

    def save_pdf_selected(self):
        paths = self._selected_paths()
        if not paths:
            self._warn(tr("no_selection", self._lang)); return
        self.save_pdf(paths)

    def save_pdf(self, paths=None):
        if not isinstance(paths, (list, tuple)):
            paths = self._ordered_paths()
        if not paths:
            self._warn(tr("scan_first", self._lang)); return
        if not self._validate_claim_ok() or not self._duplicate_ok():
            return
        default = self._build_filename(".pdf", paths=paths)
        out, _ = QtWidgets.QFileDialog.getSaveFileName(self, "PDF save karein", default, "PDF (*.pdf)")
        if not out:
            return
        if not out.lower().endswith(".pdf"):
            out += ".pdf"
        ocr = HAS_OCR_LIBS and self.chk_ocr.isChecked()
        npages = len(paths)

        def job():
            if ocr:
                self._save_ocr_pdf(paths, out)   # OCR = sabse bhaari — ab background me
            else:
                self._pages_as_pdf(paths, out)
            return out

        def done(res):
            if isinstance(res, Exception):
                if HAS_OCR_LIBS and isinstance(res, pytesseract.TesseractNotFoundError):
                    self._warn("Tesseract OCR nahi mila. OCR ke bina try karein.")
                else:
                    self._warn("PDF save fail:\n%s" % res)
                return
            self._remember_save_dir(out); self._remember_doc_name(out)
            self._record_save(out, npages); self._dirty = False; self._after_save_action(out)
            self.status.showMessage("✔ PDF save ho gaya: %s" % out, 8000)
        self._run_bg(job, done,
                     "PDF save ho raha hai…" if not ocr else "OCR PDF ban rahi hai…")

    def save_pdf_password(self):
        paths = self._ordered_paths()
        if not paths:
            self._warn("Pehle koi page scan/import karein."); return
        if not HAS_OCR_LIBS:
            self._warn("pypdf install nahi hai (password ke liye zaroori)."); return
        pw, ok = QtWidgets.QInputDialog.getText(self, "Password", "PDF ka password:", QtWidgets.QLineEdit.Password)
        if not ok or not pw:
            return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(self, "PDF save karein", self._build_filename(".pdf"), "PDF (*.pdf)")
        if not out:
            return
        if not out.lower().endswith(".pdf"):
            out += ".pdf"
        npages = len(paths)

        def job():
            self._pages_as_pdf(paths, out, password=pw)
            return out

        def done(res):
            if isinstance(res, Exception):
                self._warn("PDF save fail:\n%s" % res); return
            self._remember_save_dir(out); self._remember_doc_name(out)
            self._record_save(out, npages); self._dirty = False; self._after_save_action(out)
            self.status.showMessage("✔ Password-PDF save ho gaya: %s" % out, 8000)
        self._run_bg(job, done, "Password-PDF ban rahi hai…")

    def save_images(self):
        paths = self._ordered_paths()
        if not paths:
            self._warn("Pehle koi page scan/import karein."); return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Images save karein", self._build_filename(".jpg"),
            "JPEG (*.jpg);;PNG (*.png);;TIFF (*.tif)")
        if not out:
            return
        base, ext = os.path.splitext(out)
        if not ext:
            ext = ".jpg"
        fmt = {"jpg": "JPEG", "jpeg": "JPEG", "png": "PNG", "tif": "TIFF", "tiff": "TIFF"}\
            .get(ext.lower().lstrip("."), "JPEG")
        npages = len(paths)

        def job():
            for i, p in enumerate(paths):
                target = (base + ext) if i == 0 else ("%s_%d%s" % (base, i + 1, ext))
                with Image.open(p) as im:
                    im.convert("RGB").save(target, fmt)
            return npages

        def done(res):
            if isinstance(res, Exception):
                self._warn("Images save fail:\n%s" % res); return
            self._remember_save_dir(out)
            self.status.showMessage("✔ %d image save ho gayi." % res, 8000)
        self._run_bg(job, done, "Images save ho rahi hain…")

    def export_ocr_text(self):
        if not HAS_OCR_LIBS:
            self._warn("pytesseract install nahi hai."); return
        paths = self._ordered_paths()
        if not paths:
            self._warn("Pehle koi page scan/import karein."); return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Text save karein", self._build_filename(".txt"), "Text (*.txt)")
        if not out:
            return
        def job():
            chunks = []
            for p in paths:
                with Image.open(p) as im:
                    chunks.append(pytesseract.image_to_string(im, lang="hin+eng"))
            with open(out, "w", encoding="utf-8") as fh:
                fh.write("\n\n----\n\n".join(chunks))
            return out

        def done(res):
            if isinstance(res, Exception):
                if isinstance(res, pytesseract.TesseractNotFoundError):
                    self._warn("Tesseract OCR install nahi hai.")
                else:
                    self._warn("Text export fail:\n%s" % res)
                return
            self._remember_save_dir(out)
            self.status.showMessage("✔ Text save ho gaya: %s" % out, 8000)
        self._run_bg(job, done, "OCR text nikal rahe hain…")

    def _save_ocr_pdf(self, paths, out):
        writer = PdfWriter()
        for p in paths:
            with Image.open(p) as im:
                pdf_bytes = pytesseract.image_to_pdf_or_hocr(im, extension="pdf", lang="hin+eng")
            reader = PdfReader(io.BytesIO(pdf_bytes))
            for pg in reader.pages:
                writer.add_page(pg)
        with open(out, "wb") as fh:
            writer.write(fh)

    # ---- Tools ----
    def split_pdfs(self):
        paths = self._ordered_paths()
        if not paths:
            self._warn("Pehle koi page scan/import karein."); return
        text, ok = QtWidgets.QInputDialog.getText(
            self, "Split PDFs", "Naye document kaun se page se shuru ho? (jaise 1,5,9):", text="1")
        if not ok or not text.strip():
            return
        try:
            starts = sorted(set(int(x) for x in text.replace(" ", "").split(",") if x))
        except Exception:
            self._warn("Galat input."); return
        starts = [s for s in starts if 1 <= s <= len(paths)]
        if not starts or starts[0] != 1:
            starts = [1] + starts
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Folder", self._opts.get("save_folder", ""))
        if not folder:
            return
        os.makedirs(folder, exist_ok=True)
        base = sanitize(self.claim_edit.text().strip(), "scan") + "_" + datetime.datetime.now().strftime("%Y-%m-%d")
        ranges = []
        for i, s in enumerate(starts):
            end = (starts[i + 1] - 1) if i + 1 < len(starts) else len(paths)
            ranges.append((s, end))

        def job():
            for idx, (s, e) in enumerate(ranges, 1):
                self._pages_as_pdf(paths[s - 1:e], os.path.join(folder, "%s_part%d.pdf" % (base, idx)))
            return len(ranges)

        def done(res):
            if isinstance(res, Exception):
                self._warn("Split fail:\n%s" % res); return
            QtWidgets.QMessageBox.information(self, "Ho gaya", "%d alag PDF ban gayi:\n%s" % (res, folder))
        self._run_bg(job, done, "PDF split ho rahi hai…")

    def merge_pdfs(self):
        if not HAS_OCR_LIBS:
            self._warn("pypdf install nahi hai."); return
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Merge ke liye PDF chuno", "", "PDF (*.pdf)")
        if not files:
            return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Merged PDF save", self._build_filename(".pdf"), "PDF (*.pdf)")
        if not out:
            return
        if not out.lower().endswith(".pdf"):
            out += ".pdf"

        def job():
            writer = PdfWriter()
            for f in files:
                r = PdfReader(f)
                for pg in r.pages:
                    writer.add_page(pg)
            with open(out, "wb") as fh:
                writer.write(fh)
            return out

        def done(res):
            if isinstance(res, Exception):
                self._warn("Merge fail:\n%s" % res); return
            self._add_recent(out)
            self.status.showMessage("✔ Merged PDF ban gayi: %s" % out, 8000)
        self._run_bg(job, done, "PDFs jud rahi hain (merge)…")

    def search_pdfs(self):
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Search past PDFs")
        v = QtWidgets.QVBoxLayout(dlg)
        v.addWidget(QtWidgets.QLabel("Claim number / naam / koi shabd:"))
        ed = QtWidgets.QLineEdit()
        v.addWidget(ed)
        chk = QtWidgets.QCheckBox("PDF ke ANDAR ka text bhi dekho (dheema — sirf OCR-wali PDFs me chalega)")
        v.addWidget(chk)
        bb = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        ed.setFocus()
        if dlg.exec_() != QtWidgets.QDialog.Accepted or not ed.text().strip():
            return
        q = ed.text().strip().lower()
        root = self._opts.get("save_folder", os.path.expanduser("~"))
        pdfs = []
        try:
            for dp, _dn, fn in os.walk(root):
                for f in fn:
                    if f.lower().endswith(".pdf"):
                        pdfs.append(os.path.join(dp, f))
        except Exception:
            pass
        matches = [p for p in pdfs if q in os.path.basename(p).lower()]
        # tag match bhi dikhao
        for p, lst in (self._opts.get("tags") or {}).items():
            if p not in matches and any(q in t.lower() for t in lst) and os.path.exists(p):
                matches.append(p)
        if chk.isChecked() and HAS_OCR_LIBS:
            # pehle INDEX se turant dhoondo; sirf un-indexed files live padho
            idx = self._load_index()
            todo = []
            for p in pdfs:
                if p in matches:
                    continue
                e = idx.get(p)
                if e is not None:
                    if q in (e.get("t") or ""):
                        matches.append(p)
                else:
                    todo.append(p)
            prog = QtWidgets.QProgressDialog(
                "PDF ke andar dhoondh rahe hain… (%d files)" % len(todo), "Cancel", 0, len(todo), self)
            prog.setWindowModality(QtCore.Qt.WindowModal)
            prog.setMinimumDuration(400)
            for i, p in enumerate(todo):
                prog.setValue(i)
                QtWidgets.QApplication.processEvents()
                if prog.wasCanceled():
                    break
                try:
                    reader = PdfReader(p)
                    text = ""
                    for pg in reader.pages[:5]:
                        text += (pg.extract_text() or "").lower()
                        if q in text:
                            matches.append(p)
                            break
                except Exception:
                    pass
            prog.setValue(len(todo))
        if not matches:
            QtWidgets.QMessageBox.information(
                self, "Search",
                "Kuch nahi mila.\n\nTip: PDF ke andar ke text me wahi PDF milti hai jo "
                "'OCR (searchable)' tick karke save hui thi.")
            return
        item, ok = QtWidgets.QInputDialog.getItem(
            self, "Search results", "%d file mili — kholne ke liye chuno:" % len(matches),
            matches, 0, False)
        if ok and item:
            self._open_path(item)

    def monthly_report(self):
        month = datetime.datetime.now().strftime("%Y-%m")
        total_files, total_pages = 0, 0
        # Prefer the Excel register if present
        xp = os.path.join(self._opts.get("save_folder", "."), "register.xlsx")
        counted = False
        if HAS_XLSX and os.path.exists(xp):
            try:
                wb = openpyxl.load_workbook(xp); ws = wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0] and str(row[0]).startswith(month):
                        total_files += 1
                        try:
                            total_pages += int(row[3])
                        except Exception:
                            pass
                counted = True
            except Exception:
                counted = False
        if not counted:
            root = self._opts.get("save_folder", os.path.expanduser("~"))
            try:
                for dp, _dn, fn in os.walk(root):
                    for f in fn:
                        if f.lower().endswith(".pdf"):
                            p = os.path.join(dp, f)
                            m = datetime.datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y-%m")
                            if m == month:
                                total_files += 1
            except Exception:
                pass
        QtWidgets.QMessageBox.information(
            self, "Monthly report (%s)" % month,
            "Is mahine:\n\nKul PDF: %d\nKul pages: %s"
            % (total_files, total_pages if total_pages else "—"))

    def _draw_fit(self, painter, path, target):
        """Draw the image at `path` scaled to fit inside `target` rect, centered,
        keeping aspect ratio."""
        img = QtGui.QImage(path)
        if img.isNull():
            return
        size = img.size()
        size.scale(target.size(), QtCore.Qt.KeepAspectRatio)
        x = target.x() + (target.width() - size.width()) // 2
        y = target.y() + (target.height() - size.height()) // 2
        painter.drawImage(QtCore.QRect(x, y, size.width(), size.height()), img)

    def _do_print(self, paths, per_page=1):
        if not paths:
            self._warn("Pehle koi page scan/import karein (ya select karein)."); return
        printer = QPrinter(QPrinter.HighResolution)
        if QPrintDialog(printer, self).exec_() != QtWidgets.QDialog.Accepted:
            return
        painter = QtGui.QPainter()
        if not painter.begin(printer):
            self._warn("Printer start nahi hua."); return
        try:
            page = painter.viewport()
            if per_page <= 1:
                for i, p in enumerate(paths):
                    if i > 0:
                        printer.newPage()
                    self._draw_fit(painter, p, page)
            else:
                # per_page images stacked on ONE sheet (top/bottom halves for 2-up).
                slot_h = page.height() // per_page
                for i, p in enumerate(paths):
                    slot = i % per_page
                    if i > 0 and slot == 0:
                        printer.newPage()
                    # small inset so the two IDs aren't glued together
                    pad = int(slot_h * 0.04)
                    target = QtCore.QRect(page.x(), page.y() + slot * slot_h + pad,
                                          page.width(), slot_h - 2 * pad)
                    self._draw_fit(painter, p, target)
        finally:
            painter.end()
        self.status.showMessage("Print bhej diya.", 4000)

    def print_all(self):
        self._do_print(self._ordered_paths(), per_page=1)

    def print_selected(self):
        paths = self._selected_paths()
        if not paths:
            self._warn("Pehle thumbnails me se kuch pages select karein."); return
        self._do_print(paths, per_page=1)

    def print_ids(self):
        # ID print: 2 IDs/pages per A4 sheet. Uses selected pages if any, else all.
        paths = self._selected_paths() or self._ordered_paths()
        self._do_print(paths, per_page=2)

    def print_ids_selected(self):
        # ID print of ONLY the selected pages (2 per A4 sheet).
        paths = self._selected_paths()
        if not paths:
            self._warn("Pehle thumbnails me se ID pages select karein."); return
        self._do_print(paths, per_page=2)

    # kept for the Ctrl+P shortcut / File menu (prints all)
    def print_pages(self):
        self.print_all()

    def create_shortcut(self):
        prof = self._selected_profile()
        if prof is None:
            self._warn("Pehle ek profile chuno."); return
        script = os.path.abspath(sys.argv[0])
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        bat = os.path.join(desktop, "Scan - %s.bat" % prof.get("name", "Profile"))
        try:
            with open(bat, "w", encoding="utf-8") as fh:
                fh.write("@echo off\r\n")
                fh.write('py -3.12-32 "%s" --scan --profile "%s"\r\n' % (script, prof.get("name", "Profile")))
            QtWidgets.QMessageBox.information(self, "Ho gaya",
                "Desktop par shortcut ban gaya:\n%s\n\nDouble-click karte hi seedhe scan." % bat)
        except Exception as exc:
            self._warn("Shortcut fail:\n%s" % exc)

    def _update_status(self):
        prof = self._selected_profile()
        pname = prof.get("name") if prof else tr("none_profile", self._lang)
        flags = [k for k in ("auto_save", "batch_mode", "remove_blank", "auto_crop",
                             "deskew", "quality_enhance", "watermark", "compress",
                             "excel_log", "backup") if self._opts.get(k)]
        tail = ("   |   " + ", ".join(flags)) if flags else ""
        self.status.showMessage("%s: %d   |   %s: %s%s" % (tr("st_pages", self._lang), self.list.count(), tr("st_profile", self._lang), pname, tail))

    def _warn(self, msg):
        QtWidgets.QMessageBox.warning(self, APP_NAME, msg)

    def closeEvent(self, event):
        if self.list.count() > 0 and getattr(self, "_dirty", False):
            r = QtWidgets.QMessageBox.question(
                self, APP_NAME,
                "Bina save kiye pages hain. Band karein?" if self._lang == "hi"
                else "You have unsaved pages. Close anyway?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
            if r != QtWidgets.QMessageBox.Yes:
                event.ignore(); return
        try:
            self._conn_timer.stop()
        except Exception:
            pass
        try:
            for f in os.listdir(self._tmpdir):
                try:
                    os.remove(os.path.join(self._tmpdir, f))
                except Exception:
                    pass
            os.rmdir(self._tmpdir)
        except Exception:
            pass
        super().closeEvent(event)


def main():
    auto_profile = None
    if "--scan" in sys.argv and "--profile" in sys.argv:
        try:
            auto_profile = sys.argv[sys.argv.index("--profile") + 1]
        except Exception:
            auto_profile = None
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NobleCare.ApneScan")
    except Exception:
        pass
    # CRASH REPORTER: koi bhi anhoni error file me save ho jaati hai taaki
    # report bhej kar agli update me fix ho sake.
    def _crash_hook(t, v, tb):
        try:
            with open(CRASH_PATH, "a", encoding="utf-8") as fh:
                fh.write("\n\n==== %s | ApneScan v%s ====\n" % (datetime.datetime.now(), VERSION))
                fh.write("".join(traceback.format_exception(t, v, tb)))
        except Exception:
            pass
        try:
            QtWidgets.QMessageBox.critical(
                None, "ApneScan — error",
                "App me ek error aa gayi. Report yahan save ho gayi:\n%s\n\n"
                "Ye file feedback me bhej dein — agli update me fix ho jayega." % CRASH_PATH)
        except Exception:
            pass
    sys.excepthook = _crash_hook

    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setWindowIcon(app_icon())
    win = ScannerWindow(auto_scan_profile=auto_profile)
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
