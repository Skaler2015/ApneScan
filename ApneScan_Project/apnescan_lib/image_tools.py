"""
image_tools.py — My Files ke image (JPG/PNG/…) par chalne wale saare
image tools (pure file<->file). Zyadatar kaam imaging.py ke tested helpers
se hota hai; yahan sirf file kholna/save aur parameter-handling hai.

ApneScan document-images ke liye hai, isliye tools bhi document/form-upload
wale hain (KB tak compress, passport-size, signature, deskew, B&W, PDF banao…).
"""

import os
from PIL import Image, ImageOps, ImageEnhance, ImageFilter

from . import imaging

# HEIC/HEIF (phone photos) — optional
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
    HAS_HEIF = True
except Exception:
    HAS_HEIF = False


# --------------------------------------------------------------------------
# load / save helpers
# --------------------------------------------------------------------------
def _load(src):
    im = Image.open(src)
    try:
        im = ImageOps.exif_transpose(im)         # phone rotation theek
    except Exception:
        pass
    return im


def _load_rgb(src):
    return _load(src).convert("RGB")


def _ext(path):
    return os.path.splitext(path)[1].lower().lstrip(".")


def _save(img, out, dpi=None, quality=92):
    """Image ko out ke extension ke hisaab se save karo."""
    ext = _ext(out)
    kw = {}
    if dpi:
        kw["dpi"] = (int(dpi), int(dpi))
    if ext in ("jpg", "jpeg"):
        img = img.convert("RGB")
        img.save(out, "JPEG", quality=quality, optimize=True, **kw)
    elif ext == "png":
        img.save(out, "PNG", optimize=True, **kw)
    elif ext == "webp":
        img.save(out, "WEBP", quality=quality, **({"dpi": kw["dpi"]} if dpi else {}))
    elif ext in ("tif", "tiff"):
        img.save(out, "TIFF", **kw)
    elif ext == "bmp":
        img.convert("RGB").save(out, "BMP")
    else:
        img.save(out)
    return out


# --------------------------------------------------------------------------
# A) Convert / format
# --------------------------------------------------------------------------
def convert_format(src, out, quality=92):
    """JPG/PNG/WEBP/BMP/TIFF me badlo (out ke extension se format tay)."""
    im = _load(src)
    if _ext(out) in ("jpg", "jpeg", "bmp"):
        im = im.convert("RGB")
    return _save(im, out, quality=quality)


def heic_to_jpg(src, out, quality=92):
    """HEIC/HEIF (iPhone) -> JPG."""
    if not HAS_HEIF:
        raise RuntimeError("pillow-heif not installed")
    return _save(_load_rgb(src), out, quality=quality)


def images_to_pdf(srcs, out, dpi=200):
    """Ek ya kai images -> ek PDF (kram me)."""
    imgs = [_load_rgb(s) for s in srcs]
    if not imgs:
        raise RuntimeError("no images")
    imgs[0].save(out, "PDF", save_all=True, append_images=imgs[1:],
                 resolution=float(dpi))
    for im in imgs:
        im.close()
    return out


# --------------------------------------------------------------------------
# B) Size / compress
# --------------------------------------------------------------------------
def _jpeg_bytes(img, quality):
    import io
    b = io.BytesIO()
    img.convert("RGB").save(b, "JPEG", quality=quality, optimize=True)
    return b.getvalue()


def compress_to_kb(src, out, target_kb, min_quality=20):
    """Image ko target_kb (KB) se NEECHE laao — pehle quality ghatakar, phir
    zaroorat pade to naap (dimensions) bhi. Form-upload ke KB-limit ke liye."""
    target = int(target_kb) * 1024
    base = _load_rgb(src)
    scale = 1.0
    best = None
    for _ in range(8):                            # zyada se zyada 8 baar chhota
        w = max(1, int(base.width * scale))
        h = max(1, int(base.height * scale))
        im = base if scale == 1.0 else base.resize((w, h), Image.LANCZOS)
        lo, hi, chosen = min_quality, 95, None
        while lo <= hi:                           # binary search quality
            q = (lo + hi) // 2
            data = _jpeg_bytes(im, q)
            if len(data) <= target:
                chosen = (q, data); lo = q + 1
            else:
                hi = q - 1
        if chosen:
            best = chosen[1]
            break
        # min_quality par bhi bada -> aur chhota karo
        best = _jpeg_bytes(im, min_quality)
        if len(best) <= target:
            break
        scale *= 0.85
    with open(out, "wb") as fh:
        fh.write(best)
    return out


def resize_to(src, out, width=None, height=None, keep_aspect=True, quality=92):
    """Nayi chaudai/oonchai par resize. keep_aspect: ek diya ho to doosra
    apne aap; dono diye ho to aspect banae rakhne ko fit karta hai."""
    im = _load(src)
    w0, h0 = im.size
    if width and height:
        if keep_aspect:
            r = min(width / w0, height / h0)
            nw, nh = max(1, int(w0 * r)), max(1, int(h0 * r))
        else:
            nw, nh = int(width), int(height)
    elif width:
        nw = int(width); nh = max(1, int(h0 * (width / w0)))
    elif height:
        nh = int(height); nw = max(1, int(w0 * (height / h0)))
    else:
        nw, nh = w0, h0
    return _save(im.resize((nw, nh), Image.LANCZOS), out, quality=quality)


def resize_percent(src, out, pct, quality=92):
    im = _load(src)
    r = max(1, pct) / 100.0
    nw = max(1, int(im.width * r)); nh = max(1, int(im.height * r))
    return _save(im.resize((nw, nh), Image.LANCZOS), out, quality=quality)


def set_dpi(src, out, dpi):
    return _save(_load(src), out, dpi=dpi)


# --------------------------------------------------------------------------
# C) Rotate / fix
# --------------------------------------------------------------------------
def rotate(src, out, degrees, quality=95):
    im = _load(src).rotate(-int(degrees), expand=True, fillcolor=(255, 255, 255))
    return _save(im, out, quality=quality)


def flip(src, out, horizontal=True, quality=95):
    im = _load(src)
    im = ImageOps.mirror(im) if horizontal else ImageOps.flip(im)
    return _save(im, out, quality=quality)


def deskew(src, out, quality=95):
    return _save(imaging.ai_deskew(_load_rgb(src)), out, quality=quality)


def auto_crop(src, out, quality=95):
    return _save(imaging.ai_smart_crop(_load_rgb(src)), out, quality=quality)


def photo_to_scan(src, out, quality=95):
    """Phone-photo ko flatbed-scan jaisa (seedha + saaf + shadow hatao)."""
    return _save(imaging.ai_photo_to_scan(_load_rgb(src)), out, quality=quality)


# --------------------------------------------------------------------------
# D) Enhance / clean
# --------------------------------------------------------------------------
def auto_enhance(src, out, quality=95):
    return _save(imaging.ai_auto_enhance(_load_rgb(src)), out, quality=quality)


def grayscale(src, out, quality=95):
    return _save(_load(src).convert("L"), out, quality=quality)


def black_white(src, out):
    return _save(imaging.adaptive_bw(_load_rgb(src)), out)


def remove_shadow(src, out, quality=95):
    """Chaya hatao / background safed (flatten)."""
    return _save(imaging.flatten_background(_load_rgb(src)), out, quality=quality)


def adjust_bc(src, out, brightness=1.0, contrast=1.0, quality=95):
    im = _load_rgb(src)
    if brightness != 1.0:
        im = ImageEnhance.Brightness(im).enhance(brightness)
    if contrast != 1.0:
        im = ImageEnhance.Contrast(im).enhance(contrast)
    return _save(im, out, quality=quality)


def denoise(src, out, quality=95):
    return _save(imaging.denoise(_load_rgb(src)), out, quality=quality)


def sharpen(src, out, quality=95):
    return _save(imaging.sharpen_clarity(_load_rgb(src)), out, quality=quality)


# --------------------------------------------------------------------------
# E) Combine
# --------------------------------------------------------------------------
def merge_images(srcs, out, direction="vertical", gap=10,
                 bg=(255, 255, 255), quality=92):
    """Kai images ko ek me jodo — vertical (upar-neeche) ya horizontal
    (agal-bagal). Chhoti images ko sabse badi chaudai/oonchai par centre."""
    imgs = [_load_rgb(s) for s in srcs]
    if not imgs:
        raise RuntimeError("no images")
    if direction == "horizontal":
        H = max(i.height for i in imgs)
        W = sum(i.width for i in imgs) + gap * (len(imgs) - 1)
        canvas = Image.new("RGB", (W, H), bg)
        x = 0
        for i in imgs:
            canvas.paste(i, (x, (H - i.height) // 2)); x += i.width + gap
    else:
        W = max(i.width for i in imgs)
        H = sum(i.height for i in imgs) + gap * (len(imgs) - 1)
        canvas = Image.new("RGB", (W, H), bg)
        y = 0
        for i in imgs:
            canvas.paste(i, ((W - i.width) // 2, y)); y += i.height + gap
    for i in imgs:
        i.close()
    return _save(canvas, out, quality=quality)


# --------------------------------------------------------------------------
# F) Form-specific: passport photo, signature
# --------------------------------------------------------------------------
def passport_photo(src, out, size_mm=(35, 45), dpi=300,
                   bg=(255, 255, 255), quality=95):
    """Passport-size (default 35x45mm) par center-crop + white background.
    (Chehra apne aap centre nahi hota — pehle theek se crop karke dein to
    best; ye size/aspect standard bana deta hai.)"""
    tw = int(round(size_mm[0] / 25.4 * dpi))
    th = int(round(size_mm[1] / 25.4 * dpi))
    im = _load_rgb(src)
    # target aspect par center-crop
    ar_t = tw / th
    ar_i = im.width / im.height
    if ar_i > ar_t:                               # bahut chaudi -> kinare kaato
        nw = int(im.height * ar_t)
        x = (im.width - nw) // 2
        im = im.crop((x, 0, x + nw, im.height))
    else:                                         # bahut lambi -> upar-neeche
        nh = int(im.width / ar_t)
        y = (im.height - nh) // 2
        im = im.crop((0, y, im.width, y + nh))
    im = im.resize((tw, th), Image.LANCZOS)
    canvas = Image.new("RGB", (tw, th), bg)
    canvas.paste(im, (0, 0))
    return _save(canvas, out, dpi=dpi, quality=quality)


def signature_png(src, out):
    """Sign/stamp nikaalo — background TRANSPARENT, sirf syahi. PNG me save.
    Kuch na mile to poora (trim karke) laut jaata hai."""
    res = imaging.ai_extract_signature(_load_rgb(src))
    if res is None:
        # fallback: bas trim + white bg
        im = _load_rgb(src)
        bbox = ImageOps.invert(im.convert("L")).getbbox()
        if bbox:
            im = im.crop(bbox)
        im.save(out if out.lower().endswith(".png") else out + ".png", "PNG")
        return out
    if not out.lower().endswith(".png"):
        out = os.path.splitext(out)[0] + ".png"
    res.save(out, "PNG")
    return out


# --------------------------------------------------------------------------
# G) Mark
# --------------------------------------------------------------------------
def watermark(src, out, text, quality=95):
    return _save(imaging.apply_watermark(_load_rgb(src), text), out, quality=quality)


# --------------------------------------------------------------------------
# H) Info / privacy
# --------------------------------------------------------------------------
def info(src):
    """Image ki jaankari: format, mode, naap (px), DPI, file-size."""
    im = Image.open(src)
    dpi = im.info.get("dpi", (None, None))
    try:
        sz = os.path.getsize(src)
    except Exception:
        sz = 0
    return {
        "format": im.format,
        "mode": im.mode,
        "width": im.width,
        "height": im.height,
        "dpi": dpi[0] if dpi and dpi[0] else None,
        "bytes": sz,
    }


def strip_metadata(src, out, quality=95):
    """EXIF/GPS/camera metadata hatao (privacy) — pixel data waisi hi."""
    im = _load(src)
    clean = Image.new(im.mode, im.size)
    clean.putdata(list(im.getdata()))
    return _save(clean, out, quality=quality)
