"""Image-processing engine — pure page-cleanup transforms (PIL + numpy).

Extracted verbatim from apnescan.py. These functions take and return PIL images
(or operate on image files) and never touch application/UI state or OCR, so the
whole scan-cleanup pipeline can be maintained and tested in isolation. Behaviour
is byte-for-byte identical to the original definitions.

The OCR-based Smart Orientation engine (auto_orient / detect_orientation) is NOT
here — it stays with the OCR bootstrap and will move in a later phase.
"""

import datetime

from PIL import (
    Image, ImageChops, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps,
)

try:
    import numpy as np
    HAS_NUMPY = True
except Exception:  # pragma: no cover - numpy always present in the app build
    HAS_NUMPY = False

__all__ = [
    "is_blank_page",
    "whiten_dark_background",
    "trim_dark_borders",
    "autocrop",
    "deskew",
    "flatten_background",
    "auto_enhance",
    "auto_brightness",
    "denoise",
    "apply_enhance_mode",
    "clean_edges",
    "split_two_pages",
    "flatten_photo_shadows",
    "clean_photo",
    "_largest_gap",
    "detect_content_boxes",
    "colorfulness",
    "restore_photo",
    "save_image_keep_ext",
    "apply_watermark",
    "enhance_image",
]


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


def flatten_background(img, target=246):
    """Photocopy/scan ki GRAY-maili background ko asli SAFED banao (HD print).

    Kaise: background ko bade neighbourhood (max-filter + blur, chhote size par)
    se estimate karke image ko usse DIVIDE karte hain — kagaz har jagah ek-jaisa
    white ho jaata hai (streaks/shading bhi saaf), phir halki gamma se text
    wapas gehra kiya jaata hai. Rang wali image me teeno channels par same gain
    lagta hai (rang kharab nahi hote).

    Surakshit: photo / X-ray jaise zyada-dark pages par apne aap NO-OP
    (heuristic: kam-se-kam 55% pixels halke hone chahiye), aur pehle se safed
    page par asar na ke barabar hota hai."""
    try:
        if not HAS_NUMPY:
            return img
        g = np.asarray(img.convert("L"), dtype=np.float32)
        # Document heuristic — photo/X-ray (zyada dark) par haath mat lagao.
        if float((g > 120).mean()) < 0.55:
            return img
        h, w = g.shape
        # Background chhote size par nikalta hai (tez): max-filter text ke
        # strokes ke UPAR se kagaz utha leta hai, blur use smooth karta hai.
        small = img.convert("L").resize(
            (max(1, w // 8), max(1, h // 8)), Image.BILINEAR)
        # Pehle se SAFED page (naye 300dpi HD scan) par kaam hi mat karo —
        # print ke waqt ye per-page bachat hai (no-op = same object return).
        if float(np.percentile(np.asarray(small, dtype=np.float32), 80)) >= 243.0:
            return img
        bg = small.filter(ImageFilter.MaxFilter(9)).filter(
            ImageFilter.GaussianBlur(6)).resize((w, h), Image.BILINEAR)
        b = np.maximum(np.asarray(bg, dtype=np.float32), 40.0)
        bgmean = float(b.mean())
        gain = float(target) / b
        # ADF ki KHADI (vertical) streak-lines: har column ka apna chhota
        # correction — patli lines global blur me nahi pakdi jaati.
        gray1 = np.clip(g * gain, 0.0, 255.0)
        # (har 6th row kaafi hai — percentile hi sabse mehnga step tha)
        col = np.percentile(gray1[::6], 85, axis=0)
        cgain = np.clip(float(target) / np.maximum(col, 120.0), 0.95, 1.2)
        gray1 *= cgain[None, :]
        # Gamma LUT (np.power poori image par bahut dheema tha — LUT ~10x tez):
        # jitni maili background thi utna text gehra hota hai; end me white-
        # point stretch se kagaz PURA safed (255) tak pahunchta hai.
        gamma = min(1.7, max(1.1, 1.1 + (235.0 - bgmean) * 0.008))
        lut = np.power(np.arange(256, dtype=np.float32) / 255.0, gamma) * 255.0
        wp = float(np.percentile(
            lut[np.clip(gray1[::8, ::8], 0, 255).astype(np.uint8)], 90))
        lut = np.clip(lut * min(1.15, 255.0 / max(wp, 200.0)), 0.0, 255.0)
        lut = lut.astype(np.uint8)
        if img.mode == "L":
            out = np.clip(gray1, 0.0, 255.0).astype(np.uint8)
            return Image.fromarray(lut[out], "L")
        rgb = np.asarray(img.convert("RGB"), dtype=np.float32)
        out = np.clip(rgb * (gain * cgain[None, :])[..., None],
                      0.0, 255.0).astype(np.uint8)
        return Image.fromarray(lut[out], "RGB")
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


def auto_brightness(img):
    """F7: normalise exposure so pages are neither washed-out nor too dark.
    Nudges brightness toward a bright-but-not-blown target using the grey mean;
    gentle, clamped, and a no-op when the page is already well-exposed."""
    try:
        from PIL import ImageStat
        mean = ImageStat.Stat(img.convert("L")).mean[0] or 128.0
        if 170 <= mean <= 225:                 # already good -> leave it alone
            return img
        factor = max(0.75, min(1.6, 200.0 / max(1.0, mean)))
        return ImageEnhance.Brightness(img).enhance(factor)
    except Exception:
        return img


def denoise(img):
    """F8: remove small dust/specks/compression noise while keeping text sharp
    (a light median filter — does not blur edges noticeably)."""
    try:
        return img.convert("RGB").filter(ImageFilter.MedianFilter(3))
    except Exception:
        return img


def apply_enhance_mode(img, mode):
    """F5: background-enhancement modes — 'original' (untouched), 'white' (clean
    white background), 'enhanced' (contrast+sharpness+brightness), 'high_contrast'
    (crisp near-B&W text). Unknown/'original' returns the image unchanged."""
    try:
        if mode == "white":
            return ImageOps.autocontrast(whiten_dark_background(img).convert("RGB"), cutoff=1)
        if mode == "enhanced":
            return auto_enhance(auto_brightness(img))
        if mode == "high_contrast":
            base = ImageOps.autocontrast(auto_brightness(img).convert("RGB"), cutoff=3)
            return ImageEnhance.Contrast(base).enhance(1.35)
    except Exception:
        pass
    return img


def clean_edges(img, margin_frac=0.018, dark_level=90):
    """Scan ke kinaron par aane wali KAALI border aur kinare ke chhed (punch-hole)
    ke nishan saaf karo. Sirf bahari patti (margin) dekhi jaati hai — beech ka
    text/stamp bilkul nahi chhua jaata. Margin ke andar ke gehre (dark) pixel
    safed kar diye jaate hain."""
    try:
        rgb = img.convert("RGB")
        w, h = rgb.size
        m = max(2, int(min(w, h) * margin_frac))
        px = rgb.load()
        white = (255, 255, 255)

        def _row(y):
            for x in range(w):
                r, g, b = px[x, y][:3]
                if (r + g + b) / 3 < dark_level:
                    px[x, y] = white

        def _col(x):
            for y in range(h):
                r, g, b = px[x, y][:3]
                if (r + g + b) / 3 < dark_level:
                    px[x, y] = white
        for y in list(range(0, m)) + list(range(h - m, h)):
            _row(y)
        for x in list(range(0, m)) + list(range(w - m, w)):
            _col(x)
        return rgb
    except Exception:
        return img


def split_two_pages(img, min_aspect=1.15):
    """Ek glass par do page rakhe ho (chaudi/landscape scan) to unhe do alag
    page me kaato. Beech me sabse safed (khaali) vertical patti dhoondh kar wahi
    se kaatte hain. Do image ki list lautata hai; kaatne layak na ho to [img]."""
    try:
        rgb = img.convert("RGB")
        w, h = rgb.size
        if w < h * min_aspect:      # chaudi nahi — split ki zaroorat nahi
            return [rgb]
        gray = rgb.convert("L")
        if HAS_NUMPY:
            import numpy as _np
            a = _np.asarray(gray, dtype=_np.float32)
            col_mean = a.mean(axis=0)                 # har column ki roshni
            lo, hi = int(w * 0.35), int(w * 0.65)     # beech ke 30% me hi gutter
            band = col_mean[lo:hi]
            cut = lo + int(band.argmax())             # sabse safed column
            if col_mean[cut] < 200:                   # itni safed patti nahi mili
                cut = w // 2
        else:
            cut = w // 2
        left = rgb.crop((0, 0, cut, h))
        right = rgb.crop((cut, 0, w, h))
        return [left, right]
    except Exception:
        return [img]


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


def save_image_keep_ext(img, path, quality=95):
    """Image ko uske extension ke hisaab se sahi format me save karo.
    HD: quality 95 (baar-baar edit-save par bhi quality na gire), page ka
    GRAY mode aur dpi metadata surakshit rehte hain (PDF-size + print-size)."""
    kw = {}
    try:
        d = img.info.get("dpi")
        if d:
            kw["dpi"] = (int(round(d[0])), int(round(d[1])))
    except Exception:
        pass
    if path.lower().endswith((".jpg", ".jpeg")):
        out = img if img.mode == "L" else img.convert("RGB")
        out.save(path, "JPEG", quality=quality, **kw)
    else:
        img.save(path, "PNG", **kw)


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
