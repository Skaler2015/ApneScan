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
    "dewarp_page",
    "flatten_background",
    "smart_jpeg_quality",
    "has_real_colour",
    "adaptive_bw",
    "straighten_photo_page",
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
    "ai_auto_enhance",
    "ai_color_restore",
    "ai_denoise",
    "ai_smart_crop",
    "ai_deskew",
    "ai_auto_all",
    "ai_inpaint_region",
    "ai_extract_signature",
    "ai_scan_quality",
    "ai_photo_to_scan",
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
        # (v221) STREAKY backing band bhi poori white karo. Scanner backing me
        # kabhi halki dhaariyan (vertical/horizontal streaks) hoti hain — tab
        # upar wala first/last-logic band ke beech ka dark chhod deta tha
        # (page ke NEECHE/upar gray patti reh jaati thi). Fix: har row ka
        # 'paper-hona' (bright-fraction) SMOOTH karke asli paper ka upar/neeche
        # kinara dhoondho; us kinare ke BAHAR ki har row poori white. Streak
        # (patli bright line) smoothing me dab jaati hai, isliye galat 'paper'
        # nahi banti; page ke andar ka content (frac ooncha) surakshit.
        def _smooth(v, k=15):
            if len(v) < k:
                return v
            ker = _np.ones(k, dtype=_np.float32) / k
            return _np.convolve(v.astype(_np.float32), ker, mode="same")
        rf = _smooth(bright.mean(axis=1))
        paper_rows = _np.where(rf > 0.45)[0]
        if paper_rows.size:
            pt, pb = int(paper_rows[0]), int(paper_rows[-1])
            mask[:pt, :] = True          # paper ke UPAR sab backing -> white
            mask[pb + 1:, :] = True      # paper ke NEECHE sab backing -> white
        cf = _smooth(bright.mean(axis=0))
        paper_cols = _np.where(cf > 0.45)[0]
        if paper_cols.size:
            pl, pr = int(paper_cols[0]), int(paper_cols[-1])
            mask[:, :pl] = True          # paper ke BAAYIN sab backing -> white
            mask[:, pr + 1:] = True      # paper ke DAAYIN sab backing -> white
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
        # Document-guard: deskew sirf KAGAZ jaise page par (zyada halka
        # background) — photo/X-ray par rotation ke white kone score ko
        # jhootha badha dete the aur bina wajah ghumav lag jaata tha.
        # (subsample [::4] — resize ka averaging noise ko 'halka' dikha deta tha)
        if float((np.asarray(g, dtype=np.float32)[::4, ::4] > 150).mean()) < 0.55:
            return img

        def _score(angle):
            rot = small.rotate(angle, resample=Image.BILINEAR, fillcolor=255)
            arr = np.asarray(rot, dtype=np.float32)
            return float(np.var(np.diff(arr.sum(axis=1))))

        score0 = _score(0.0)
        # (F2) BADA tilt bhi: pehle mota search ±44° (2° step). Bada angle
        # sirf tab mana jaata hai jab uska score 0° se SAAF (1.35x) behtar
        # ho — warna photo/table jaise page par galat 30° ghumav lag jaata.
        center = 0.0
        best_c, best_cs = 0.0, score0
        for i in range(-22, 23):
            a = i * 2.0
            if a == 0.0:
                continue
            s = _score(a)
            if s > best_cs:
                best_cs, best_c = s, a
        if abs(best_c) > 5.0 and best_cs > 1.35 * score0:
            center = best_c
        # fine search: center ke aas-paas ±5° (0.5° step)
        best_angle, best_score = center, _score(center)
        for i in range(-10, 11):
            angle = center + i * 0.5
            s = _score(angle)
            if s > best_score:
                best_score, best_angle = s, angle
        if abs(best_angle) < 0.25:
            return img
        # Quality-guard: ghumane par text-lines SAAF behtar honi chahiye —
        # photo/naqsha jaise page par bina wajah rotation nahi lagti.
        if best_score < 1.15 * score0:
            return img
        return img.rotate(best_angle, resample=Image.BICUBIC, fillcolor=(255, 255, 255), expand=True)
    except Exception:
        return img


def dewarp_page(img, max_shift_frac=0.04):
    """(Conservative) Kitab/register ke MUDE hue page ki tedhi text-lines
    seedhi karo. Kaise: page khadi pattiyon (24 strips) me bat'ta hai; har
    patti ke text-profile ko beech wali patti se milaya jaata hai
    (cross-correlation) — kitna upar/neeche khiska hai. In shifts par smooth
    (quadratic) curve fit karke HAR column apni jagah wapas khiska diya
    jaata hai (bilinear remap — koi nearest-neighbour nahi).
    Suraksha: flat page, photo, ya kam-bharose par ORIGINAL wapas."""
    try:
        if not HAS_NUMPY:
            return img
        g = np.asarray(img.convert("L"), dtype=np.float32)
        h, w = g.shape
        if h < 400 or w < 400:
            return img
        ink = np.clip(200.0 - g, 0.0, 200.0)
        NS = 24
        sw = w // NS
        if sw < 8:
            return img
        prof = np.zeros((NS, h), np.float32)
        for i in range(NS):
            prof[i] = ink[:, i * sw:(i + 1) * sw].sum(axis=1)
        tot = prof.sum(axis=1)
        c = NS // 2
        if tot[c] <= 0 or float((tot > tot[c] * 0.15).mean()) < 0.5:
            return img
        max_sh = max(6, int(h * max_shift_frac))
        ref = prof[c] - prof[c].mean()
        rn = float(np.sqrt((ref * ref).sum())) or 1.0
        shifts = np.zeros(NS, np.float32)
        conf = np.zeros(NS, np.float32)
        for i in range(NS):
            p = prof[i] - prof[i].mean()
            pn = float(np.sqrt((p * p).sum())) or 1.0
            best, bs = 0.0, 0
            for s in range(-max_sh, max_sh + 1, 2):
                if s >= 0:
                    v = float((p[s:h] * ref[0:h - s]).sum())
                else:
                    v = float((p[0:h + s] * ref[-s:h]).sum())
                v /= (pn * rn)
                if v > best:
                    best, bs = v, s
            shifts[i] = bs
            conf[i] = best
        good = conf > 0.35
        if int(good.sum()) < NS // 2:
            return img
        xs = (np.arange(NS) + 0.5) * sw
        co = np.polyfit(xs[good], shifts[good], 2)
        colshift = np.polyval(co, np.arange(w)).astype(np.float32)
        colshift = colshift - colshift[w // 2]
        mx = float(np.abs(colshift).max())
        if mx < 3.0 or mx > max_sh:
            return img          # page pehle se seedha / bharosa nahi
        rows = np.arange(h, dtype=np.float32)[:, None] + colshift[None, :]
        r0 = np.clip(np.floor(rows).astype(np.int32), 0, h - 1)
        fr = rows - r0
        r1 = np.clip(r0 + 1, 0, h - 1)
        cols = np.arange(w)[None, :]
        outg = g[r0, cols] * (1.0 - fr) + g[r1, cols] * fr
        # Quality-guard: sudhaar ke BAAD text-lines saaf zyada 'tight' honi
        # chahiye (ink-profile ka std badhe) — na badhe to original wapas
        # (seedha page kabhi kharab nahi hota).
        def _tight(arr):
            return float(np.clip(200.0 - arr, 0.0, 200.0).sum(axis=1).std())
        if _tight(outg) < 1.05 * _tight(g):
            return img
        if img.mode == "L":
            return Image.fromarray(np.clip(outg, 0, 255).astype(np.uint8), "L")
        # RGB: channel-by-channel (32-bit app me memory kam rahe)
        src = img.convert("RGB")
        outc = np.empty((h, w, 3), dtype=np.uint8)
        for ci, ch in enumerate(src.split()):
            a = np.asarray(ch, dtype=np.float32)
            outc[:, :, ci] = np.clip(a[r0, cols] * (1.0 - fr) + a[r1, cols] * fr,
                                     0, 255).astype(np.uint8)
        return Image.fromarray(outc, "RGB")
    except Exception:
        return img


def _box_mean(a, k=8):
    """k×k box-window ka mean (integral-image se — tez)."""
    c = np.cumsum(np.cumsum(a, axis=0, dtype=np.float64), axis=1)
    c = np.pad(c, ((1, 0), (1, 0)), mode="constant")
    return ((c[k:, k:] - c[:-k, k:] - c[k:, :-k] + c[:-k, :-k]) / float(k * k))


def _ssim_lite(a, b, k=8):
    """SSIM (structural similarity) — box-window version, [::2] subsample par
    chalta hai (32-bit memory-safe). 1.0 = bilkul same."""
    a = a[::2, ::2].astype(np.float64); b = b[::2, ::2].astype(np.float64)
    mu_a = _box_mean(a, k); mu_b = _box_mean(b, k)
    va = _box_mean(a * a, k) - mu_a * mu_a
    vb = _box_mean(b * b, k) - mu_b * mu_b
    cov = _box_mean(a * b, k) - mu_a * mu_b
    C1, C2 = 6.5025, 58.5225
    s = ((2 * mu_a * mu_b + C1) * (2 * cov + C2)) / ((mu_a * mu_a + mu_b * mu_b + C1) * (va + vb + C2))
    return float(s.mean())


try:
    from pyzbar import pyzbar as _sjq_zbar
except Exception:  # pragma: no cover - pyzbar optional
    _sjq_zbar = None


def smart_jpeg_quality(img, qualities=(95, 92, 88, 84, 80, 76)):
    """INTELLIGENT JPEG QUALITY: har page ke liye quality KHUD chuno —
    quality neeche utarte jao aur har baar ORIGINAL se compare karo; jis
    quality par pehli DIKHNE-LAYAK kharabi aati us se theek pehle ruk jao.

    Har trial par jaanch:
      - mean/99.5-percentile pixel difference (poore page par),
      - EDGE-MAP difference (text/patli line/QR ke kinare — sabse pehle
        yahi toot'te hain),
      - SSIM (structural similarity) >= 0.985,
      - QR/barcode page par: compress ke baad bhi decode SAME aana chahiye
        (pyzbar uplabdh ho to).
    Binary search (~3 encode/page). Kuch bhi fail ho to unchi quality par
    wapas — quality HAMESHA size se pehle. Returns (quality, report)."""
    rep = {}
    try:
        if not HAS_NUMPY:
            return 92, rep
        rgb = img if img.mode in ("L", "RGB") else img.convert("RGB")
        # Saari metrics HALF-RES par (4x tez; kinare/structure wahi pakde
        # jaate hain) — encode/decode zaroor FULL-res hota hai (asli output).
        g0 = np.asarray(rgb.convert("L"), dtype=np.float32)[::2, ::2]
        gx = np.abs(np.diff(g0, axis=1)); gy = np.abs(np.diff(g0, axis=0))
        em = np.zeros_like(g0); em[:, :-1] += gx; em[:-1, :] += gy
        strong = em > 40.0
        n_strong = int(strong.sum())
        qr0 = None
        if _sjq_zbar is not None:
            try:
                qr0 = sorted(set((d.type, bytes(d.data)) for d in _sjq_zbar.decode(rgb))) or None
            except Exception:
                qr0 = None

        import io as _io

        def _metrics(q):
            buf = _io.BytesIO()
            rgb.save(buf, "JPEG", quality=q)
            buf.seek(0)
            comp = Image.open(buf); comp.load()
            g1 = np.asarray(comp.convert("L"), dtype=np.float32)[::2, ::2]
            diff = np.abs(g1 - g0)
            mm = {"mean": float(diff.mean()),
                  "p99": float(np.percentile(diff, 99.5)),
                  "edge": (float(diff[strong].mean()) if n_strong > 100 else 0.0),
                  "ssim": _ssim_lite(g0, g1)}
            return mm, comp

        # Baseline: qualities[0] (95) khud kitna badalta hai — scanner-noise
        # wale page par SSIM noise se girta hai, content se nahi; isliye
        # aage ke sab trials 95-wale se RELATIVE naape jaate hain.
        base, _ = _metrics(qualities[0])
        # SSIM sirf LOOSE backstop hai (bahut saaf page par baseline ~0.999
        # hota hai aur chhota delta bhi jhootha alarm deta tha) — asli
        # pehredaar EDGE-difference hai.
        ssim_floor = max(0.93, base["ssim"] - 0.06)

        def _passes(q):
            mm, comp = _metrics(q)
            # PRIMARY gate = EDGE difference: text/patli line/QR ke kinare
            # (asli 'pixel phatna' yahi hai). mean/p99/ssim dheele hain —
            # wo scanner-noise ke smooth hone ko bhi 'kharabi' gin lete the.
            if mm["edge"] > 2.5:
                return False
            if mm["mean"] > 3.0 or mm["p99"] > 16.0:
                return False
            if mm["ssim"] < ssim_floor:
                return False
            if qr0 is not None:
                try:
                    if sorted(set((d.type, bytes(d.data)) for d in _sjq_zbar.decode(comp))) != qr0:
                        return False   # QR/barcode padhna band ho gaya
                except Exception:
                    return False
            return True

        ok = 0                       # qualities[0] hamesha fallback
        lo, hi = 1, len(qualities) - 1
        while lo <= hi:              # monotonic maan kar binary search
            mid = (lo + hi) // 2
            if _passes(qualities[mid]):
                ok = mid; lo = mid + 1
            else:
                hi = mid - 1
        rep["q"] = qualities[ok]; rep["edges"] = n_strong; rep["qr"] = bool(qr0)
        return qualities[ok], rep
    except Exception:
        return 92, rep


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
        sw, sh = max(1, w // 8), max(1, h // 8)
        small = img.convert("L").resize((sw, sh), Image.BILINEAR)
        # Kagaz me RANG ki tint (peela/neela) hai kya? — white-balance ke liye.
        tint = 0.0
        if img.mode != "L":
            sc = np.asarray(img.convert("RGB").resize((sw, sh), Image.BILINEAR),
                            dtype=np.float32)
            bright = np.asarray(small, dtype=np.float32) > 150
            if float(bright.mean()) > 0.2:
                ch = sc[bright].mean(axis=0)
                tint = float(ch.max() - ch.min())
        # Pehle se SAFED + bina-tint page (naye 300dpi HD scan) par kaam hi mat
        # karo — print ke waqt ye per-page bachat hai (same object return).
        if (float(np.percentile(np.asarray(small, dtype=np.float32), 80)) >= 243.0
                and tint < 12.0):
            return img

        def _bg_field(chan_small):
            # max-filter text ke strokes ke UPAR se kagaz utha leta hai,
            # blur use smooth karta hai (sab chhote size par — tez).
            f = chan_small.filter(ImageFilter.MaxFilter(9)).filter(
                ImageFilter.GaussianBlur(6)).resize((w, h), Image.BILINEAR)
            return np.maximum(np.asarray(f, dtype=np.float32), 40.0)

        b = _bg_field(small)
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
        # Quality-guard (stroke preservation): sudhaar ke baad text GHATNA
        # nahi chahiye — gehre pixel aadhe se kam reh jaayen to processing
        # reject karke original wapas (kabhi-kabhi ajeeb page par safety).
        outL = lut[np.clip(gray1, 0.0, 255.0).astype(np.uint8)]
        dark_before = float((g < 110).mean())
        if dark_before > 0.002 and float((outL < 110).mean()) < 0.5 * dark_before:
            return img
        if img.mode == "L":
            return Image.fromarray(outL, "L")
        # Colour: HAR CHANNEL ki apni background estimate — isse WHITE BALANCE
        # bhi ho jaata hai (peela/purana kagaz asli white; thanda/garam bulb ki
        # tint gayab) aur stamp/sign ke rang bane rehte hain. Channel-by-channel
        # (32-bit app me memory kam rakhne ke liye).
        chans = img.convert("RGB").split()
        outc = np.empty((h, w, 3), dtype=np.uint8)
        for i, cch in enumerate(chans):
            cs = cch.resize((sw, sh), Image.BILINEAR)
            bc = _bg_field(cs)
            ca = np.asarray(cch, dtype=np.float32)
            outc[:, :, i] = lut[np.clip(
                ca * (float(target) / bc) * cgain[None, :],
                0.0, 255.0).astype(np.uint8)]
        return Image.fromarray(outc, "RGB")
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


def naps2_clean(img):
    """NAPS2 (HP WIA driver) jaisa SAAF + CHATAK look — eSCL ka kaccha/feeka scan
    theek karo.

    NAPS2 apne scanner-driver se global white-balance + contrast + halki
    saturation lagata hai (isliye background PURA safed aur rang chatak). Ye
    function wahi karta hai:
      - autocontrast (per-channel, cutoff=1): background WHITE + poori tonal
        range (feeka-pan gaya) + halka white-balance.
      - Color (saturation) + Contrast + Sharpness: rang chatak, text crisp.

    flatten_background se ALAG: ye background ko DIVIDE nahi karta, isliye rangeen
    letterhead (gulabi/laal header, mohar) WASH-OUT nahi hota — ulta vivid hota
    hai. Asli photo/X-ray par halka rehta hai (kam sharpen, kam saturation) taaki
    over-process na ho. Grayscale page gray hi rehta hai (file chhoti)."""
    try:
        # Grayscale scan: rang nahi — sirf contrast + halka sharpen.
        if img.mode == "L":
            out = ImageOps.autocontrast(img, cutoff=1)
            return ImageEnhance.Sharpness(out).enhance(1.25)
        rgb = img.convert("RGB")
        # Document jaisa (kaafi kagaz/near-white) ya photo jaisa?
        paper = 0.5
        try:
            if HAS_NUMPY:
                L = np.asarray(rgb.convert("L"), dtype=np.float32)
                paper = float((L > 200).mean())
        except Exception:
            paper = 0.5
        rgb = ImageOps.autocontrast(rgb, cutoff=1)   # bg->white, range poori, WB
        if paper >= 0.30:                            # document: chatak
            rgb = ImageEnhance.Color(rgb).enhance(1.18)
            rgb = ImageEnhance.Contrast(rgb).enhance(1.08)
            rgb = ImageEnhance.Sharpness(rgb).enhance(1.3)
        else:                                        # photo/X-ray: halka haath
            rgb = ImageEnhance.Color(rgb).enhance(1.06)
            rgb = ImageEnhance.Sharpness(rgb).enhance(1.1)
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


def has_real_colour(img, frac=0.0006, chroma_thr=48):
    """Kya page par ASLI rang hai — chhota stamp/mohar/sign/logo bhi?
    Purana tarika (global colorfulness < 6) poore page ka AVERAGE leta tha,
    isliye 1% area wali neeli mohar bhi 'rang-heen' gin kar page GRAY ho
    jaata tha (user ki shikayat: 'colour me scan, PDF black&white').
    Ab PIXEL-level: kitne pixel saaf-saaf rangeen hain — mohar jitna chhota
    hissa (0.4%+) bhi page ko COLOUR rakh deta hai. Scanner ke halke
    colour-noise/tint se alag (chroma>40 + roshan pixel hi ginte hain)."""
    try:
        if img.mode in ("1", "L"):
            return False
        if not HAS_NUMPY:
            return colorfulness(img) >= 6.0
        a = np.asarray(img.convert("RGB"), dtype=np.float32)[::4, ::4]
        r, g, b = a[..., 0], a[..., 1], a[..., 2]
        mx = np.maximum(np.maximum(r, g), b)
        mn = np.minimum(np.minimum(r, g), b)
        sat = mx - mn                       # rang ki taakat (chroma)
        colored = (sat > float(chroma_thr)) & (mx > 80.0)
        return float(colored.mean()) >= float(frac)
    except Exception:
        return True     # shak ho to RANG rakho — mohar kabhi mat udao


def adaptive_bw(img, window=41, k=0.18):
    """Professional B&W (Sauvola adaptive threshold): har ilaake ka APNA
    threshold (local mean/std se) — global-160 wale purane tarike me patle
    akshar toot'te the aur halki syahi gayab ho jaati thi. Yahan:
      - patli lines / chhote font surakshit,
      - halka pen / stamp bhi aata hai,
      - background pure white, text deep black.
    Badi image par threshold-field aadhe size par banta hai (field smooth hota
    hai — quality same, memory/time 4x kam; 32-bit app ke liye zaroori).
    numpy na ho, ya result ajeeb lage (sab kala / sab safed), to purana global
    threshold (160) — kabhi kharab output nahi."""
    try:
        if not HAS_NUMPY:
            raise RuntimeError("no numpy")
        gsrc = img.convert("L")
        g = np.asarray(gsrc, dtype=np.float32)
        h, w = g.shape
        scale = 2 if max(h, w) > 2200 else 1
        gs = (np.asarray(gsrc.resize((max(1, w // 2), max(1, h // 2)),
                                     Image.BILINEAR), dtype=np.float32)
              if scale == 2 else g)
        hh, ww = gs.shape
        win = max(15, min(window, (min(hh, ww) // 2) * 2 - 1))
        if win % 2 == 0:
            win += 1
        pad = win // 2
        gp = np.pad(gs, pad + 1, mode="edge").astype(np.float64)
        s1 = gp.cumsum(0).cumsum(1)
        s2 = (gp * gp).cumsum(0).cumsum(1)

        def _box(S):
            return (S[win:, win:] - S[:-win, win:]
                    - S[win:, :-win] + S[:-win, :-win])

        n = float(win * win)
        mean = (_box(s1)[:hh, :ww] / n).astype(np.float32)
        var = (_box(s2)[:hh, :ww] / n).astype(np.float32) - mean * mean
        std = np.sqrt(np.maximum(var, 0.0))
        thr = mean * (1.0 + k * (std / 128.0 - 1.0))
        if scale == 2:
            thr = np.asarray(
                Image.fromarray(thr, "F").resize((w, h), Image.BILINEAR),
                dtype=np.float32)
        bw = g > thr
        black = 1.0 - float(bw.mean())
        # Sanity: bilkul khaali ya aadhe se zyada kala page = kuch galat hai.
        if not (0.002 <= black <= 0.5):
            raise RuntimeError("sauvola out of range")
        return Image.fromarray((bw.astype(np.uint8) * 255), "L").convert("1")
    except Exception:
        try:
            return img.convert("L").point(
                lambda v: 255 if v >= 160 else 0, mode="1")
        except Exception:
            return img


def straighten_photo_page(img):
    """Phone-photo me TIRCHHE rakhe page ko flatbed-jaisa seedha karo
    (perspective / keystone correction). Kagaz ka chamakta hua quad numpy se
    dhoondh kar BICUBIC warp hota hai (kabhi nearest-neighbour nahi).
    Sirf tab lagta hai jab 4 kone bharose se milen — warna original wapas
    (photo kabhi kharab nahi hoti)."""
    try:
        if not HAS_NUMPY:
            return img
        rgb = img.convert("RGB")
        w, h = rgb.size
        small = rgb.resize((max(1, w // 6), max(1, h // 6)), Image.BILINEAR)
        sg = np.asarray(small.convert("L"), dtype=np.float32)
        sh, sw = sg.shape
        # Paper-mask: andhere background par chamakta kagaz.
        thr = (float(np.percentile(sg, 20)) + float(np.percentile(sg, 90))) / 2.0
        mask = sg > thr
        frac = float(mask.mean())
        # ~poora frame bright (scanner/white table) ya bahut kam paper -> no-op.
        if not (0.25 <= frac <= 0.90):
            return img
        ys, xs = np.nonzero(mask)
        s = xs + ys
        d = xs.astype(np.int64) - ys.astype(np.int64)
        fx, fy = w / float(sw), h / float(sh)

        def _pt(i):
            return (float(xs[i]) * fx, float(ys[i]) * fy)

        tl, br = _pt(int(s.argmin())), _pt(int(s.argmax()))
        tr, bl = _pt(int(d.argmax())), _pt(int(d.argmin()))
        # Shoelace area — degenerate ya ~full-frame quad par no-op.
        qx = [tl[0], tr[0], br[0], bl[0]]
        qy = [tl[1], tr[1], br[1], bl[1]]
        area = 0.5 * abs(sum(qx[i] * qy[(i + 1) % 4] - qx[(i + 1) % 4] * qy[i]
                             for i in range(4)))
        if not (0.30 * w * h <= area <= 0.965 * w * h):
            return img

        def _dist(a, b):
            return ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2) ** 0.5

        W2 = int(max(_dist(tl, tr), _dist(bl, br)))
        H2 = int(max(_dist(tl, bl), _dist(tr, br)))
        if W2 < 200 or H2 < 200:
            return img
        # PIL QUAD source order: NW, SW, SE, NE.
        data = (tl[0], tl[1], bl[0], bl[1], br[0], br[1], tr[0], tr[1])
        return rgb.transform((W2, H2), Image.QUAD, data, Image.BICUBIC)
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
    # Tirchha page seedha (perspective warp) — bharosa na ho to no-op.
    img = straighten_photo_page(img)
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


# ---------------------------------------------------------------------------
# (v207) AI TOOLS — charan 1: sab offline, PIL+numpy, koi internet/limit nahi.
# Har function PIL image leta/deta hai; galti par original wapas (no-crash).
# ---------------------------------------------------------------------------

def ai_auto_enhance(img):
    """AI Enhance: histogram padh-kar levels (percentile stretch) + adaptive
    gamma (page ki roshni dekh kar) + halki dhaar. auto_enhance se zyada
    samajhdaar — feeke page zyada sudhrenge, achhe page par halka haath."""
    try:
        rgb = img.convert("RGB")
        if not HAS_NUMPY:
            return auto_enhance(rgb)
        lum = np.asarray(rgb.convert("L"), dtype=np.float32)
        lo = float(np.percentile(lum, 1.5))
        hi = float(np.percentile(lum, 97.0))
        if hi - lo < 10.0:
            return auto_enhance(rgb)
        arr = np.asarray(rgb, dtype=np.float32)
        arr = (arr - lo) * (255.0 / max(1.0, hi - lo))
        np.clip(arr, 0.0, 255.0, out=arr)
        # adaptive gamma: document ka mean ~200 ki taraf, par halke se (clamp)
        mean = float(arr.mean())
        if 5.0 < mean < 250.0:
            import math
            g = math.log(200.0 / 255.0) / math.log(max(2.0, mean) / 255.0)
            g = min(1.30, max(0.80, g))
            if abs(g - 1.0) > 0.02:
                arr = 255.0 * np.power(arr / 255.0, g)
        out = Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8), "RGB")
        return ImageEnhance.Sharpness(out).enhance(1.25)
    except Exception:
        return img


def ai_color_restore(img):
    """Rang wapsi: peela/purana kagaz phir safed. Kagaz (sabse chamakdaar
    hissa) ka asli rang naap kar har channel ko balance karta hai — text ke
    rang (neeli syahi, laal mohar) waise ke waise rehte hain."""
    try:
        rgb = img.convert("RGB")
        if not HAS_NUMPY:
            return ImageOps.autocontrast(rgb, cutoff=1)
        arr = np.asarray(rgb, dtype=np.float32)
        lum = arr.mean(axis=2)
        thr = max(120.0, float(np.percentile(lum, 80.0)))
        mask = lum >= thr
        if float(mask.mean()) < 0.02:
            return ImageOps.autocontrast(rgb, cutoff=1)
        paper = arr[mask].mean(axis=0)          # kagaz ka [R,G,B]
        scale = np.clip(247.0 / np.maximum(1.0, paper), 0.75, 1.60)
        out = np.clip(arr * scale, 0, 255).astype(np.uint8)
        return Image.fromarray(out, "RGB")
    except Exception:
        return img


def ai_denoise(img):
    """AI De-noise (kinara-bachau): jahan text/rekha hai wahan tez rahe, jahan
    khaali jagah hai wahan daane/dhabbe saaf. Edge-mask se original aur smooth
    ka blend — plain median se behtar (akshar dhundhle nahi hote)."""
    try:
        rgb = img.convert("RGB")
        if not HAS_NUMPY:
            return rgb.filter(ImageFilter.MedianFilter(3))
        smooth = rgb.filter(ImageFilter.MedianFilter(3)).filter(
            ImageFilter.GaussianBlur(0.8))
        edges = rgb.convert("L").filter(ImageFilter.FIND_EDGES).filter(
            ImageFilter.GaussianBlur(2.5))
        w = np.clip(np.asarray(edges, dtype=np.float32) / 255.0 * 4.0,
                    0.0, 1.0)[..., None]        # 1 = kinara => original rakho
        a = np.asarray(rgb, dtype=np.float32)
        s = np.asarray(smooth, dtype=np.float32)
        out = (a * w + s * (1.0 - w)).astype(np.uint8)
        return Image.fromarray(out, "RGB")
    except Exception:
        return img


def ai_smart_crop(img, margin_frac=0.01):
    """Smart crop: content (text/rekha) khud dhoondh kar document ki seema par
    crop — scanner lid ka kaala/safed faaltu kinara gayab. Kuch na mile ya
    shak ho to purana autocrop fallback."""
    try:
        rgb = img.convert("RGB")
        if not HAS_NUMPY:
            return autocrop(rgb)
        g = rgb.convert("L")
        sc = max(1, max(g.size) // 700)
        small = g.resize((max(1, g.width // sc), max(1, g.height // sc)))
        a = np.asarray(small, dtype=np.float32)
        border = np.concatenate([a[0, :], a[-1, :], a[:, 0], a[:, -1]])
        bg = float(np.median(border))
        diff = np.abs(a - bg)
        grad = np.zeros_like(a)
        grad[:-1, :] += np.abs(np.diff(a, axis=0))
        grad[:, :-1] += np.abs(np.diff(a, axis=1))
        m = (diff > 28.0) | (grad > 22.0)
        rows = m.mean(axis=1); cols = m.mean(axis=0)
        ri = np.where(rows > 0.015)[0]; ci = np.where(cols > 0.015)[0]
        if ri.size == 0 or ci.size == 0:
            return autocrop(rgb)
        t, b = int(ri[0]), int(ri[-1]); l, r = int(ci[0]), int(ci[-1])
        mg = int(margin_frac * max(small.size)) + 1
        t = max(0, t - mg); l = max(0, l - mg)
        b = min(small.height - 1, b + mg); r = min(small.width - 1, r + mg)
        if (r - l) * (b - t) < 0.20 * small.width * small.height:
            return autocrop(rgb)      # itna chhota? — galat pakad, safe raho
        box = (l * sc, t * sc,
               min(rgb.width, (r + 1) * sc), min(rgb.height, (b + 1) * sc))
        if box[2] - box[0] < 50 or box[3] - box[1] < 50:
            return autocrop(rgb)
        return rgb.crop(box)
    except Exception:
        return img


def ai_deskew(img):
    """Deskew+: pehle contrast normalise (feeki chhapai par bhi lines dikhen),
    phir projection-profile se 0.1° tak baareek khoj. Purane deskew se zyada
    sateek — halki 0.3-0.5° tilt bhi pakdi jaati hai. Wahi photo-guards."""
    try:
        if not HAS_NUMPY:
            return deskew(img)
        g = ImageOps.autocontrast(img.convert("L"), cutoff=2)
        if float((np.asarray(g, dtype=np.float32)[::4, ::4] > 150).mean()) < 0.55:
            return img                # photo/X-ray jaisa — mat chhedo
        small = g.resize((max(1, g.width // 4), max(1, g.height // 4)))

        def _score(angle):
            rot = small.rotate(angle, resample=Image.BILINEAR, fillcolor=255)
            arr = np.asarray(rot, dtype=np.float32)
            return float(np.var(np.diff(arr.sum(axis=1))))

        score0 = _score(0.0)
        best_c, best_cs = 0.0, score0
        for i in range(-22, 23):
            a = i * 2.0
            if a == 0.0:
                continue
            s = _score(a)
            if s > best_cs:
                best_cs, best_c = s, a
        center = best_c if (abs(best_c) > 5.0 and best_cs > 1.35 * score0) else 0.0
        best_angle, best_score = center, _score(center)
        for i in range(-10, 11):
            a = center + i * 0.5
            s = _score(a)
            if s > best_score:
                best_score, best_angle = s, a
        c2 = best_angle
        for i in range(-5, 6):
            a = c2 + i * 0.1
            if abs(a - best_angle) < 1e-6:
                continue
            s = _score(a)
            if s > best_score:
                best_score, best_angle = s, a
        if abs(best_angle) < 0.15 or best_score < 1.12 * score0:
            return img
        return img.rotate(best_angle, resample=Image.BICUBIC,
                          fillcolor=(255, 255, 255), expand=True)
    except Exception:
        return img


def ai_inpaint_region(img, box):
    """(v209 charan 2) Daag/ungli mitao — chuna hua box aas-paas ke rang se
    bhar jaata hai (multi-scale harmonic/diffusion inpainting, sirf numpy).
    Documents par ungli ka kona, staple/punch ke nishaan, daag ke liye.
    Smooth background par bilkul gayab; texture par halka soft fill."""
    try:
        rgb = img.convert("RGB")
        if not HAS_NUMPY:
            return rgb
        x0, y0, x1, y1 = [int(v) for v in box]
        x0 = max(0, x0); y0 = max(0, y0)
        x1 = min(rgb.width, x1); y1 = min(rgb.height, y1)
        if x1 - x0 < 3 or y1 - y0 < 3:
            return rgb
        # patch = box + aas-paas ka ring (source pixels)
        pad = max(10, (x1 - x0 + y1 - y0) // 6)
        px0 = max(0, x0 - pad); py0 = max(0, y0 - pad)
        px1 = min(rgb.width, x1 + pad); py1 = min(rgb.height, y1 + pad)
        patch = rgb.crop((px0, py0, px1, py1))
        a = np.asarray(patch, dtype=np.float32).copy()
        m = np.zeros(a.shape[:2], dtype=bool)
        hy0, hy1 = y0 - py0, y1 - py0
        hx0, hx1 = x0 - px0, x1 - px0
        m[hy0:hy1, hx0:hx1] = True
        ring = a[~m]
        if ring.size == 0:
            return rgb
        # 4-kinaron se interpolation: text-line jo hole ke aar-paar jaati hai
        # wo left↔right interpolation me bach jaati hai (sirf diffusion me
        # dhul jaati thi). Kinare ke paas uska weight zyada (inverse-distance).
        hh, hw = hy1 - hy0, hx1 - hx0
        Lv = a[hy0:hy1, hx0 - 1] if hx0 > 0 else None          # left boundary col
        Rv = a[hy0:hy1, hx1] if hx1 < a.shape[1] else None      # right boundary col
        Tv = a[hy0 - 1, hx0:hx1] if hy0 > 0 else None           # top boundary row
        Bv = a[hy1, hx0:hx1] if hy1 < a.shape[0] else None      # bottom boundary row
        cc = np.arange(hw, dtype=np.float32)
        rr = np.arange(hh, dtype=np.float32)
        tx = ((cc + 1.0) / (hw + 1.0))[None, :, None]
        ty = ((rr + 1.0) / (hh + 1.0))[:, None, None]
        if Lv is not None and Rv is not None:
            horiz = Lv[:, None, :] * (1.0 - tx) + Rv[:, None, :] * tx
        else:
            horiz = (Lv if Lv is not None else Rv)
            horiz = None if horiz is None else np.repeat(horiz[:, None, :], hw, axis=1)
        if Tv is not None and Bv is not None:
            vert = Tv[None, :, :] * (1.0 - ty) + Bv[None, :, :] * ty
        else:
            vert = (Tv if Tv is not None else Bv)
            vert = None if vert is None else np.repeat(vert[None, :, :], hh, axis=0)
        if horiz is None and vert is None:
            a[m] = np.median(ring.reshape(-1, 3), axis=0)
        elif vert is None:
            fill = horiz
            if Lv is not None and Rv is not None and Tv is None and Bv is None:
                pass                      # horiz hi sab kuch carry karta hai
            a[hy0:hy1, hx0:hx1] = fill
        else:
            # base = upar-neeche ka smooth interpolation (khadi rekha/feature
            # apne aap columns ke raaste carry hoti hai)
            fill = vert.copy()
            if Lv is not None and Rv is not None:
                # left-right RESIDUAL carry: aar-paar jaati text-line/rekha
                # (jo base me nahi hai) poori taakat se hole ke paar jaaye
                rL = (Lv - fill[:, 0, :])[:, None, :]
                rR = (Rv - fill[:, -1, :])[:, None, :]
                fill = fill + rL * (1.0 - tx) + rR * tx
            a[hy0:hy1, hx0:hx1] = np.clip(fill, 0.0, 255.0)
        # bahut halki diffusion — sirf seams naram (structure na dhule)
        for _ in range(8):
            p = np.pad(a, ((1, 1), (1, 1), (0, 0)), mode="edge")
            nb = (p[:-2, 1:-1] + p[2:, 1:-1] + p[1:-1, :-2] + p[1:-1, 2:]) / 4.0
            a[m] = a[m] * 0.70 + nb[m] * 0.30
        out = rgb.copy()
        out.paste(Image.fromarray(np.clip(a, 0, 255).astype(np.uint8)), (px0, py0))
        return out
    except Exception:
        return img


def ai_scan_quality(img):
    """(v211) Scan-guard chowkidaar: page ki turant jaanch. Kharabi-codes ki
    list lautata hai: 'dark' (bahut gehra/kaala), 'faint' (bahut feeka),
    'blur' (dhundhla). Khaali list = sab theek. Thresholds conservative —
    jhoothi chetavni se achha hai kabhi-kabhi chup rehna."""
    issues = []
    try:
        g = img.convert("L")
        small = g.resize((max(1, min(420, g.width)), max(1, min(420, g.height))))
        if not HAS_NUMPY:
            return issues
        a = np.asarray(small, dtype=np.float32)
        mean = float(a.mean())
        dyn0 = float(np.percentile(a, 97) - np.percentile(a, 3))
        if mean < 85:
            issues.append("dark")
        elif mean > 232 and 6.0 < dyn0 < 80.0:
            # feeka CONTENT (khaali/blank page nahi — wo alag filter ka kaam)
            issues.append("faint")
        # dhundhlapan: sabse tez kinare (p99.9 gradient) bhi kamzor hain PAR
        # page par content zaroor hai (dyn) — saaf text par p99.9 ~100+,
        # dhundhle par ~10-30. Mean-gradient sparse text par dhokha deta tha.
        dx = np.abs(np.diff(a, axis=1))
        p999 = float(np.percentile(dx, 99.9))
        dyn = float(np.percentile(a, 97) - np.percentile(a, 3))
        if p999 < 30.0 and dyn > 60.0:
            issues.append("blur")
    except Exception:
        pass
    return issues


def ai_photo_to_scan(img):
    """(v211) Photo → scan jaisa, EK click/apne aap: EXIF seedha + size cap +
    4-kone perspective + shadow flatten (clean_photo) + rang wapsi +
    AI enhance + kinara-bachau denoise."""
    try:
        im = clean_photo(img)
        im = ai_color_restore(im)
        im = ai_auto_enhance(im)
        im = ai_denoise(im)
        return im.convert("RGB")
    except Exception:
        try:
            return img.convert("RGB")
        except Exception:
            return img


def ai_extract_signature(img):
    """(v210 charan 3) Sign/mohar nikaalo: kagaz ka rang naap kar background
    poori tarah TRANSPARENT — sirf syahi (sign/mohar) bachti hai, asli rang
    aur naram kinaron (anti-aliased alpha) ke saath. RGBA lautata hai;
    kuch na mile to None."""
    try:
        rgb = img.convert("RGB")
        if not HAS_NUMPY:
            return None
        a = np.asarray(rgb, dtype=np.float32)
        lum = a.mean(axis=2)
        # kagaz ka rang = chamakdaar 40% pixels ka median (sign chhota hota
        # hai, zyada hissa kagaz hi hota hai)
        thr = float(np.percentile(lum, 60.0))
        paper_px = a[lum >= thr]
        if paper_px.size == 0:
            return None
        paper = np.median(paper_px.reshape(-1, 3), axis=0)
        dist = np.sqrt(((a - paper) ** 2).sum(axis=2))       # 0 = kagaz
        # naram alpha-ramp: 18 tak kagaz (0), 55+ pakki syahi (255)
        alpha = np.clip((dist - 18.0) * (255.0 / (55.0 - 18.0)), 0.0, 255.0)
        alpha_im = Image.fromarray(alpha.astype(np.uint8), "L")
        alpha_im = alpha_im.filter(ImageFilter.MedianFilter(3))   # akela daana ud jaye
        alpha = np.asarray(alpha_im, dtype=np.uint8)
        if float((alpha > 128).mean()) < 0.0015:              # kuch mila hi nahi
            return None
        out = np.dstack([a.astype(np.uint8), alpha])
        res = Image.fromarray(out, "RGBA")
        # content par trim (6px hashiya)
        bbox = alpha_im.point(lambda v: 255 if v > 24 else 0).getbbox()
        if bbox:
            l, t, r, b = bbox
            res = res.crop((max(0, l - 6), max(0, t - 6),
                            min(res.width, r + 6), min(res.height, b + 6)))
        return res
    except Exception:
        return None


def ai_auto_all(img):
    """AI Auto (sab ek saath): Deskew+ → Smart crop → Rang wapsi →
    AI Enhance → AI De-noise. Ek click me poora page taiyaar."""
    try:
        im = ai_deskew(img)
        im = ai_smart_crop(im)
        im = ai_color_restore(im)
        im = ai_auto_enhance(im)
        im = ai_denoise(im)
        return im.convert("RGB")
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
