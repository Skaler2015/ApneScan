"""Smart Document Orientation engine — multi-method auto-rotate (OCR + numpy).

Extracted verbatim from apnescan.py. Detects the correct 0/90/180/270 page
orientation via Tesseract OSD with numpy text-line + layout fallbacks, and
returns an upright image (lossless quarter-turns, DPI preserved). Low confidence
keeps the original. Behaviour is byte-for-byte identical to the original code.

OCR decoupling: this module needs the app's Tesseract-availability check but must
not import apnescan.py (that would be circular). It ships a self-contained
fallback and apnescan.py injects its own cached tesseract_available() over it at
import time, so the exact same result is used everywhere.
"""

import os
import re
import datetime

from PIL import ImageOps

try:
    import numpy as np
    HAS_NUMPY = True
except Exception:  # pragma: no cover
    HAS_NUMPY = False

try:
    import pytesseract
    HAS_OCR_LIBS = True
except Exception:  # pragma: no cover
    HAS_OCR_LIBS = False

from apnescan_lib.imaging import is_blank_page, deskew

__all__ = ["auto_orient", "detect_orientation", "ORIENT_LOG", "tesseract_available"]


def _default_tesseract_available():
    """Fallback OCR-availability check. apnescan.py injects its own cached
    tesseract_available() over this name at import time, so orientation uses the
    identical result the rest of the app uses — with no circular import."""
    if not HAS_OCR_LIBS:
        return False
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


# Rebound by apnescan.py to its own cached tesseract_available() after import.
tesseract_available = _default_tesseract_available


def _orient_keep_dpi(src, out):
    """Carry the original DPI/format hints onto the rotated image (no quality loss)."""
    try:
        if src is not out and src.info.get("dpi"):
            out.info["dpi"] = src.info["dpi"]
    except Exception:
        pass
    return out


def _orient_osd(img):
    """Tesseract OSD orientation on a downscaled copy (fast) — returns
    (rotate_degrees, confidence, script). Detection is on a small copy; the
    rotation itself is applied to the full-res original elsewhere."""
    try:
        rgb = img.convert("RGB")
        w, h = rgb.size
        m = max(w, h)
        if m > 1000:                      # detect small, rotate original -> fast + lossless
            s = 1000.0 / m
            rgb = rgb.resize((max(1, int(w * s)), max(1, int(h * s))))
        try:                              # boost contrast on the small copy so OSD copes
            rgb = ImageOps.autocontrast(rgb, cutoff=2)   # with dark/light/old docs
        except Exception:
            pass
        try:
            from pytesseract import Output as _Out
            d = pytesseract.image_to_osd(rgb, output_type=_Out.DICT)
            return (int(d.get("rotate", 0) or 0),
                    float(d.get("orientation_conf", 0) or 0), str(d.get("script", "")))
        except Exception:
            osd = pytesseract.image_to_osd(rgb)
            mr = re.search(r"Rotate:\s*(\d+)", osd or "")
            mc = re.search(r"Orientation confidence:\s*([\d.]+)", osd or "")
            return (int(mr.group(1)) if mr else 0,
                    float(mc.group(1)) if mc else 0.0, "")
    except Exception:
        return (0, 0.0, "")


def _orient_osd_pct(raw):
    """Map Tesseract OSD's raw orientation_conf (~0..15) to a 0-100% scale for
    the threshold + logging. raw ~2.7 -> ~80%."""
    try:
        return int(min(99, max(0, round(float(raw) * 28 + 5))))
    except Exception:
        return 0


def _orient_text(img):
    """METHOD 2 — text-line structure analysis (OpenCV-style, done with numpy so
    no cv2 dependency). Counts horizontal vs vertical text 'bands' in the ink
    projection profiles to tell an UPRIGHT/upside-down page (0/180 axis) from a
    SIDEWAYS one (90/270). Returns (axis_angle 0|90, confidence%). It cannot tell
    up-from-down, so 90 here just means 'a quarter-turn is needed'."""
    if not HAS_NUMPY:
        return (0, 0)
    try:
        g = img.convert("L").resize((400, 400))
        a = np.asarray(g, dtype=np.float32)
        thr = max(60.0, a.mean() - 0.5 * a.std())      # adaptive ink threshold
        ink = (a < thr).astype(np.float32)
        if ink.sum() < 120:                            # too little text -> no signal
            return (0, 0)

        def bands(p):                                  # how many text/gap alternations
            p = p - p.mean()
            return int(np.count_nonzero(np.diff(np.sign(p)) != 0))
        rb = bands(ink.sum(axis=1))                    # rows oscillate -> horizontal lines
        cb = bands(ink.sum(axis=0))                    # cols oscillate -> vertical lines
        tot = rb + cb
        if tot < 6:
            return (0, 0)
        if rb >= cb:                                   # horizontal text -> upright axis
            return (0, int(min(72, round(60.0 * rb / tot))))
        return (90, int(min(72, round(60.0 * cb / tot))))   # vertical text -> sideways
    except Exception:
        return (0, 0)


def _orient_layout(img):
    """METHOD 3 — document layout (margin / white-space distribution). A weaker
    tie-breaker: strong ink imbalance across one axis suggests the page is
    sideways. Returns (angle 0|90, confidence%)."""
    if not HAS_NUMPY:
        return (0, 0)
    try:
        g = np.asarray(img.convert("L").resize((240, 240)), dtype=np.float32)
        ink = (g < 150).astype(np.float32)
        if ink.sum() < 60:
            return (0, 0)
        rv = float(np.var(ink.sum(axis=1)))            # per-row ink variance
        cv = float(np.var(ink.sum(axis=0)))            # per-col ink variance
        if rv <= 0 and cv <= 0:
            return (0, 0)
        if cv > rv * 1.7:
            return (90, int(min(50, round(28 * cv / (rv + 1e-6)))))
        return (0, int(min(40, round(18 * rv / (cv + 1e-6)))))
    except Exception:
        return (0, 0)


def detect_orientation(img, use_ocr=True, use_text=True, use_layout=True, threshold=80):
    """Smart multi-method orientation detection (offline, priority-ordered).
    Returns {'angle':0/90/180/270, 'conf':0-100, 'method':'osd'/'text'/'layout'/'none',
    'decision':'rotate'/'keep'}. Rotates ONLY when confidence >= threshold, so an
    uncertain page is never turned the wrong way."""
    res = {"angle": 0, "conf": 0, "method": "none", "decision": "keep"}
    # METHOD 1 — Tesseract OSD (only one that resolves full 0/90/180/270 + up/down)
    if use_ocr and tesseract_available():
        try:
            rot, raw, _script = _orient_osd(img)
            res = {"angle": (rot if rot in (90, 180, 270) else 0),
                   "conf": _orient_osd_pct(raw), "method": "osd", "decision": "keep"}
        except Exception:
            pass
    # METHODS 2 & 3 — fallbacks when OSD is unavailable or not confident enough
    if res["method"] == "none" or res["conf"] < threshold:
        for use, fn, name in ((use_text, _orient_text, "text"),
                              (use_layout, _orient_layout, "layout")):
            if not use:
                continue
            try:
                a, c = fn(img)
            except Exception:
                a, c = 0, 0
            if a in (90,) and c > res["conf"]:
                res = {"angle": a, "conf": c, "method": name, "decision": "keep"}
    # METHOD 4 — smart decision: only rotate when reliably confident
    res["decision"] = "rotate" if (res["angle"] in (90, 180, 270) and res["conf"] >= threshold) else "keep"
    return res


ORIENT_LOG = os.path.join(os.path.expanduser("~"), ".apnescan_orient.log")


def _orient_log(label, res):
    """Log one orientation decision (page, angle, confidence, method, decision).
    Self-capping file; never raises."""
    try:
        line = "%s\t%s\tangle=%s conf=%s%% method=%s -> %s\n" % (
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), (label or "page"),
            res.get("angle"), res.get("conf"), res.get("method"), res.get("decision"))
        if os.path.exists(ORIENT_LOG) and os.path.getsize(ORIENT_LOG) > 300000:
            try:
                with open(ORIENT_LOG, "r", encoding="utf-8") as f:
                    tail = f.readlines()[-1500:]
                with open(ORIENT_LOG, "w", encoding="utf-8") as f:
                    f.writelines(tail)
            except Exception:
                pass
        with open(ORIENT_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def auto_orient(img, mode="accurate", ocr_based=True, layout_based=True,
                deskew_small=False, min_conf=None, threshold=None,
                text_based=None, page_label=""):
    """Intelligent Smart-Orientation engine (offline). Detects the correct
    0/90/180/270 orientation via Tesseract OSD (English/Hindi/mixed) with numpy
    text-line + layout fallbacks, and returns an UPRIGHT image. 90° turns are
    lossless (no resize/recompress); DPI + colour depth preserved. Low confidence
    -> ORIGINAL kept (never a wrong rotation). Blank pages skip OCR. Every decision
    is logged. Backward compatible: auto_orient(img) still works.

    'mode' fast = OSD only (quickest); accurate = OSD + text + layout."""
    try:
        orig = img
        if threshold is None:                          # legacy min_conf -> %; else default 80
            threshold = 80
        if text_based is None:
            text_based = (mode != "fast")
        if mode == "fast":
            layout_based = False                       # keep fast mode lean
        # blank page -> no reliable signal; keep as-is
        try:
            if is_blank_page(orig.convert("RGB"), 0.0008):
                if page_label:
                    _orient_log(page_label, {"angle": 0, "conf": 0, "method": "blank", "decision": "keep"})
                return orig
        except Exception:
            pass
        res = detect_orientation(orig, use_ocr=ocr_based, use_text=text_based,
                                 use_layout=layout_based, threshold=int(threshold))
        out = orig
        if res["decision"] == "rotate" and res["angle"] in (90, 180, 270):
            out = orig.rotate(-res["angle"], expand=True)   # lossless quarter-turn
        if deskew_small:
            out = deskew(out)
        _orient_log(page_label, res)
        return _orient_keep_dpi(orig, out)
    except Exception:
        return img
