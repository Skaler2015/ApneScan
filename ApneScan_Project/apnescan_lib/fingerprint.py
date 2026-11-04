"""Document fingerprint engine — text/image/barcode signatures for AI Memory.

Extracted verbatim from apnescan.py. Pure, stateless helpers the SQLite-backed
Document Memory uses to recognise "the same document" across scans: normalised
text keywords + hash, Jaccard/containment text similarity, aHash/dHash/pHash
image fingerprints (numpy DCT), barcode/QR decoding (pyzbar) and filename
subject/affix analysis. No UI, no database access, no app state — behaviour is
byte-for-byte identical to the original definitions.
"""

import re
import hashlib

from PIL import Image

try:
    import numpy as np
    HAS_NUMPY = True
except Exception:  # pragma: no cover
    HAS_NUMPY = False

try:
    from pyzbar.pyzbar import decode as _zbar_decode
    HAS_ZBAR = True
except Exception:  # pragma: no cover
    HAS_ZBAR = False


_DM_STOP = set((
    "the a an of to in on for and or is are was at by be with from as it this that "
    "ka ki ke ko hai ho na ne me se aur ya ek do par bhi "
).split())
_DM_WORD = re.compile(r"[A-Za-z0-9ऀ-ॿ]+")
_DM_DCT_CACHE = {}


def _dm_norm(t):
    """Lowercase, keep alnum + Devanagari, collapse to single spaces."""
    if not t:
        return ""
    return " ".join(_DM_WORD.findall(t.lower()))


def _dm_tokens(t):
    return [w for w in _dm_norm(t).split() if len(w) >= 2]


def _dm_keywords(t, k=18):
    """Top keywords by frequency (stopwords removed, len >= 3)."""
    freq = {}
    for w in _dm_tokens(t):
        if len(w) < 3 or w in _DM_STOP:
            continue
        freq[w] = freq.get(w, 0) + 1
    top = sorted(freq.items(), key=lambda x: (-x[1], x[0]))[:k]
    return [w for w, _ in top]


def _dm_text_hash(t):
    n = _dm_norm(t)
    return hashlib.sha1(n.encode("utf-8")).hexdigest()[:16] if n else ""


def _dm_jaccard(a, b):
    if not a or not b:
        return 0.0
    sa, sb = set(a), set(b)
    u = sa | sb
    return (len(sa & sb) / float(len(u))) if u else 0.0


def _dm_text_sim(a, b):
    """0..100 similarity of two OCR texts (token Jaccard + bigram shingles).
    Robust to OCR noise / word reordering. Inputs capped to ~500 tokens."""
    ta, tb = _dm_tokens(a)[:500], _dm_tokens(b)[:500]
    if not ta or not tb:
        return 0.0
    jac = _dm_jaccard(ta, tb)
    ba, bb = set(zip(ta, ta[1:])), set(zip(tb, tb[1:]))
    shin = _dm_jaccard(ba, bb) if (ba and bb) else jac
    return round(100.0 * (0.6 * jac + 0.4 * shin), 1)


def _dm_gray(img, size):
    rs = getattr(Image, "LANCZOS", getattr(Image, "BILINEAR", 2))
    return np.asarray(img.convert("L").resize((size, size), rs), dtype=np.float64)


def _dm_bits_to_hex(bits):
    v = 0
    for b in bits:
        v = (v << 1) | (1 if b else 0)
    return "%016x" % v


def _dm_ahash(img):
    """Average hash (64-bit / 16 hex)."""
    if not HAS_NUMPY:
        return ""
    try:
        a = _dm_gray(img, 8)
        return _dm_bits_to_hex((a > a.mean()).flatten())
    except Exception:
        return ""


def _dm_dhash(img):
    """Difference hash (horizontal gradient, 64-bit)."""
    if not HAS_NUMPY:
        return ""
    try:
        rs = getattr(Image, "LANCZOS", getattr(Image, "BILINEAR", 2))
        a = np.asarray(img.convert("L").resize((9, 8), rs), dtype=np.int32)
        return _dm_bits_to_hex((a[:, 1:] > a[:, :-1]).flatten())
    except Exception:
        return ""


def _dm_dct_matrix(N):
    m = _DM_DCT_CACHE.get(N)
    if m is not None:
        return m
    n = np.arange(N)
    k = n.reshape(-1, 1)
    M = np.cos(np.pi * (2 * n + 1) * k / (2.0 * N))
    M[0, :] *= 1.0 / np.sqrt(2)
    M *= np.sqrt(2.0 / N)
    _DM_DCT_CACHE[N] = M
    return M


def _dm_phash(img):
    """Perceptual hash: 2D-DCT of a 32x32 image, low-freq 8x8 block (64-bit)."""
    if not HAS_NUMPY:
        return ""
    try:
        a = _dm_gray(img, 32)
        M = _dm_dct_matrix(32)
        d = M.dot(a).dot(M.T)
        low = d[:8, :8].copy()
        low[0, 0] = 0.0
        return _dm_bits_to_hex((low > np.median(low)).flatten())
    except Exception:
        return ""


def _dm_hamming(h1, h2):
    if not h1 or not h2 or len(h1) != len(h2):
        return 64
    try:
        return bin(int(h1, 16) ^ int(h2, 16)).count("1")
    except Exception:
        return 64


def _dm_hash_sim(h1, h2):
    """0..100 similarity for 64-bit hex hashes."""
    if not h1 or not h2:
        return 0.0
    return round(100.0 * (1.0 - _dm_hamming(h1, h2) / 64.0), 1)


def _dm_decode_codes(img):
    """Return (barcode, qrcode) text if present. Best-effort (needs pyzbar).
    Never raises."""
    if not HAS_ZBAR:
        return "", ""
    bc, qr = "", ""
    try:
        for r in _zbar_decode(img):
            try:
                data = r.data.decode("utf-8", "ignore")
            except Exception:
                data = str(r.data)
            if not data:
                continue
            if getattr(r, "type", "") == "QRCODE":
                qr = qr or data[:200]
            else:
                bc = bc or data[:200]
    except Exception:
        pass
    return bc, qr


# ---- Naming-pattern learning (offline) --------------------------------------
# Learns the fixed part of how the user names a document TYPE (e.g. every Aadhaar
# is "<person> - Aadhaar") and applies it to a brand-new, never-seen document by
# pulling the person/subject out of its OCR text. Pure functions -> easy to test.
_DM_NONNAME = set((
    "government india republic name naam patient holder card number no date dob "
    "aadhaar aadhar pan invoice bill receipt hospital ward summary discharge "
    "prescription report lab claim rghs echs the of male female address from amount "
    "shri smt kumari doctor dr mr mrs total gst"
).split())


def _dm_clean_name(s):
    """Keep up to 3 real name-words (drop labels/keywords), Title-Cased."""
    words = [w for w in re.findall(r'[A-Za-z]+', s or "")
             if len(w) >= 2 and w.lower() not in _DM_NONNAME][:3]
    if len(words) < 2:
        return ""
    return " ".join(w[:1].upper() + w[1:].lower() for w in words)


def _dm_guess_subject(text):
    """Best-effort person/subject from OCR text (label-guided, else first clean
    Title-Case run). Offline, heuristic — never raises."""
    if not text:
        return ""
    try:
        m = re.search(r'(?:patient\s+name|holder\s+name|\bname\b|\bnaam\b)\s*[:\-]\s*', text, re.I)
        if m:
            m2 = re.match(r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,2})', text[m.end():m.end() + 40])
            if m2:
                c = _dm_clean_name(m2.group(1))
                if c:
                    return c
        for m in re.finditer(r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})\b', text):
            c = _dm_clean_name(m.group(1))
            if c:
                return c
    except Exception:
        pass
    return ""


def _dm_affix_pattern(filenames, min_samples=3):
    """From several filenames of ONE doc-type, find the shared leading/trailing
    text around the variable slot. Returns (prefix, suffix) or None."""
    fns = [re.sub(r'[_]+', ' ', f).strip() for f in filenames if f and f.strip()]
    fns = [f for f in fns if f]
    if len(fns) < min_samples:
        return None

    def lcp(strs):
        s0 = strs[0]
        n = min(len(s) for s in strs)
        i = 0
        while i < n and all(s[i].lower() == s0[i].lower() for s in strs):
            i += 1
        return s0[:i]
    pre = lcp(fns)
    suf = lcp([f[::-1] for f in fns])[::-1]
    # snap to a word/separator boundary so a name isn't sliced mid-word
    if pre and not pre.endswith((' ', '-', '.', '/')):
        cut = max(pre.rfind(' '), pre.rfind('-'))
        pre = pre[:cut + 1] if cut >= 0 else ""
    if suf and not suf.startswith((' ', '-', '.', '/')):
        cand = [x for x in (suf.find(' '), suf.find('-')) if x >= 0]
        suf = suf[min(cand):] if cand else ""
    if len(pre.strip()) + len(suf.strip()) < 2:
        return None
    if any(len(pre) + len(suf) >= len(f) for f in fns):   # no room for a variable
        return None
    return (pre, suf)
