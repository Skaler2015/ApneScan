"""OCR text-recognition helpers — page titles, doc-type classify, name learning.

Extracted verbatim from apnescan.py. Reads a page's OCR text (top fraction),
classifies WHAT the document is (DOC_TYPES keyword table, tuned for ECHS/RGHS
hospital paperwork), extracts document numbers, and matches pages to previously
learned names by word-signature (Jaccard/containment) or visual signature.
Behaviour is byte-for-byte identical to the original definitions.

pytesseract note: this module imports the same pytesseract module object the app
configures at startup (_configure_tesseract sets pytesseract.pytesseract
.tesseract_cmd on the shared singleton), so OCR behaves exactly the same here.
"""

import re

from PIL import Image

try:
    import pytesseract
    HAS_OCR_LIBS = True
except Exception:  # pragma: no cover
    HAS_OCR_LIBS = False


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
    """Kisi form ki FIXED chhapi layout (letterhead, field-labels, title) ko
    pehchanne wale khaas shabd. Patient/customer-specific data (naam, number)
    zyadatar nikal jata hai, isliye SAME form ke do scan inme se zyadatar shabd
    share karte hain.

    Ab ENGLISH + HINDI (Devanagari) DONO pakadta hai — pehle sirf English tha,
    isliye Hindi forms ka naam theek se yaad nahi rehta tha."""
    low = (text or "").lower()
    en = re.findall(r"[a-z]{3,}", low)          # 3+ (pehle 4+ tha)
    hi = re.findall(r"[ऀ-ॿ]{2,}", text or "")   # Devanagari shabd
    stop = {"name", "date", "time", "age", "years", "year", "male", "female",
            "address", "sign", "signature", "page", "self", "spouse", "mobile",
            "phone", "number", "the", "and", "for", "with", "this", "that",
            "नाम", "दिनांक", "पता", "उम्र", "हस्ताक्षर", "मोबाइल", "फोन"}
    return set(w for w in (en + hi) if w not in stop)


def _jaccard(a, b):
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / union if union else 0.0


def _containment(sig, stored):
    """Stored form-signature ka kitna hissa NAYE page me maujood hai (0..1).
    Ye Jaccard se behtar hai jab naye scan me zyada/kam text ho — kyunki ek hi
    form dobara scan karne par uska poora signature naye me aa hi jata hai."""
    if not sig or not stored:
        return 0.0
    return len(sig & stored) / float(len(stored))


def learned_name_for(text, learned, thr=0.45):
    """learned = list of (wordset, name). Jis stored signature se ye page sabse
    zyada milta hai uska naam do (agar kaafi milta ho). Jaccard YA containment —
    dono me se jo behtar ho use lete hain (thr thoda narm bhi kiya hai)."""
    sig = sig_words(text)
    if not sig:
        return None
    best_name, best_score = None, 0.0
    for words, name in learned:
        ws = words if isinstance(words, set) else set(words)
        # containment ko halka zyada wazan (same form dobara aane par ~1.0 hota hai)
        sc = max(_jaccard(sig, ws), _containment(sig, ws) * 0.95)
        if sc > best_score:
            best_score, best_name = sc, name
    # thr=0.40 (narm) — par containment 0.6+ ho to bhi maan lo
    return best_name if best_score >= 0.40 else None


# ---------------------------------------------------------------------------
# Visual (shakl se) name-matching — OCR ki zaroorat NAHI. Page ki design/
# layout + rang dekh kar pehle diye gaye naam se milaan. Handwritten/X-ray/
# ID jaise pages ke liye (jahan OCR fail hota hai) sabse kaam ka.
# ---------------------------------------------------------------------------
def visual_signature(img):
    """Page ki 'shakl' ka fingerprint (OCR nahi). v2 — zyada barik taaki
    ek jaise letterhead wale alag document (Referral vs Lab-report) bhi
    thoda alag pehchane jaayein:
      - gray 24x24 average-hash bits  -> layout/design (576 bits)
      - dark-pixel %                   -> ghana table vs khaali vs kaali X-ray
      - colour 8x8 grid ke RGB         -> rang / type (blue ink, dark X-ray…)
      - aspect-ratio                   -> portrait/landscape"""
    try:
        im = img.convert("RGB")
        g = im.convert("L").resize((24, 24))
        px = list(g.getdata())
        avg = (sum(px) / len(px)) if px else 128
        bits = "".join("1" if p >= avg else "0" for p in px)       # 576 bits
        dark = round(sum(1 for p in px if p < 100) / float(len(px)), 3)
        c = im.resize((8, 8))
        color = [v for rgb in c.getdata() for v in (rgb if isinstance(rgb, tuple) else (rgb, rgb, rgb))]
        w, h = im.size
        ar = round(w / float(h), 2) if h else 1.0
        return {"v": 2, "bits": bits, "dark": dark, "color": color, "ar": ar}
    except Exception:
        return None


def _vsig_sim(a, b):
    """Do visual signature kitne milte hain (0..1)."""
    if not a or not b:
        return 0.0
    ba, bb = a.get("bits", ""), b.get("bits", "")
    if not ba or len(ba) != len(bb):
        return 0.0                                    # alag version/size -> match nahi
    lay = sum(1 for x, y in zip(ba, bb) if x == y) / float(len(ba))
    ca, cb = a.get("color", []), b.get("color", [])
    col = 1.0 - (sum(abs(x - y) for x, y in zip(ca, cb)) / (len(ca) * 255.0)) \
        if (ca and len(ca) == len(cb)) else 0.0
    dsim = 1.0 - abs(a.get("dark", 0) - b.get("dark", 0))          # ink-density milaan
    asp = 1.0 if abs(a.get("ar", 1) - b.get("ar", 1)) < 0.18 else 0.6
    return (0.52 * lay + 0.28 * col + 0.20 * dsim) * asp


def visual_name_for(sig, learned_visual, thr=0.86, margin=0.09):
    """Sabse milta-julta naam do — PAR sirf tab jab match SAAF ho. Shakl-only
    matching sirf tab bharosemand hai jab document bilkul alag ho (ID/X-ray)
    ya wahi page dobara ho; ek jaise letterhead wale docs ke liye ye jaan-
    boojhkar 'koi naam nahi' deta (unhe TEXT/OCR se naam milta hai):
      - best score >= thr, AUR
      - ya to best >= 0.94 (yaani wahi document dobara), ya best se alag NAAM
        wala doosra signature kaafi peeche ho (margin). Warna koi naam nahi."""
    if not sig:
        return None, 0.0
    scored = []
    for e in (learned_visual or []):
        scored.append((_vsig_sim(sig, e.get("vsig")), e.get("name") or ""))
    if not scored:
        return None, 0.0
    scored.sort(key=lambda t: t[0], reverse=True)
    best_sc, best_name = scored[0]
    if best_sc < thr:
        return None, best_sc
    if best_sc >= 0.94:                    # लगभग वही page दोबारा — pakka
        return best_name, best_sc
    for sc, nm in scored[1:]:              # pehla ALAG-naam wala competitor
        if nm and nm != best_name:
            if best_sc - sc < margin:
                return None, best_sc       # do naam bahut paas -> ambiguous, chhodo
            break
    return best_name, best_sc


def extract_doc_number(text):
    """Page ke text me se sabse sambhavit 'document number' (invoice/bill/claim/
    receipt no.) nikaalo. 'No/No./Number/Bill/Invoice/Claim/Receipt' ke aage wala
    number pehle dekha jata hai; na mile to koi lamba alphanumeric code."""
    t = text or ""
    m = re.search(r"(?:invoice|bill|claim|receipt|ref(?:erence)?|no|number|"
                  r"बिल|रसीद|क्लेम|संख्या|नंबर)\.?\s*(?:संख्या|नंबर|no|number)?\.?"
                  r"\s*[:#-]?\s*([A-Za-z]{0,4}\d[\w\-/]{2,20})",
                  t, re.IGNORECASE)
    if m:
        return re.sub(r"[^A-Za-z0-9\-]", "", m.group(1))[:24]
    m = re.search(r"\b([A-Z]{2,5}[-/]?\d{3,})\b", t)   # jaise INV-1234, ABC/2026/45
    if m:
        return re.sub(r"[^A-Za-z0-9\-]", "", m.group(1))[:24]
    return ""


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
