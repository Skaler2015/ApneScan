"""Intelligent Naming Engine — build clean {Name}-{Type} filenames from OCR text.

Extracted verbatim from apnescan.py. Pure text processing (only :mod:`re`): it
detects the document type, extracts a person/business name and any document
number, learns the user's filename style, and applies OCR corrections — all
offline and stateless. Behaviour is byte-for-byte identical to the original.
"""

import re


# ---- INTELLIGENT NAMING ENGINE (offline) ----
_NI_TYPES = [
    ("Aadhaar",           ["aadhaar", "aadhar", "uidai", "आधार", "unique identification"], 9),
    ("PAN",               ["permanent account number", "income tax department", "pan card"], 9),
    ("Passport",          ["passport", "republic of india\npassport"], 9),
    ("Driving Licence",   ["driving licence", "driving license", "transport department", " dl no", "form 7"], 9),
    ("Voter ID",          ["election commission", "elector", "voter", "epic no"], 9),
    ("Cheque",            ["account payee", "pay to the", "cheque", " ifsc", "bearer"], 8),
    ("Invoice",           ["tax invoice", "gstin", "invoice"], 6),
    ("Salary Slip",       ["salary slip", "pay slip", "payslip", "net pay", "earnings\ndeductions"], 8),
    ("Bank Statement",    ["statement of account", "bank statement", "opening balance", "closing balance"], 8),
    ("Utility Bill",      ["electricity bill", "water bill", "gas bill", "units consumed", "meter reading"], 8),
    ("Receipt",           ["receipt", "received with thanks", "received from"], 5),
    ("Bill",              ["bill of supply", " bill "], 4),
    ("Prescription",      ["prescription", " rx ", "advice", "tab.", "cap.", "1-0-1"], 6),
    ("Discharge Summary", ["discharge summary", "date of discharge", "discharge"], 8),
    ("Consent Form",      ["consent form", "i hereby consent", "informed consent"], 8),
    ("Medical Certificate", ["medical certificate", "fitness certificate", "certified that", "certify that"], 7),
    ("Insurance Claim",   ["claim form", "claim number", "claim id"], 8),
    ("Insurance",         ["policy no", "sum insured", "insurance", "policy holder"], 6),
    ("RGHS",              ["rghs"], 9),
    ("ESIC",              ["esic", "employees state insurance"], 9),
    ("Salary Slip",       ["salary", "basic pay"], 3),
    ("Agreement",         ["agreement", "hereby agree", "terms and conditions", "party of the first"], 7),
    ("Quotation",         ["quotation", "quote no", "we are pleased to quote"], 8),
    ("Purchase Order",    ["purchase order", "po number", "p.o. no"], 8),
    ("Delivery Challan",  ["delivery challan", "challan"], 8),
    ("Student ID",        ["student id", "roll no", "enrollment no", "student identity"], 8),
    ("Employee ID",       ["employee id", "emp id", "employee code", "employee identity"], 8),
    ("ID Card",           ["identity card", "id card"], 6),
    ("Tax Document",      ["form 16", "tds certificate", "income tax return", "itr"], 8),
    ("Application",       ["application for", "applicant", "i am applying"], 6),
    ("Letter",            ["dear sir", "subject:", "yours faithfully", "yours sincerely"], 5),
]
# words that look like a person's title -> strip
_NI_TITLES = {"mr", "mrs", "ms", "dr", "shri", "sri", "smt", "km", "kumari", "mstr", "master", "m/s", "prof"}
# tokens that must never appear inside an extracted "name"
_NI_STOP = set((
    "government india republic name naam patient holder card number no date dob "
    "aadhaar aadhar pan invoice bill receipt hospital ward summary discharge tax "
    "prescription report lab claim rghs esic echs the of male female address from "
    "amount total gst gstin cheque bank policy insurance department form son wife "
    "daughter father mother address city state pin phone mobile email dated bill "
    "quotation challan salary slip statement licence license passport voter "
    "permanent account income order estimate details enrollment enrolment "
    "certificate agreement application letter document medical general "
    "rx tab cap mg ml dose advice diagnosis opd ipd reg regn"
).split())
_NI_MEDICAL = ("patient", "hospital", "doctor", "clinic", "diagnosis", "prescription",
               "discharge", "blood", "report", "test", "mg", "tablet", "opd")


def _ni_detect_type(text):
    """Return (filename_label, confidence 0..100). Only categories from the
    allowed list; unknown -> 'Document' (or 'Medical Document' if clearly medical).
    Never returns a specific report name (CBC/LFT/MRI...)."""
    if not text:
        return ("Document", 0)
    low = "\n" + text.lower() + "\n"
    best, best_pri, hits = None, -1, 0
    for label, keys, pri in _NI_TYPES:
        for k in keys:
            if k in low:
                if pri > best_pri:
                    best, best_pri = label, pri
                hits += 1
                break
    if best:
        conf = min(98, 70 + best_pri * 3)
        return (best, conf)
    # medical-looking but not a listed category -> generic medical document
    if sum(1 for w in _NI_MEDICAL if w in low) >= 2:
        return ("Medical Document", 55)
    return ("Document", 30)


def _ni_clean_ocr(s, corrections=None):
    """Apply the user's LEARNED OCR corrections (and a few universal ones)."""
    if not s:
        return s
    out = s
    if corrections:
        for wrong, right in corrections.items():
            if wrong and right:
                out = re.sub(r'\b' + re.escape(wrong) + r'\b', right, out, flags=re.I)
    return out


def _ni_titlecase(w):
    return w[:1].upper() + w[1:].lower() if w else w


def _ni_clean_name(raw, corrections=None, max_words=3):
    """Turn a raw name fragment into a clean Title-Case name, dropping titles /
    stop-words / OCR garbage. Only keeps capitalized (name-like) words. '' if
    not a plausible name."""
    if not raw:
        return ""
    raw = _ni_clean_ocr(raw, corrections)
    words = re.findall(r"[A-Za-z]+", raw)
    out = []
    for w in words:
        lw = w.lower()
        if lw in _NI_TITLES or lw in _NI_STOP or len(w) < 2:
            continue
        if not (w[0].isupper() or w.isupper()):   # real names are capitalized
            continue
        out.append(_ni_titlecase(w))
        if len(out) >= max_words:
            break
    return " ".join(out)


def _ni_extract_person(text, corrections=None, allow_fallback=True):
    """Best person name from OCR (label-guided first, then, if allowed, the first
    clean Title-Case run). allow_fallback=False for invoices/bills so a business
    header isn't mistaken for a person."""
    if not text:
        return ""
    # label-guided: Name / Patient Name / S/O etc.
    m = re.search(r'(?:patient\s*name|holder\s*name|\bname\b|\bnaam\b|\bnaame\b)\s*[:\-]?\s*'
                  r'((?:(?:mr|mrs|ms|dr|shri|sri|smt|kumari)\.?\s+)?[A-Za-z][A-Za-z]+(?:\s+[A-Za-z]+){0,2})',
                  text, re.I)
    if m:
        c = _ni_clean_name(m.group(1), corrections, 3)
        if len(c.split()) >= 2:
            return c
    # title-prefixed name anywhere: "Mr Ram Kumar", "Dr. Suresh Meena"
    m = re.search(r'\b(?:Mr|Mrs|Ms|Dr|Shri|Sri|Smt|Kumari)\.?\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,2})', text)
    if m:
        c = _ni_clean_name(m.group(1), corrections, 3)
        if len(c.split()) >= 2:
            return c
    if not allow_fallback:
        return ""
    # first clean Title-Case run of 2-3 words
    for m in re.finditer(r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})\b', text):
        c = _ni_clean_name(m.group(1), corrections, 3)
        if len(c.split()) >= 2:
            return c
    return ""


_NI_BIZ_SUFFIX = ("store", "stores", "traders", "trader", "medical", "medicals", "pharma",
                  "pharmacy", "enterprises", "enterprise", "agencies", "agency", "ltd",
                  "pvt", "limited", "company", "co", "industries", "hospital", "clinic",
                  "labs", "laboratory", "diagnostics", "motors", "electronics", "hardware")


def _ni_extract_business(text, corrections=None):
    """Business/seller name for invoices/bills — usually a prominent top line
    ending in a business word; never 'Tax Invoice'/'GSTIN'."""
    if not text:
        return ""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for ln in lines[:8]:
        low = ln.lower()
        if any(bad in low for bad in ("tax invoice", "invoice", "gstin", "bill", "receipt", "estimate")):
            continue
        if any(sfx in low.split() or low.endswith(sfx) for sfx in _NI_BIZ_SUFFIX):
            words = re.findall(r"[A-Za-z&][A-Za-z&.]*", ln)
            words = [w for w in words if w.lower() not in ("the", "of")][:5]
            if 1 <= len(words) <= 5:
                # keep short ALL-CAPS acronyms (ABC), Title-Case the rest
                return " ".join(w if (w.isupper() and len(w) <= 4) else _ni_titlecase(w)
                                for w in words)
    return ""


def _ni_extract_number(text, label):
    """Extract the document number relevant to the type. Clean, short."""
    if not text:
        return ""
    t = text
    def g(pat):
        m = re.search(pat, t, re.I)
        return (m.group(1).strip(" .:-") if m else "")
    if label in ("Invoice", "GST Invoice"):
        # require the 'no/number/#' label AND at least one digit -> never grabs GSTIN
        return g(r'invoice\s*(?:no\.?|number|#)\s*[:#\-]?\s*([A-Za-z0-9/\-]*\d[A-Za-z0-9/\-]*)')
    if label == "Cheque":
        return g(r'cheque\s*(?:no|number|#)?\s*[:#\-]?\s*(\d{5,8})') or g(r'\b(\d{6})\b')
    if label in ("Insurance Claim",):
        return g(r'claim\s*(?:no|number|id|#)?\s*[:#\-]?\s*([A-Za-z0-9/\-]{3,20})')
    if label == "Insurance":
        return g(r'policy\s*(?:no|number|#)?\s*[:#\-]?\s*([A-Za-z0-9/\-]{4,22})')
    if label == "Passport":
        return g(r'\b([A-PR-WYa-pr-wy][0-9]{7})\b')
    if label == "PAN":
        return g(r'\b([A-Z]{5}[0-9]{4}[A-Z])\b')
    if label == "Application":
        return g(r'application\s*(?:no|number|id|#)?\s*[:#\-]?\s*([A-Za-z0-9/\-]{3,20})')
    if label == "Purchase Order":
        return g(r'(?:p\.?o\.?|purchase order)\s*(?:no|number|#)?\s*[:#\-]?\s*([A-Za-z0-9/\-]{2,18})')
    return ""


# ---- naming STYLE (separator / order / caps) learned from user's filenames ----
def _ni_default_style():
    return {"sep": " - ", "order": "name_first", "caps": "title"}


def _ni_learn_style(filenames):
    """Infer the user's separator + word order from their past filenames.
    (v224) UNDERSCORE ab kabhi separator nahi banta — user chahta hai naam me
    SPACE ho, isliye purani underscore-wali files se '_' seekhna band. '_'
    wali files ko space-jaisa gina jaata hai."""
    st = _ni_default_style()
    seps = {" - ": 0, "-": 0, " ": 0}
    for fn in filenames:
        if not fn:
            continue
        if " - " in fn:
            seps[" - "] += 1
        elif "-" in fn:
            seps["-"] += 1
        elif "_" in fn or " " in fn.strip():   # '_' ab space jaisa maana jaata
            seps[" "] += 1
    if any(seps.values()):
        st["sep"] = max(seps, key=seps.get)
    return st


def _ni_join(parts, style):
    parts = [p for p in parts if p]
    if not parts:
        return ""
    sep = style.get("sep", " - ") if style else " - "
    if "_" in sep:              # (v224) purana stored '_' style bhi -> space
        sep = " "
    return sep.join(parts)


def _ni_generate(text, style=None, corrections=None):
    """Generate a clean professional filename from OCR text alone.
    Returns (filename, type_label, type_confidence)."""
    style = style or _ni_default_style()
    label, tconf = _ni_detect_type(text)
    docnum = _ni_extract_number(text, label)
    is_biz = label in ("Invoice", "GST Invoice", "Bill", "Receipt", "Quotation",
                       "Purchase Order", "Delivery Challan")
    # for invoices/bills only a LABELLED person counts (else a business header
    # would be mistaken for a person); the business name is the usual subject.
    person = _ni_extract_person(text, corrections, allow_fallback=not is_biz)
    business = _ni_extract_business(text, corrections) if is_biz else ""

    subject = person or business
    # special professional layouts
    if label == "Cheque" and person:
        name = _ni_join(["Cheque", person, docnum], style)
    elif label in ("Invoice", "GST Invoice") and business and not person:
        inv = ("Invoice " + docnum).strip()
        name = _ni_join([business, inv], style)
    elif subject:
        typ = label
        if docnum and label in ("Invoice", "GST Invoice", "Insurance Claim"):
            typ = (label + " " + docnum).strip()
        name = _ni_join([subject, typ], style)
    else:
        # no person/business -> use the type (+ number) alone
        name = _ni_join([label, docnum], style) if docnum else label
    name = re.sub(r'\s+', ' ', name).strip(" -_")
    # (v224) underscore ko kabhi space se replace nahi karte — naam me space rehta
    return (name, label, tconf)


def _ni_lev(a, b):
    """Small Levenshtein distance (for spotting OCR typos in a name token)."""
    if a == b:
        return 0
    if not a or not b:
        return max(len(a), len(b))
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _ni_learn_corrections(extracted, final):
    """Compare an OCR-extracted name with the user's final name and return
    [(wrong, correct)] token pairs that look like OCR typos (small edit distance).
    Only same-position tokens; conservative so real edits aren't 'corrected'."""
    ew = re.findall(r"[A-Za-z]+", extracted or "")
    fw = re.findall(r"[A-Za-z]+", final or "")
    out = []
    if not ew or not fw or abs(len(ew) - len(fw)) > 1:
        return out
    for a, b in zip(ew, fw):
        if a.lower() == b.lower() or len(a) < 4 or len(b) < 4:
            continue
        d = _ni_lev(a.lower(), b.lower())
        if 1 <= d <= 2 and abs(len(a) - len(b)) <= 2:
            out.append((a, b[:1].upper() + b[1:].lower()))
    return out
