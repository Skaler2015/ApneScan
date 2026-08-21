"""
pdf_text_edit.py — Digital/native PDF ke ASLI text ko badalne ke liye
pure helper functions (PyMuPDF/fitz par).

ApneScan zyadatar SCANNED PDF (photo-jaisi) banata hai — un me text asli
"akshar" nahi, tasveer ke pixels hote hain, isliye unka text seedha nahi
badalta. Ye module SIRF un PDF par kaam karta hai jin me asli text-layer ho
(Word/print-to-PDF se bani). Har edit = purane word par redaction (safed) +
usi jagah naya text daalna (font size/rang milaakar).

UI (dialog) apnescan.py me hai; ye file sirf file<->file / doc-level kaam
karti hai taaki test karna aasaan rahe.
"""

import fitz  # PyMuPDF


# ---- base-14 font chunna (span ke flags se bold/italic/serif) ---------------
# fitz flags bits: 1=superscript, 2=italic, 4=serifed, 8=monospaced, 16=bold
def _base14_for(flags):
    bold = bool(flags & 16)
    italic = bool(flags & 2)
    mono = bool(flags & 8)
    serif = bool(flags & 4)
    if mono:
        return {(0, 0): "cour", (1, 0): "cobo",
                (0, 1): "coit", (1, 1): "cobi"}[(int(bold), int(italic))]
    if serif:
        return {(0, 0): "tiro", (1, 0): "tibo",
                (0, 1): "tiit", (1, 1): "tibi"}[(int(bold), int(italic))]
    return {(0, 0): "helv", (1, 0): "hebo",
            (0, 1): "heit", (1, 1): "hebi"}[(int(bold), int(italic))]


def _color_rgb(c):
    """fitz span color (int sRGB) -> (r,g,b) 0..1 tuple."""
    try:
        c = int(c)
    except Exception:
        return (0, 0, 0)
    r = ((c >> 16) & 255) / 255.0
    g = ((c >> 8) & 255) / 255.0
    b = (c & 255) / 255.0
    return (r, g, b)


def has_text_layer(src, min_chars=6):
    """True agar PDF me asli (selectable) text hai — yaani digital PDF.
    Scanned PDF (sirf image) par False."""
    doc = fitz.open(src)
    try:
        total = 0
        for page in doc:
            total += len((page.get_text("text") or "").strip())
            if total >= min_chars:
                return True
        return total >= min_chars
    finally:
        doc.close()


def page_spans(doc, page_no):
    """Ek page ke saare text 'spans' — har ek editable tukda.
    Lauta: list of dict {id, text, bbox(x0,y0,x1,y1), size, flags,
    color(int), font, origin(x,y)}.  'id' = (block, line, span) index.
    """
    page = doc[page_no]
    out = []
    d = page.get_text("dict")
    for bi, block in enumerate(d.get("blocks", [])):
        if block.get("type", 0) != 0:      # 0 = text block
            continue
        for li, line in enumerate(block.get("lines", [])):
            for si, span in enumerate(line.get("spans", [])):
                txt = span.get("text", "")
                if not txt.strip():
                    continue
                bbox = tuple(span.get("bbox", (0, 0, 0, 0)))
                origin = span.get("origin")
                if not origin:
                    origin = (bbox[0], bbox[3])
                out.append({
                    "id": (bi, li, si),
                    "text": txt,
                    "bbox": bbox,
                    "size": float(span.get("size", 11.0)),
                    "flags": int(span.get("flags", 0)),
                    "color": span.get("color", 0),
                    "font": span.get("font", ""),
                    "origin": (float(origin[0]), float(origin[1])),
                })
    return out


def _find_span(doc, page_no, span_id):
    for sp in page_spans(doc, page_no):
        if sp["id"] == span_id:
            return sp
    return None


def replace_span(doc, page_no, span, new_text, fit=True,
                 scanned=False, fill=(1, 1, 1)):
    """Ek span (purana text) ko new_text se badlo — usi jagah, usi size/rang.
    span = page_spans() (digital) ya OCR-span (scanned) dict.

    digital PDF (scanned=False): purane text ko redaction se hatate hain.
    scanned PDF (scanned=True) : page ek tasveer hai — purane word ko 'fill'
        rang (kagaz ka background) ke bhare box se dhak dete hain, phir uske
        upar naya text likhte hain (bilkul Adobe/WPS ki tarah). NOTE: ye
        ek SAAF sudhaar hai, koi chhupa hua/undetectable badlaav nahi —
        gaur se dekhne par ya forensic jaanch me edit pakda ja sakta hai.
    fit=True: naya text chaura ho to size thoda ghata kar bbox me fit.
    """
    page = doc[page_no]
    bbox = fitz.Rect(span["bbox"])

    if scanned:
        # tasveer wale word ko background-rang ke box se dhako (whiteout)
        cover = bbox + (-1.5, -1.5, 1.5, 1.5)
        try:
            page.draw_rect(cover, color=fill, fill=fill, width=0)
        except Exception:
            page.draw_rect(cover, color=(1, 1, 1), fill=(1, 1, 1), width=0)
    else:
        # digital text ko redaction se hatao (sirf isi span ke bbox par)
        page.add_redact_annot(bbox, fill=(1, 1, 1))
        try:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        except Exception:
            page.apply_redactions()

    if not new_text:
        return                              # sirf mitana tha

    size = span["size"]
    fontname = _base14_for(span["flags"])
    color = _color_rgb(span["color"])
    origin = fitz.Point(span["origin"])

    # naya text bbox me fit karne ke liye size adjust (agar zaroori)
    if fit:
        avail = bbox.width
        if avail > 2:
            tl = fitz.get_text_length(new_text, fontname=fontname, fontsize=size)
            if tl > avail:
                size = max(4.0, size * (avail / tl) * 0.98)

    # naya text usi baseline par daalo
    page.insert_text(origin, new_text, fontname=fontname,
                     fontsize=size, color=color)


def apply_edits(src, out, edits):
    """File->file: src kholo, edits lagao, out me save.
    edits = list of {page, span, new_text}.  span = page_spans() dict.
    Ek page par kai edits ho to unhe niche-se-upar (ya kisi bhi kram me)
    apply karna theek hai kyunki har span ka apna bbox hai; par surakshit
    rahne ke liye hum har edit se pehle span ko dobara nahi dhoondhte —
    caller current bbox de.
    """
    doc = fitz.open(src)
    try:
        for e in edits:
            replace_span(doc, int(e["page"]), e["span"], e.get("new_text", ""),
                         scanned=e.get("scanned", False),
                         fill=e.get("fill", (1, 1, 1)))
        doc.save(out, garbage=4, deflate=True)
    finally:
        doc.close()
    return out
