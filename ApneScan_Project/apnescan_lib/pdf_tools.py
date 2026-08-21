"""PDF tools engine — pure file→file operations for ApneScan's My Files panel.

Every function takes a source PDF path (and params) and writes a new file, never
touching the original. Uses pypdf for structure ops (extract/delete/rotate/
protect) and PyMuPDF (fitz) for render ops (grayscale/crop/repair/flatten/
png/extract-images/n-up). No PyQt, no app state — easy to test in isolation.

Page numbers in the public API are 1-based (as the user sees them); parse_ranges
turns "3-5, 8, 11-" into 0-based indices.
"""

import os
import re

try:
    from pypdf import PdfReader, PdfWriter
    HAS_PYPDF = True
except Exception:  # pragma: no cover
    HAS_PYPDF = False

try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except Exception:  # pragma: no cover
    HAS_FITZ = False

__all__ = [
    "page_count", "parse_ranges", "extract_pages", "delete_pages", "rotate_pdf",
    "split_each_page", "split_every_n", "grayscale_pdf", "crop_margins",
    "repair_pdf", "flatten_pdf", "protect_pdf", "remove_password", "pdf_to_png",
    "extract_images", "n_up", "page_numbers_file", "watermark_file",
    "HAS_PYPDF", "HAS_FITZ",
]


def page_count(src):
    if HAS_FITZ:
        d = fitz.open(src)
        try:
            return d.page_count
        finally:
            d.close()
    r = PdfReader(src)
    return len(r.pages)


def parse_ranges(spec, total):
    """'3-5, 8, 11-' + total -> sorted unique 0-based indices. Empty/invalid ->
    []. Open ranges ('11-') go to the last page; '-3' means 1..3."""
    out = set()
    for part in re.split(r"[,\s]+", (spec or "").strip()):
        if not part:
            continue
        m = re.match(r"^(\d*)-(\d*)$", part)
        if m:
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(2)) if m.group(2) else total
        elif part.isdigit():
            a = b = int(part)
        else:
            continue
        for p in range(min(a, b), max(a, b) + 1):
            if 1 <= p <= total:
                out.add(p - 1)
    return sorted(out)


def _writer_from(reader, indices):
    w = PdfWriter()
    for i in indices:
        w.add_page(reader.pages[i])
    return w


def extract_pages(src, out, indices):
    """indices = 0-based list of pages to KEEP (in order)."""
    r = PdfReader(src)
    if not indices:
        raise ValueError("no pages selected")
    w = _writer_from(r, [i for i in indices if 0 <= i < len(r.pages)])
    with open(out, "wb") as f:
        w.write(f)
    return out


def delete_pages(src, out, del_indices):
    """Keep everything EXCEPT del_indices (0-based)."""
    r = PdfReader(src)
    keep = [i for i in range(len(r.pages)) if i not in set(del_indices)]
    if not keep:
        raise ValueError("that would delete every page")
    w = _writer_from(r, keep)
    with open(out, "wb") as f:
        w.write(f)
    return out


def rotate_pdf(src, out, degrees, indices=None):
    """Rotate given pages (0-based; None = all) by degrees (90/180/270)."""
    r = PdfReader(src)
    w = PdfWriter()
    sel = set(indices) if indices is not None else None
    for i, pg in enumerate(r.pages):
        if sel is None or i in sel:
            try:
                pg.rotate(int(degrees) % 360)
            except Exception:
                pass
        w.add_page(pg)
    with open(out, "wb") as f:
        w.write(f)
    return out


def split_each_page(src, outdir, base):
    """Har page ki alag PDF: base_01.pdf, base_02.pdf … Returns list of paths."""
    r = PdfReader(src)
    n = len(r.pages)
    outs = []
    for i in range(n):
        w = PdfWriter(); w.add_page(r.pages[i])
        p = os.path.join(outdir, "%s_%02d.pdf" % (base, i + 1))
        with open(p, "wb") as f:
            w.write(f)
        outs.append(p)
    return outs


def split_every_n(src, outdir, base, n):
    """Har n page ka ek chunk. Returns list of paths."""
    n = max(1, int(n))
    r = PdfReader(src)
    total = len(r.pages)
    outs = []
    part = 0
    for start in range(0, total, n):
        part += 1
        w = PdfWriter()
        for i in range(start, min(start + n, total)):
            w.add_page(r.pages[i])
        p = os.path.join(outdir, "%s_part%02d.pdf" % (base, part))
        with open(p, "wb") as f:
            w.write(f)
        outs.append(p)
    return outs


def grayscale_pdf(src, out, dpi=150):
    """Har page ko grayscale render karke nayi PDF — rangeen PDF chhoti + B&W
    print-ready. (fitz zaroori.)"""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) not available")
    doc = fitz.open(src)
    newd = fitz.open()
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    try:
        for pg in doc:
            pix = pg.get_pixmap(matrix=mat, colorspace=fitz.csGRAY, alpha=False)
            rect = pg.rect
            np = newd.new_page(width=rect.width, height=rect.height)
            np.insert_image(np.rect, pixmap=pix)
        newd.save(out, deflate=True, garbage=4)
    finally:
        doc.close(); newd.close()
    return out


def _content_bbox(pix, thresh=245, pad=4):
    """Render pixmap se non-white content ka bounding box (pixels)."""
    import struct
    w, h, n = pix.width, pix.height, pix.n
    samples = pix.samples
    minx, miny, maxx, maxy = w, h, 0, 0
    found = False
    step = max(1, (w * h) // 400000)   # bade page par thoda sub-sample (tez)
    for y in range(0, h, 1):
        row = y * pix.stride
        for x in range(0, w, step):
            off = row + x * n
            # gray/rgb: koi bhi channel dark -> content
            dark = False
            for c in range(min(3, n)):
                if samples[off + c] < thresh:
                    dark = True; break
            if dark:
                found = True
                if x < minx: minx = x
                if x > maxx: maxx = x
                if y < miny: miny = y
                if y > maxy: maxy = y
    if not found:
        return None
    minx = max(0, minx - pad); miny = max(0, miny - pad)
    maxx = min(w, maxx + pad); maxy = min(h, maxy + pad)
    return (minx, miny, maxx, maxy)


def crop_margins(src, out, dpi=100):
    """Har page ke content ke charo taraf ki khaali (white) jagah kaat do."""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) not available")
    doc = fitz.open(src)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    try:
        for pg in doc:
            pix = pg.get_pixmap(matrix=mat, colorspace=fitz.csGRAY, alpha=False)
            bb = _content_bbox(pix)
            if bb:
                minx, miny, maxx, maxy = bb
                r = pg.rect
                cb = fitz.Rect(r.x0 + minx / zoom, r.y0 + miny / zoom,
                               r.x0 + maxx / zoom, r.y0 + maxy / zoom)
                cb = cb & r
                if cb.width > 10 and cb.height > 10:
                    pg.set_cropbox(cb)
        doc.save(out, deflate=True, garbage=4)
    finally:
        doc.close()
    return out


def repair_pdf(src, out):
    """Kharab/na-khulne wali PDF ko dobara likho (clean + rebuild)."""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) not available")
    doc = fitz.open(src)
    try:
        doc.save(out, garbage=4, deflate=True, clean=True)
    finally:
        doc.close()
    return out


def flatten_pdf(src, out, dpi=150):
    """Har page ko image bana kar nayi PDF — form/annotation/mohar 'pakki' (badli
    na ja sake). Colour rehta hai."""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) not available")
    doc = fitz.open(src)
    newd = fitz.open()
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    try:
        for pg in doc:
            pix = pg.get_pixmap(matrix=mat, alpha=False)
            rect = pg.rect
            np = newd.new_page(width=rect.width, height=rect.height)
            np.insert_image(np.rect, pixmap=pix)
        newd.save(out, deflate=True, garbage=4)
    finally:
        doc.close(); newd.close()
    return out


def protect_pdf(src, out, password):
    """PDF par password lagao (kholne ke liye password chahiye)."""
    if not password:
        raise ValueError("empty password")
    r = PdfReader(src)
    w = PdfWriter()
    for pg in r.pages:
        w.add_page(pg)
    w.encrypt(str(password))
    with open(out, "wb") as f:
        w.write(f)
    return out


def pdf_to_png(src, outdir, base, dpi=200):
    """Har page ki hi-res PNG image. Returns list of paths."""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) not available")
    doc = fitz.open(src)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    outs = []
    try:
        for i, pg in enumerate(doc):
            pix = pg.get_pixmap(matrix=mat, alpha=False)
            p = os.path.join(outdir, "%s_%02d.png" % (base, i + 1))
            pix.save(p)
            outs.append(p)
    finally:
        doc.close()
    return outs


def extract_images(src, outdir, base):
    """PDF ke ANDAR ki embedded images nikaalo. Returns list of paths."""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) not available")
    doc = fitz.open(src)
    outs = []
    seen = set()
    try:
        k = 0
        for pno in range(doc.page_count):
            for img in doc.get_page_images(pno):
                xref = img[0]
                if xref in seen:
                    continue
                seen.add(xref)
                try:
                    pix = fitz.Pixmap(doc, xref)
                    if pix.n >= 5:               # CMYK/alpha -> RGB
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    k += 1
                    p = os.path.join(outdir, "%s_img%02d.png" % (base, k))
                    pix.save(p)
                    outs.append(p)
                    pix = None
                except Exception:
                    continue
    finally:
        doc.close()
    return outs


def n_up(src, out, per=2):
    """2-in-1 / 4-in-1: source ke kai page ek A4 sheet par (print-friendly)."""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) not available")
    per = 2 if per not in (2, 4) else per
    doc = fitz.open(src)
    newd = fitz.open()
    A4 = fitz.paper_rect("a4")            # portrait
    try:
        cols, rows = (1, 2) if per == 2 else (2, 2)
        cw, ch = A4.width / cols, A4.height / rows
        slot = 0; page = None
        for pno in range(doc.page_count):
            if slot == 0:
                page = newd.new_page(width=A4.width, height=A4.height)
            r = slot // cols
            c = slot % cols
            cell = fitz.Rect(c * cw, r * ch, (c + 1) * cw, (r + 1) * ch)
            cell = cell + (4, 4, -4, -4)          # thoda gap
            page.show_pdf_page(cell, doc, pno)
            slot = (slot + 1) % per
        newd.save(out, deflate=True, garbage=4)
    finally:
        doc.close(); newd.close()
    return out


def remove_password(src, out, password):
    """Password pata ho to bina-password copy banao."""
    r = PdfReader(src)
    if r.is_encrypted:
        if not r.decrypt(str(password or "")):
            raise ValueError("wrong password")
    w = PdfWriter()
    for pg in r.pages:
        w.add_page(pg)
    with open(out, "wb") as f:
        w.write(f)
    return out


def page_numbers_file(src, out, header=""):
    """Har page ke neeche 'Page i / n' (aur chaaho to upar header text)."""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) not available")
    doc = fitz.open(src)
    try:
        n = doc.page_count
        for i, pg in enumerate(doc):
            r = pg.rect
            foot = "Page %d / %d" % (i + 1, n)
            pg.insert_text((r.width / 2.0 - 28, r.height - 22), foot,
                           fontsize=9, color=(0.25, 0.25, 0.25))
            if header:
                pg.insert_text((36, 28), str(header)[:120],
                               fontsize=10, color=(0.25, 0.25, 0.25))
        doc.save(out, deflate=True, garbage=4)
    finally:
        doc.close()
    return out


def watermark_file(src, out, text, opacity=0.12):
    """Har page par tirchha halka watermark (jaise 'COPY' / clinic naam)."""
    if not HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) not available")
    text = (text or "").strip() or "COPY"
    doc = fitz.open(src)
    try:
        for pg in doc:
            r = pg.rect
            fs = max(24, int(min(r.width, r.height) / max(6, len(text)) * 1.6))
            try:
                pg.insert_textbox(
                    fitz.Rect(0, r.height / 2 - fs, r.width, r.height / 2 + fs),
                    text, fontsize=fs, color=(0.5, 0.5, 0.5),
                    align=1, rotate=45, fill_opacity=opacity)
            except Exception:
                # purane fitz me fill_opacity/rotate na ho to simple text
                pg.insert_text((r.width * 0.25, r.height * 0.5), text,
                               fontsize=fs, color=(0.7, 0.7, 0.7))
        doc.save(out, deflate=True, garbage=4)
    finally:
        doc.close()
    return out
