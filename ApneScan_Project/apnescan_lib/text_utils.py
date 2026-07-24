"""Text / filename helpers — pure string transforms, no state, no PyQt.

These build safe file and folder names from user- or OCR-supplied text. They are
deliberately dependency-light (only :mod:`re`) so they can be unit-tested in
isolation. Behaviour is byte-for-byte identical to the original definitions that
lived in ``apnescan.py``.
"""

import re

__all__ = ["sanitize", "underscore_name", "name_key", "folder_safe_name"]


def sanitize(text, fallback="scan"):
    """Return *text* reduced to a filesystem-safe token.

    Keeps only alphanumerics, ``-`` and ``_``; falls back to *fallback* when
    nothing usable remains.
    """
    safe = "".join(c for c in (text or "") if c.isalnum() or c in "-_")
    return safe or fallback


def underscore_name(s):
    """Filename-friendly: spaces -> underscore, drop odd chars, collapse repeats.

    Used for PDF / file names (which keep the underscore convention).
    """
    s = re.sub(r"[^\w\s.\-]", "", s or "")
    s = re.sub(r"\s+", "_", s.strip())
    s = re.sub(r"_+", "_", s)
    return s.strip("_.")


def name_key(s):
    """Normalized key to match the 'same' document title across scans."""
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def folder_safe_name(s):
    """Human-readable folder name for NEW patient folders.

    Keeps SPACES (no underscores), drops only the characters Windows forbids and
    collapses repeats: ``'Rajendra_Kumar'`` / ``'Rajendra   Kumar'`` ->
    ``'Rajendra Kumar'``. (File names still use :func:`underscore_name`.)
    """
    s = (s or "").strip()
    s = re.sub(r"[_]+", " ", s)                      # underscores -> readable spaces
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", s)      # Windows-illegal chars
    s = re.sub(r"\s+", " ", s).strip()
    s = s.rstrip(". ")                               # Windows: no trailing dot/space
    return s or "New Folder"
