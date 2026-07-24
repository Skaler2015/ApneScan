"""Smart patient-folder search engine — pure ranking logic, no state, no PyQt.

Ranks patient folders the way premium hospital software does: the right folder
surfaces at the top regardless of spaces vs underscores, case, or typos. Works
identically for OLD folders (``793_Rajendra_Kumar_MR``) and NEW ones
(``793 Rajendra Kumar MR``) because both normalise to the same tokens.

Only depends on :mod:`re` and :mod:`difflib`, so the whole engine is unit-testable
in isolation. Behaviour is identical to the original definitions in ``apnescan.py``.
"""

import re
import difflib

__all__ = [
    "_norm_search", "_vowel_skeleton", "_term_category", "_folder_search_score",
]


def _norm_search(s):
    """Normalise a folder name OR a search query to the SAME internal form:
    lowercase, treat ``_`` and ``-`` as spaces, collapse repeats, trim.

    ``'793_Rajendra_Kumar_MR'`` -> ``'793 rajendra kumar mr'``.
    """
    s = (s or "").lower()
    s = re.sub(r"[_\-]+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _vowel_skeleton(s):
    """Drop vowels so vowel-only spelling differences collapse together
    (``'rajinder'``/``'rajender'``/``'rajendra'`` -> ``'rjndr'``).

    Keeps the first char even if it is a vowel, so ``'ayush'``/``'aiyush'`` stay
    comparable.
    """
    if not s:
        return s
    return s[0] + re.sub(r"[aeiou]", "", s[1:])


def _term_category(t, tokens, num, name_part, name_norm):
    """How strongly one query token matches a folder. Lower = better.

    ``0`` exact folder-number · ``1`` exact word · ``2`` name-starts-with ·
    ``3`` word-starts-with · ``4`` contains · ``5`` fuzzy(typo) · ``None`` = no match.
    """
    if num is not None and t == num:
        return 0
    if t in tokens:
        return 1
    if name_part.startswith(t):
        return 2
    if any(tok.startswith(t) for tok in tokens):
        return 3
    if t in name_norm:
        return 4
    # P6 fuzzy — typo/spelling tolerance. Two signals, both guarded by a shared
    # first letter so it stays safe across 100k folders:
    #   (a) close overall similarity (rajendar/rajendr/rajedra -> rajendra), or
    #   (b) same consonant-skeleton, i.e. only the vowels differ — the common
    #       Indian-name case (rajinder/rajender/rajendra all -> 'rjndr').
    if len(t) >= 4:
        tv = _vowel_skeleton(t)
        for tok in tokens:
            if len(tok) < 4 or tok[0] != t[0]:
                continue
            if difflib.SequenceMatcher(None, t, tok).ratio() >= 0.8:
                return 5
            sk = _vowel_skeleton(tok)
            if tv and sk and difflib.SequenceMatcher(None, tv, sk).ratio() >= 0.85:
                return 5
    return None


def _folder_search_score(terms, name_norm):
    """Rank a folder for a multi-word query.

    Returns ``(matched, sort_key)`` where a SMALLER key ranks higher. A folder
    matches only if EVERY query token finds a home in it. Priority order realised
    by the key: 1 exact folder-number · 2 exact patient-name · 3 name-starts ·
    4 word-starts · 5 contains · 6 fuzzy.
    """
    tokens = name_norm.split()
    if not tokens or not terms:
        return (False, None)
    num = tokens[0] if tokens[0].isdigit() else None
    name_part = " ".join(tokens[1:] if num else tokens)   # patient name (no number)
    qjoin = " ".join(terms)
    exact_name = (qjoin == name_part or qjoin == name_norm)
    exact_number = False
    cats = []
    for t in terms:
        c = _term_category(t, tokens, num, name_part, name_norm)
        if c is None:
            return (False, None)                          # a token matched nothing
        if c == 0:
            exact_number = True
        cats.append(c)
    key = (0 if exact_number else 1,      # P1: exact folder number to the very top
           0 if exact_name else 1,        # P2: exact full patient name next
           max(cats),                     # weakest token's strength (P3/P4/P5/P6)
           sum(cats),                     # overall tightness of the match
           len(name_norm),                # shorter (more specific) name first
           name_norm)                     # stable alphabetical tie-break
    return (True, key)
