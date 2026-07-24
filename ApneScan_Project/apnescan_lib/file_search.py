"""File Search Engine — Everything-style instant search for the My Files panel.

Reliability-first design: a file that exists must always be findable — by full
name, any partial (prefix / middle / suffix, including partial NUMBERS like
'3166' or '6901' inside '3166901'), across separators, case-insensitively, in
Hindi/English/mixed Unicode, with typo tolerance and multi-keyword queries.

The class separates the classic stages — normalisation, tokenisation, query
parsing (extension / size / date filters), matching, ranking, caching and
logging — into small reusable methods. Everything is pure Python + re/difflib,
so the whole engine is unit-testable without the app.

Ranking (best first): exact filename → filename starts with query → partial
filename → folder-path match → OCR/content match. Extension terms ('pdf',
'jpg') act as filters; 'today'/'yesterday'/'last week'/'last month' filter by
modified date; '>1mb'/'<500kb' filter by size.
"""

import os
import re
import time
import difflib
import datetime

# Every separator / decoration character users put in filenames. All of these
# are treated as spaces so they can never hide a match.
_SEP_RX = re.compile(r"""[_\-.,()\[\]{}#@!%&+='"/\\|~^;:?*<>$\s]+""")
_WS_RX = re.compile(r"\s+")

_KNOWN_EXTS = ("pdf", "jpg", "jpeg", "png", "tif", "tiff", "docx", "xlsx")
# size filter is parsed on the RAW query (norm() would strip the < / > sign)
_SIZE_RX = re.compile(r"([<>])\s*(\d+(?:\.\d+)?)\s*(kb|mb|gb)\b")
_DATE_WORDS = ("today", "yesterday")
_DATE_PHRASE_RX = re.compile(r"\blast\s+(week|month)\b")

SEARCH_LOG = os.path.join(os.path.expanduser("~"), ".apnescan_search.log")


class SearchEngine(object):
    """Reusable, stateless-per-query search engine with a small result cache."""

    def __init__(self):
        self._cache = {}          # query-key -> results (cleared on index change)

    # ------------------------------------------------------------------
    # 1. NORMALISATION
    # ------------------------------------------------------------------
    @staticmethod
    def norm(s):
        """'3166901_Garima-Shek (2).jpg' -> '3166901 garima shek 2 jpg'.

        Unicode-safe (casefold), every separator/special char becomes a space,
        runs collapse. Never raises on odd input.
        """
        try:
            s = (s or "")
            s = _SEP_RX.sub(" ", s)
            s = _WS_RX.sub(" ", s).strip()
            return s.casefold()
        except Exception:
            return ""

    @staticmethod
    def compact(norm_s):
        """Separator-less form: '3166901 garima shek 2 jpg' -> '3166901garimashek2jpg'.

        Lets a query match ACROSS separators ('garimashek', or digits that span
        a boundary), so no partial string can ever be hidden by a '_' or '-'.
        """
        return (norm_s or "").replace(" ", "")

    # ------------------------------------------------------------------
    # 2. TOKENISATION
    # ------------------------------------------------------------------
    @classmethod
    def tokenize(cls, s):
        """Searchable tokens of a name: '3166901_Garima_Shek.jpg' ->
        ['3166901', 'garima', 'shek', 'jpg']."""
        return cls.norm(s).split()

    # ------------------------------------------------------------------
    # 3. QUERY PARSING (text terms + extension / size / date filters)
    # ------------------------------------------------------------------
    @classmethod
    def parse_query(cls, q):
        """Split a raw query into (terms, filters).

        filters = {'ext': 'pdf'|None, 'size': ('>'|'<', bytes)|None,
                   'days': (min_age_days, max_age_days)|None}
        'pdf' alone shows only PDFs; 'today'/'yesterday'/'last week'/'last
        month' filter by modified time; '>1mb'/'<500kb' by size.
        """
        filters = {"ext": None, "size": None, "days": None}
        q = (q or "").casefold()
        m = _DATE_PHRASE_RX.search(q)
        if m:
            filters["days"] = (0, 7) if m.group(1) == "week" else (0, 31)
            q = _DATE_PHRASE_RX.sub(" ", q)
        sm = _SIZE_RX.search(q)          # BEFORE norm(): norm strips the < / > sign
        if sm:
            mul = {"kb": 1024, "mb": 1048576, "gb": 1073741824}[sm.group(3)]
            filters["size"] = (sm.group(1), float(sm.group(2)) * mul)
            q = _SIZE_RX.sub(" ", q)
        terms = []
        for t in cls.norm(q).split():
            if t in _DATE_WORDS:
                filters["days"] = (0, 1) if t == "today" else (1, 2)
                continue
            if t in _KNOWN_EXTS and filters["ext"] is None:
                filters["ext"] = "jpeg" if t == "jpg" else t
                # an extension word doubles as a filter; it is NOT also required
                # to appear in the filename
                continue
            terms.append(t)
        return terms, filters

    # ------------------------------------------------------------------
    # 4. MATCHING (never == / startswith-only: full intelligent substring)
    # ------------------------------------------------------------------
    @staticmethod
    def _fuzzy_tok(t, tok):
        """Typo tolerance: 'garma'~'garima', 'rajedra'~'rajendra'."""
        if len(t) < 4 or len(tok) < 4 or t[0] != tok[0]:
            return False
        if difflib.SequenceMatcher(None, t, tok).ratio() >= 0.78:
            return True
        tv = t[0] + re.sub(r"[aeiou]", "", t[1:])
        kv = tok[0] + re.sub(r"[aeiou]", "", tok[1:])
        return bool(tv and kv and difflib.SequenceMatcher(None, tv, kv).ratio() >= 0.85)

    @staticmethod
    def _digit_subseq(t, tok):
        """Weakest number match: '169' hits '3166901' because 1,6,9 appear in
        order inside the number (subsequence). Digits-only, term >= 3 chars, so
        text search never gets noisy — and it ranks last anyway."""
        if len(t) < 3 or not t.isdigit() or not tok.isdigit() or len(tok) <= len(t):
            return False
        it = iter(tok)
        return all(ch in it for ch in t)

    @classmethod
    def term_strength(cls, t, norm_name, compact_name, tokens, fuzzy=True):
        """How strongly one term matches a name. Lower = better.
        0 exact-name · 1 exact-word · 2 name-starts · 3 word-starts ·
        4 substring anywhere (incl. across separators / inside numbers) ·
        5 fuzzy/number-subsequence · None = no match.

        fuzzy=False skips the expensive typo pass (used for the instant first
        pass over huge indexes; a fuzzy second pass runs only when needed).
        tokens may be None — they are only computed when actually needed, which
        keeps the hot miss-path to a single substring check."""
        # cheap gate: ANY match of categories 0-4 implies t is a substring of
        # the compact (separator-less) form — one 'in' decides the common case.
        if t in compact_name:
            if t == norm_name or t == compact_name:
                return 0
            if tokens is None:
                tokens = norm_name.split()
            if t in tokens:
                return 1
            if norm_name.startswith(t) or compact_name.startswith(t):
                return 2
            if any(tok.startswith(t) for tok in tokens):
                return 3
            return 4
        if fuzzy:
            if tokens is None:
                tokens = norm_name.split()
            if t.isdigit():                 # numbers: cheap subsequence only —
                for tok in tokens:          # SequenceMatcher on digits is noise
                    if cls._digit_subseq(t, tok):
                        return 5
            else:
                for tok in tokens:
                    if cls._fuzzy_tok(t, tok):
                        return 5
        return None

    @classmethod
    def match_file(cls, terms, name_norm, name_compact, name_tokens,
                   folder_norm="", folder_compact="", ocr_text="", fuzzy=True):
        """Match every term against filename first, then folder, then OCR.

        Returns (tier, strengths) or None. tier: 0 = all terms in the NAME,
        1 = needed the folder path, 2 = needed OCR text. strengths drive the
        fine ranking inside a tier.
        """
        tier = 0
        strengths = []
        for t in terms:
            c = cls.term_strength(t, name_norm, name_compact, name_tokens, fuzzy)
            if c is None and folder_norm:
                if t in folder_norm or t in folder_compact:
                    c = 4
                    tier = max(tier, 1)
            if c is None and ocr_text:
                if t in ocr_text:
                    c = 4
                    tier = max(tier, 2)
            if c is None:
                return None
            strengths.append(c)
        return tier, strengths

    # ------------------------------------------------------------------
    # 5. RANKING
    # ------------------------------------------------------------------
    @staticmethod
    def rank_key(tier, strengths, name_norm):
        """Sort key (smaller = higher). Exact filename beats starts-with beats
        partial beats folder beats OCR; shorter names win ties."""
        return (tier, max(strengths) if strengths else 9,
                sum(strengths), len(name_norm), name_norm)

    # ------------------------------------------------------------------
    # 6. FILTERS (extension / size / date)
    # ------------------------------------------------------------------
    @staticmethod
    def passes_filters(filters, ext, size, mtime):
        """Apply parsed ext/size/date filters to one file's metadata."""
        try:
            fe = filters.get("ext")
            if fe:
                e = (ext or "").lstrip(".").casefold()
                if e == "jpg":
                    e = "jpeg"
                want = "jpeg" if fe == "jpg" else fe
                if e != want and not (want == "jpeg" and e in ("jpg", "jpeg")) \
                        and not (want in ("tif", "tiff") and e in ("tif", "tiff")):
                    return False
            fs = filters.get("size")
            if fs and size is not None:
                op, val = fs
                if op == ">" and not (size > val):
                    return False
                if op == "<" and not (size < val):
                    return False
            fd = filters.get("days")
            if fd and mtime:
                age_days = max(0.0, (time.time() - mtime) / 86400.0)
                lo, hi = fd
                if not (lo <= age_days < hi):
                    return False
            return True
        except Exception:
            return True    # a broken stat must never hide a file from search

    # ------------------------------------------------------------------
    # 7. CACHE + LOG
    # ------------------------------------------------------------------
    def cached(self, key):
        return self._cache.get(key)

    def remember(self, key, results):
        if len(self._cache) > 64:
            self._cache.clear()
        self._cache[key] = results

    def clear_cache(self):
        self._cache.clear()

    @staticmethod
    def log(query, duration_s, n_results, error=""):
        """Append one search to a small self-capping log file. Never raises."""
        try:
            line = "%s\t%r\t%.0fms\t%d results%s\n" % (
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), query,
                duration_s * 1000.0, n_results,
                ("\tERROR: %s" % error) if error else "")
            if os.path.exists(SEARCH_LOG) and os.path.getsize(SEARCH_LOG) > 200000:
                try:
                    with open(SEARCH_LOG, "r", encoding="utf-8") as f:
                        tail = f.readlines()[-1000:]
                    with open(SEARCH_LOG, "w", encoding="utf-8") as f:
                        f.writelines(tail)
                except Exception:
                    pass
            with open(SEARCH_LOG, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass


# A shared instance for the app (the engine itself is stateless per query; the
# instance only carries the small result cache).
ENGINE = SearchEngine()
