"""Suggested-name library — the scalable, local-first data model behind
ApneScan's redesigned Rename system.

Pure Python (no PyQt, no network). All data lives in memory and is persisted
through a caller-supplied save callback (ApneScan stores it inside its normal
config file, so nothing new to install and it loads at startup).

Backward compatibility: the old rename suggestions were a plain list of strings
under ``config["name_history"]``. On first load we migrate those strings into
rich records and we keep ``name_history`` mirrored (ordered by priority/usage)
so every existing reader — auto-naming, the line-edit completer, etc. — keeps
working unchanged.

Each suggestion record::

    {
      "id": "n_ab12cd34",
      "name": "OPD",
      "category": "Medical",
      "keyword": "opd",
      "usageCount": 14,
      "lastUsed": 1734900000.0,
      "createdAt": 1734000000.0,
      "updatedAt": 1734900000.0,
      "priority": 1,            # smaller = higher (drag order); 0 = unset/auto
      "isActive": True,
      "source": "user",        # user | seed | import | ocr
      "confidence": None,      # 0..100 for ocr-sourced, else None
      "aliases": ["OPD", "Out Patient Department"]
    }

The engine is deliberately independent of storage details so it can be unit
tested and reused. Search/suggestions are O(n) over an in-memory list — instant
for the hundreds/thousands of names a real user accumulates; never touches disk
or network on a keystroke.
"""

import re
import time
import uuid

__all__ = ["NameLibrary", "DEFAULT_CATEGORIES", "TEMPLATE_VARS",
           "expand_template", "clean_custom_name", "invalid_name_reason",
           "WIN_FORBIDDEN"]

# Windows-forbidden filename characters (shown to the user verbatim).
WIN_FORBIDDEN = r'\/:*?"<>|'
_FORBIDDEN_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')

DEFAULT_CATEGORIES = ["Medical", "ECHS", "Financial", "General", "Custom"]

# Fresh-install starter suggestions (only seeded when the user has NOTHING yet —
# never added on top of an existing library / name_history).
_SEED = [
    ("OPD", "Medical", "opd"),
    ("Prescription", "Medical", "prescription"),
    ("Diagnosis", "Medical", "diagnosis"),
    ("Investigation", "Medical", "investigation"),
    ("Lab Report", "Medical", "lab report"),
    ("X-Ray", "Medical", "x-ray"),
    ("Discharge Summary", "Medical", "discharge"),
    ("Doctor Certificate", "Medical", "certificate"),
    ("ECHS Card", "ECHS", "echs"),
    ("Referral", "ECHS", "referral"),
    ("Invoice", "Financial", "invoice"),
    ("Medical Bill", "Financial", "bill"),
]

# Template variables the naming-template expander understands. Extend freely —
# unknown variables are left blank and the separators are tidied afterwards.
TEMPLATE_VARS = [
    "PatientName", "DocumentType", "Date", "Time", "PageNo",
    "TotalPages", "FolderName", "Category", "DoctorName",
]

_VAR_RE = re.compile(r"\{([A-Za-z0-9_]+)\}")


def clean_custom_name(s, max_len=50):
    """Trim, collapse whitespace and strip Windows-forbidden characters.
    Never raises. Returns a filename-friendly string (spaces kept)."""
    s = _FORBIDDEN_RE.sub("", s or "")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip(" .")
    if max_len and len(s) > max_len:
        s = s[:max_len].strip()
    return s


def invalid_name_reason(s, existing_lower=None, max_len=50):
    """Return a short human reason the *raw* name is invalid, or None if OK.
    ``existing_lower`` (optional set of lowercased names) flags duplicates."""
    raw = s or ""
    if not raw.strip():
        return "empty"
    bad = sorted(set(ch for ch in raw if _FORBIDDEN_RE.match(ch) and ch.strip()))
    if bad:
        return "badchars:" + "".join(bad)
    if max_len and len(raw.strip()) > max_len:
        return "toolong"
    if existing_lower is not None and raw.strip().lower() in existing_lower:
        return "duplicate"
    return None


def expand_template(tpl, ctx):
    """Expand a naming template like ``{PatientName}_{DocumentType}_{Date}``.

    ``ctx`` maps variable names to values; missing/blank variables vanish and the
    surrounding separators are cleaned so you never get ``Name__Date`` or a
    leading/trailing ``_``. Returns a filesystem-safe string (forbidden chars
    removed, spaces kept)."""
    ctx = ctx or {}

    def _sub(m):
        return str(ctx.get(m.group(1), "") or "").strip()

    out = _VAR_RE.sub(_sub, tpl or "")
    # collapse separators left behind by empty variables
    out = re.sub(r"[ _\-]*_[ _\-]*", "_", out) if "_" in (tpl or "") else out
    out = re.sub(r"__+", "_", out)
    out = re.sub(r"--+", "-", out)
    out = re.sub(r"\s+", " ", out)
    out = out.strip(" _-.")
    return clean_custom_name(out, max_len=0)


def _now():
    return time.time()


def _new_id():
    return "n_" + uuid.uuid4().hex[:10]


def _norm(s):
    return re.sub(r"\s+", " ", (s or "").strip()).lower()


class NameLibrary:
    """In-memory suggested-name store. Construct with the app config dict and a
    save callback; everything else operates on records in memory and calls
    ``save`` when something changes."""

    def __init__(self, config, save_cb=None):
        self._cfg = config if isinstance(config, dict) else {}
        self._save_cb = save_cb
        self._items = []
        self._load()

    # ---- persistence -----------------------------------------------------
    def _load(self):
        raw = self._cfg.get("name_library")
        if isinstance(raw, list) and raw and isinstance(raw[0], dict):
            self._items = [self._coerce(r) for r in raw if r.get("name")]
        else:
            self._items = []
            self._migrate_from_history()
            if not self._items:
                self._seed_defaults()
        self._reindex()
        self._persist(mirror_only=True)   # ensure name_history mirror exists

    def _coerce(self, r):
        """Fill any missing fields on a stored record (forward/backward safe)."""
        name = (r.get("name") or "").strip()
        return {
            "id": r.get("id") or _new_id(),
            "name": name,
            "category": r.get("category") or "General",
            "keyword": (r.get("keyword") or "").strip(),
            "usageCount": int(r.get("usageCount") or 0),
            "lastUsed": r.get("lastUsed"),
            "createdAt": r.get("createdAt") or _now(),
            "updatedAt": r.get("updatedAt") or _now(),
            "priority": int(r.get("priority") or 0),
            "isActive": bool(r.get("isActive", True)),
            "source": r.get("source") or "user",
            "confidence": r.get("confidence"),
            "aliases": [a for a in (r.get("aliases") or []) if a],
        }

    def _migrate_from_history(self):
        """Turn the legacy ``name_history`` string list into rich records.
        Order in the list implies recency -> seed a descending usage/recency so
        the first few still rank first."""
        hist = self._cfg.get("name_history")
        if not isinstance(hist, list):
            return
        n = len(hist)
        for i, nm in enumerate(hist):
            if not isinstance(nm, str) or not nm.strip():
                continue
            self._items.append(self._coerce({
                "name": nm.strip(),
                "category": "General",
                "usageCount": max(1, n - i),      # earlier = used more
                "priority": i + 1,
                "source": "user",
            }))

    def _seed_defaults(self):
        t = _now()
        for i, (nm, cat, kw) in enumerate(_SEED):
            self._items.append(self._coerce({
                "name": nm, "category": cat, "keyword": kw,
                "priority": i + 1, "source": "seed", "createdAt": t,
            }))

    def _persist(self, mirror_only=False):
        # rich store
        if not mirror_only:
            self._cfg["name_library"] = self._items
        # backward-compat mirror: ordered active names (priority then usage)
        self._cfg["name_history"] = [it["name"] for it in self._ranked(active_only=True)]
        if self._save_cb:
            try:
                self._save_cb()
            except Exception:
                pass

    def _reindex(self):
        # normalise duplicate names (keep the most-used), stable ids
        seen = {}
        for it in self._items:
            k = _norm(it["name"])
            if not k:
                continue
            if k in seen:
                a = seen[k]
                a["usageCount"] = max(a["usageCount"], it["usageCount"]) + it["usageCount"] // 2
                if it.get("keyword") and not a.get("keyword"):
                    a["keyword"] = it["keyword"]
            else:
                seen[k] = it
        self._items = list(seen.values())

    # ---- lookup / ranking ------------------------------------------------
    def _ranked(self, active_only=True):
        items = [it for it in self._items if (it["isActive"] or not active_only)]

        def key(it):
            pr = it["priority"] if it["priority"] else 10 ** 6
            return (pr, -it["usageCount"], -(it["lastUsed"] or 0),
                    _norm(it["name"]))
        return sorted(items, key=key)

    def all(self, active_only=False):
        return list(self._ranked(active_only=active_only))

    def by_id(self, nid):
        for it in self._items:
            if it["id"] == nid:
                return it
        return None

    def categories(self):
        cats = list(DEFAULT_CATEGORIES)
        for it in self._items:
            c = it.get("category") or "General"
            if c not in cats:
                cats.append(c)
        return cats

    def frequently_used(self, limit=8):
        items = [it for it in self._items if it["isActive"] and it["usageCount"] > 0]
        items.sort(key=lambda it: (-it["usageCount"], _norm(it["name"])))
        return items[:limit]

    def recently_used(self, limit=8):
        items = [it for it in self._items if it["isActive"] and it["lastUsed"]]
        items.sort(key=lambda it: -(it["lastUsed"] or 0))
        return items[:limit]

    def search(self, q, category=None, active_only=False, limit=None):
        """Instant local filter. Matches name / keyword / aliases (substring,
        case-insensitive). Ranked: name-prefix first, then usage/priority."""
        q = _norm(q)
        cat = category if category and category != "All" else None
        out = []
        for it in self._ranked(active_only=active_only):
            if cat and (it.get("category") or "General") != cat:
                continue
            if q:
                hay = " ".join([it["name"], it.get("keyword", "")]
                               + it.get("aliases", [])).lower()
                if q not in hay:
                    continue
                pref = 0 if it["name"].lower().startswith(q) else (
                    1 if q in it["name"].lower() else 2)
            else:
                pref = 0
            out.append((pref, it))
        out.sort(key=lambda t: t[0])          # stable: keeps _ranked order within tier
        res = [it for _p, it in out]
        return res[:limit] if limit else res

    def suggestions(self, query="", limit=24, ocr_text=None):
        """Names for the Rename dialog chips. Frequently/recently used surface
        first; if OCR text is given, a detected document type is injected at the
        very top (future-ready, optional)."""
        base = self.search(query, active_only=True, limit=None)
        if not query and ocr_text:
            for nm, conf in self.ocr_suggestions(ocr_text):
                # move a matching library name to the front, else prepend a ghost
                hit = next((it for it in base if _norm(it["name"]) == _norm(nm)), None)
                if hit:
                    base.remove(hit)
                    base.insert(0, hit)
        return base[:limit] if limit else base

    def ocr_suggestions(self, text):
        """Future-ready hook: map OCR text -> [(name, confidence)]. Uses the
        offline document-type detector when available; degrades to [] silently.
        Never mandatory."""
        out = []
        try:
            from apnescan_lib.naming_engine import _ni_detect_type
            label, conf = _ni_detect_type(text or "")
            if label and label not in ("Document", ""):
                out.append((label, conf))
        except Exception:
            pass
        # keyword hits from the user's own library
        low = (text or "").lower()
        if low:
            for it in self._items:
                kw = (it.get("keyword") or "").strip().lower()
                if kw and kw in low:
                    out.append((it["name"], it.get("confidence") or 60))
        # de-dup keep first/highest
        seen, uniq = set(), []
        for nm, conf in out:
            k = _norm(nm)
            if k in seen:
                continue
            seen.add(k)
            uniq.append((nm, conf))
        return uniq[:6]

    # ---- mutations -------------------------------------------------------
    def add(self, name, category="General", keyword="", aliases=None,
            source="user", confidence=None, priority=0, save=True):
        name = clean_custom_name(name, max_len=0)
        if not name:
            return None
        existing = self._find(name)
        if existing:
            # never duplicate — just enrich
            if keyword and not existing.get("keyword"):
                existing["keyword"] = keyword.strip()
            if category and category != "General":
                existing["category"] = category
            if aliases:
                for a in aliases:
                    if a and a not in existing["aliases"]:
                        existing["aliases"].append(a)
            existing["updatedAt"] = _now()
            if save:
                self._persist()
            return existing
        rec = self._coerce({
            "name": name, "category": category or "General",
            "keyword": (keyword or "").strip(),
            "aliases": aliases or [], "source": source,
            "confidence": confidence, "priority": priority,
            "createdAt": _now(), "updatedAt": _now(),
        })
        self._items.append(rec)
        if save:
            self._persist()
        return rec

    def edit(self, nid, **fields):
        it = self.by_id(nid)
        if not it:
            return None
        if "name" in fields:
            newname = clean_custom_name(fields["name"], max_len=0)
            if newname:
                clash = self._find(newname)
                if clash and clash is not it:
                    return None            # would duplicate — reject
                it["name"] = newname
        for f in ("category", "keyword", "source"):
            if f in fields and fields[f] is not None:
                it[f] = fields[f]
        if "priority" in fields and fields["priority"] is not None:
            it["priority"] = int(fields["priority"])
        if "isActive" in fields and fields["isActive"] is not None:
            it["isActive"] = bool(fields["isActive"])
        if "aliases" in fields and fields["aliases"] is not None:
            it["aliases"] = [a for a in fields["aliases"] if a]
        if "confidence" in fields:
            it["confidence"] = fields["confidence"]
        it["updatedAt"] = _now()
        self._persist()
        return it

    def delete(self, nid):
        n = len(self._items)
        self._items = [it for it in self._items if it["id"] != nid]
        if len(self._items) != n:
            self._persist()
            return True
        return False

    def bump_usage(self, name, category=None, save=True):
        """Record that ``name`` was actually used for a rename — creates the
        record if new, increments usage, stamps lastUsed."""
        name = clean_custom_name(name, max_len=0)
        if not name:
            return None
        it = self._find(name)
        if not it:
            it = self.add(name, category=category or "General", save=False)
        it["usageCount"] += 1
        it["lastUsed"] = _now()
        it["updatedAt"] = _now()
        if save:
            self._persist()
        return it

    def set_order(self, ordered_ids):
        """Apply a drag-and-drop priority order (first = highest)."""
        rank = {nid: i + 1 for i, nid in enumerate(ordered_ids)}
        for it in self._items:
            if it["id"] in rank:
                it["priority"] = rank[it["id"]]
        self._persist()

    def _find(self, name):
        k = _norm(name)
        for it in self._items:
            if _norm(it["name"]) == k:
                return it
            if any(_norm(a) == k for a in it.get("aliases", [])):
                return it
        return None

    # ---- import / export -------------------------------------------------
    def export_data(self):
        return {
            "app": "ApneScan",
            "kind": "naming_library",
            "version": 1,
            "exportedAt": _now(),
            "names": [dict(it) for it in self._items],
        }

    def validate_import(self, data):
        """Return (ok, records|reason). Never trusts the file blindly."""
        if not isinstance(data, dict):
            return (False, "not a valid library file")
        if data.get("kind") != "naming_library":
            return (False, "not an ApneScan naming library")
        names = data.get("names")
        if not isinstance(names, list):
            return (False, "no names in file")
        good = [self._coerce(r) for r in names
                if isinstance(r, dict) and (r.get("name") or "").strip()]
        if not good:
            return (False, "file has no usable names")
        return (True, good)

    def import_data(self, data, mode="merge"):
        """mode='merge' keeps existing (adds new, sums usage on clashes);
        mode='replace' swaps the whole library. Returns (added, updated)."""
        ok, payload = self.validate_import(data)
        if not ok:
            return (0, 0, payload)
        added = updated = 0
        if mode == "replace":
            self._items = payload
            added = len(payload)
        else:
            for rec in payload:
                cur = self._find(rec["name"])
                if cur:
                    cur["usageCount"] = max(cur["usageCount"], rec["usageCount"])
                    if rec.get("keyword") and not cur.get("keyword"):
                        cur["keyword"] = rec["keyword"]
                    for a in rec.get("aliases", []):
                        if a not in cur["aliases"]:
                            cur["aliases"].append(a)
                    cur["updatedAt"] = _now()
                    updated += 1
                else:
                    self._items.append(rec)
                    added += 1
        self._reindex()
        self._persist()
        return (added, updated, None)

    # ---- rename history --------------------------------------------------
    def add_history(self, old, new, page="", doc=""):
        h = self._cfg.setdefault("rename_history", [])
        h.insert(0, {
            "old": old or "", "new": new or "", "ts": _now(),
            "page": page or "", "doc": doc or "",
        })
        del h[200:]                 # keep last 200
        if self._save_cb:
            try:
                self._save_cb()
            except Exception:
                pass

    def history(self, limit=100):
        h = self._cfg.get("rename_history") or []
        return h[:limit]

    def pop_history(self):
        h = self._cfg.get("rename_history") or []
        if not h:
            return None
        item = h.pop(0)
        if self._save_cb:
            try:
                self._save_cb()
            except Exception:
                pass
        return item
