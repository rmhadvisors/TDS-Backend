"""
Nature-of-payment keywords added by users, for words the statutory list (tds_statutory_rates) does not know.

Stored in backend/data/user_keywords.json (override the path with TDS_USER_KEYWORDS_FILE) as
    [{"keyword": "HOUSEKEEPING", "section": "194C"}, ...]

A user keyword matches any ledger / party text the same way a statutory keyword does (start of a word, so
"HOUSEKEEP" hits "Housekeeping Charges" in every company). It is consulted only when no statutory keyword matches
the text, and keywords themselves come after the data's own signals (TDS ledger, masters, earlier deductions) —
so adding one never changes a section those already decide; it only fills in what would otherwise be a guess.
"""
import json
import os
import re
import threading
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from app.services.tds_statutory_rates import (
    STATUTORY_TDS_REGISTRY,
    StatutoryTDSRule,
    _GLOBAL_EXCL,
    _TDS_ACCOUNT,
    _kw_regex,
    get_statutory_rule_by_section,
)

_DEFAULT_FILE = Path(__file__).resolve().parents[2] / "data" / "user_keywords.json"
_LOCK = threading.Lock()
_CACHE: Dict[str, object] = {"path": None, "mtime": None, "entries": [], "compiled": []}


def _file() -> Path:
    return Path(os.environ.get("TDS_USER_KEYWORDS_FILE") or _DEFAULT_FILE)


def section_options() -> List[Dict[str, str]]:
    """Sections a user can assign a keyword to: code (sub-section where the rule has one) and description."""
    out, seen = [], set()
    for r in STATUTORY_TDS_REGISTRY:
        code = r.sub_section or r.section_code
        if code not in seen:
            seen.add(code)
            out.append({"section": code, "description": r.description})
    return out


def _normalize_keyword(keyword: str) -> str:
    return " ".join((keyword or "").upper().split())


def _load() -> List[Tuple[str, StatutoryTDSRule, "re.Pattern[str]"]]:
    path = _file()
    try:
        mtime = path.stat().st_mtime
    except OSError:
        mtime = None
    if _CACHE["path"] == str(path) and _CACHE["mtime"] == mtime:
        return _CACHE["compiled"]  # type: ignore[return-value]
    entries: List[Dict[str, str]] = []
    if mtime is not None:
        try:
            entries = [e for e in json.loads(path.read_text(encoding="utf-8")) if isinstance(e, dict)]
        except (OSError, ValueError):
            entries = []
    compiled = []
    for e in entries:
        kw, rule = _normalize_keyword(e.get("keyword", "")), get_statutory_rule_by_section(e.get("section", ""))
        if kw and rule:
            compiled.append((kw, rule, _kw_regex(kw)))
    # longest keyword first: "SECURITY GUARD" before "SECURITY"
    compiled.sort(key=lambda x: -len(x[0]))
    _CACHE.update(path=str(path), mtime=mtime, entries=entries, compiled=compiled)
    return compiled


def list_user_keywords() -> List[Dict[str, str]]:
    return [{"keyword": kw, "section": rule.sub_section or rule.section_code} for kw, rule, _ in sorted(_load(), key=lambda x: x[0])]


def _save(entries: List[Dict[str, str]]) -> None:
    path = _file()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(entries, indent=2), encoding="utf-8")
    tmp.replace(path)
    # Two saves within the file system's timestamp resolution leave the mtime unchanged, so the cache is dropped
    # here rather than trusting the mtime check in _load
    _CACHE["path"] = None


def add_user_keyword(keyword: str, section: str) -> Dict[str, str]:
    """Adds (or re-points) a keyword. Raises ValueError for an unusable keyword or unknown section."""
    kw = _normalize_keyword(keyword)
    if len(re.sub(r"[^A-Z0-9]", "", kw)) < 3:
        raise ValueError("Keyword must have at least 3 letters or digits.")
    if _TDS_ACCOUNT.match(kw) or any(p.search(kw) for p in _GLOBAL_EXCL):
        raise ValueError(f"'{kw}' names a TDS account or a non-TDS item and cannot be used as a keyword.")
    rule = get_statutory_rule_by_section(section)
    if rule is None:
        raise ValueError(f"Unknown TDS section '{section}'.")
    code = rule.sub_section or rule.section_code
    with _LOCK:
        entries = [{"keyword": k, "section": r.sub_section or r.section_code} for k, r, _ in _load() if k != kw]
        entries.append({"keyword": kw, "section": code})
        _save(entries)
    return {"keyword": kw, "section": code}


def remove_user_keyword(keyword: str) -> bool:
    kw = _normalize_keyword(keyword)
    with _LOCK:
        current = _load()
        if not any(k == kw for k, _, _ in current):
            return False
        _save([{"keyword": k, "section": r.sub_section or r.section_code} for k, r, _ in current if k != kw])
    return True


def match_user_rule(text: Optional[str]) -> Optional[Tuple[StatutoryTDSRule, str]]:
    """(rule, keyword) of the longest user keyword found in `text`, or None. TDS accounts and globally excluded
    (non-TDS) texts never match, as for statutory keywords."""
    if not text:
        return None
    u = text.upper()
    if _TDS_ACCOUNT.match(u) or any(p.search(u) for p in _GLOBAL_EXCL):
        return None
    for kw, rule, pat in _load():
        if pat.search(u):
            return rule, kw
    return None
