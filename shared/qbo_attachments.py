"""QBO transaction attachments (the uploaded bill scans) - the ONE resolver, shared.

QBO serves an attachment only through a `TempDownloadUri` that EXPIRES in minutes, so a
durable link can't be stored. Instead we keep a disk INDEX of every Attachable,
`(entity type, txn id) -> [{Id, FileName}]` (a slow company-wide sweep, cached a week),
and fetch a FRESH `TempDownloadUri` per file at click-time by re-reading that attachable
by its Id.

This logic was proven in project-pnl (bill-scan links on the P&L, Mac + Windows); it is
lifted here so the ledger dashboard can reuse the SAME on-disk cache (never a second
sweep). project-pnl keeps its own copy until it can be pointed at this module.

Read-only on QBO. No secrets, no realm printed.
"""
from __future__ import annotations

import datetime as dt
import glob
import json
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

_TTL_S = 7 * 24 * 3600
# Reuse the P&L-built cache when it's there (same realm) so we never double-sweep;
# a fresh sweep of our own is written to the clean location.
_DIRS = (Path.home() / "Library/Logs/Proficient/qbo-attachments",
         Path.home() / "Library/Logs/Proficient/project-pnl")

Index = Dict[Tuple[str, str], List[dict]]


def _read(path: Path) -> Optional[Index]:
    try:
        if not (path.exists() and time.time() - path.stat().st_mtime < _TTL_S):
            return None
        raw = json.loads(path.read_text())
        by_key: Index = {}
        for item in raw.get("items", []):
            for etype, evalue in item["refs"]:
                by_key.setdefault((etype, evalue), []).append(
                    {"Id": item["id"], "FileName": item["file"]})
        return by_key
    except Exception:                                  # noqa: BLE001 - unreadable → treat as absent
        return None


def index_from_cache(company_id: str = "") -> Optional[Index]:
    """The `(etype, txnId) -> [{Id, FileName}]` index from a FRESH disk cache, or None.
    No QBO call. Blank company_id → take the freshest realm's cache on disk."""
    if company_id:
        for d in _DIRS:
            got = _read(d / f"attachable_index_{company_id}.json")
            if got is not None:
                return got
        return None
    cands = []
    for d in _DIRS:
        cands += [Path(c) for c in glob.glob(str(d / "attachable_index_*.json"))]
    cands = [c for c in cands if time.time() - c.stat().st_mtime < _TTL_S]
    if not cands:
        return None
    cands.sort(key=lambda c: c.stat().st_mtime, reverse=True)
    return _read(cands[0])


def build_index(access: str, company_id: str, query_all, force: bool = False) -> Index:
    """The index from a fresh cache if we have one, else a full Attachable sweep (slow:
    every scan ever uploaded), which is then cached a week. `query_all` is injected to
    avoid importing the QBO client at module load. `force` skips the cache (a new sweep)."""
    got = None if force else index_from_cache(company_id)
    if got is not None:
        return got
    by_key: Index = {}
    items = []
    for a in query_all(access, company_id, "Attachable"):
        if not a.get("FileName"):
            continue                                   # a bare note, not a file
        refs = [((r.get("EntityRef") or {}).get("type"),
                 (r.get("EntityRef") or {}).get("value"))
                for r in a.get("AttachableRef") or []]
        refs = [x for x in refs if x[0] and x[1]]
        if not refs:
            continue
        items.append({"id": a["Id"], "file": a["FileName"], "refs": refs})
        for key in refs:
            by_key.setdefault(key, []).append({"Id": a["Id"], "FileName": a["FileName"]})
    try:
        d = _DIRS[0]
        d.mkdir(parents=True, exist_ok=True)
        (d / f"attachable_index_{company_id}.json").write_text(
            json.dumps({"fetched": dt.datetime.now().isoformat(), "items": items}))
    except OSError:
        pass
    return by_key


def count_for(idx: Index, txn_id: str, tx_type: str = "Bill") -> int:
    """How many files a transaction has in the index (no QBO call)."""
    etype = "Purchase" if tx_type == "Expense" else tx_type
    return len(idx.get((etype, str(txn_id)), []))


def fresh_links(access: str, company_id: str, idx: Index, txn_id: str,
                api_get, tx_type: str = "Bill") -> List[dict]:
    """`[{name, url}]` for a transaction's attachments, each url a FRESH (minutes-lived)
    `TempDownloadUri` fetched by re-reading the attachable by Id. `api_get` is injected
    (the retrying GET). Skips any file whose fresh link can't be fetched."""
    etype = "Purchase" if tx_type == "Expense" else tx_type
    out = []
    for a in idx.get((etype, str(txn_id)), []):
        try:
            fresh = api_get(f"/v3/company/{company_id}/attachable/{a['Id']}", access)
            uri = (fresh.get("Attachable") or {}).get("TempDownloadUri")
            if uri:
                out.append({"name": a.get("FileName") or "attachment", "url": uri})
        except Exception:                              # noqa: BLE001 - one bad file shouldn't sink the rest
            continue
    return out
