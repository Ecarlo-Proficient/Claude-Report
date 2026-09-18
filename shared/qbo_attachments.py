"""QBO transaction attachments (the uploaded bill scans) - the ONE resolver, shared.

QBO serves an attachment only through a `TempDownloadUri` that EXPIRES in minutes, so a
durable link can't be stored. Instead we keep a disk INDEX of every Attachable,
`(entity type, txn id) -> [{Id, FileName}]` (a company-wide sweep - field-limited, paged
in parallel, ~40 s - cached a week),
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


_FIELDS = "Id, FileName, AttachableRef"   # all the index stores; SELECT * is 40x the bytes per page
_PAGE = 1000                               # QBO's max page
_WORKERS = 5                               # pages in flight (QBO allows ~10 concurrent per realm)


def _sweep(access: str, company_id: str, api_get, progress=None) -> List[dict]:
    """Every Attachable as `{Id, FileName, AttachableRef}`: field-limited and paged in
    parallel, with a progress line every 10 pages. Measured 2026-09-15 on 79 pages:
    `SELECT *` = 5-10 s and 7 MB per page, sequential and silent = 6-13 min that read as a
    hang (`sync-all` "always hangs" - owner); this = ~2.4 s / 170 KB per page, 5 at a time
    = ~40 s. Same rows, same index. Pages are read by a fixed STARTPOSITION plan off
    COUNT(*), then the tail is re-read sequentially until a short page, so a file uploaded
    mid-sweep is still caught."""
    import math
    from concurrent.futures import ThreadPoolExecutor

    def q(sql: str) -> dict:
        return api_get(f"/v3/company/{company_id}/query", access, {"query": sql}).get("QueryResponse", {})

    def page(i: int) -> List[dict]:
        return q(f"SELECT {_FIELDS} FROM Attachable ORDERBY Id "
                 f"STARTPOSITION {i * _PAGE + 1} MAXRESULTS {_PAGE}").get("Attachable", [])

    total = int(q("SELECT COUNT(*) FROM Attachable").get("totalCount") or 0)
    pages = max(1, math.ceil(total / _PAGE))
    rows: List[dict] = []
    with ThreadPoolExecutor(max_workers=_WORKERS) as ex:
        for i, batch in enumerate(ex.map(page, range(pages)), 1):
            rows.extend(batch)
            if progress and (i % 10 == 0 or i == pages):
                progress(f"  attachments: page {i}/{pages} ({len(rows):,} of ~{total:,} records)")
    nxt = pages
    while len(batch) == _PAGE:                         # grew during the sweep: read on to a short page
        batch = page(nxt); rows.extend(batch); nxt += 1
    return rows


def build_index(access: str, company_id: str, api_get, force: bool = False, progress=None) -> Index:
    """The index from a fresh cache if we have one, else a full Attachable sweep (every
    scan ever uploaded - `_sweep`, ~40 s), which is then cached a week. `api_get` (the
    retrying GET) is injected to avoid importing the QBO client at module load; `progress`
    is a print-like callable for the page counter (None = silent). `force` skips the cache."""
    got = None if force else index_from_cache(company_id)
    if got is not None:
        return got
    by_key: Index = {}
    items = []
    for a in _sweep(access, company_id, api_get, progress):
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
