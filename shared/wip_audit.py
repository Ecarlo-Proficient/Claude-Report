"""
wip_audit.py - the WIP change log: every contract / CO / ETC / billed / costs change on a WIP
tab, written the moment the tab is written, readable per job "just like QuickBooks' audit
log" (the owner 2026-09-17).

WHO WRITES: `wip/wip_writer.write_test_cp` (the ONE guarded tab writer) after a successful
save - it diffs the tab as it stood before the rewrite against the rows it wrote, one entry
per changed field, with the field's SOURCE (the document the reader stamped, "typed on the
tab (kept)" for an owner edit, "QuickBooks" for money fields) and the run. A line that
appears or leaves the tab is logged as field 'line'. Hand backfills (one-off scripts) use
`log_changes` too, with their own actor.

WHO READS: the ledger dashboard (`/api/wip/audit?no=`) for the change-log section on the
RP review job page and the project page. `shared/` is the only importable common code, so
the wip tool (writer) and the ledger (reader) share this one module - no tool imports another.

Absent-safe: no ledger DB -> the writer logs nothing and says so once; readers return [].
"""
from __future__ import annotations

import os
import sqlite3
from pathlib import Path
from typing import Iterable, List, Optional

LEDGER_DB = Path(os.environ.get("ACB_LEDGER_DB",
                                Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3"))

FIELDS = ("contract", "approved_cos", "etc", "co_costs", "billed", "costs", "retainage", "line")
_COLS = "at, tab, project_no, field, old, new, source, actor, run, note"


def _ensure(con: sqlite3.Connection) -> None:
    con.execute("CREATE TABLE IF NOT EXISTS wip_field_audit ("
                "id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, tab TEXT, project_no TEXT NOT NULL, "
                "field TEXT NOT NULL, old REAL, new REAL, source TEXT, actor TEXT, run TEXT, note TEXT)")
    con.execute("CREATE INDEX IF NOT EXISTS ix_wip_field_audit_proj ON wip_field_audit (project_no, id)")


def log_changes(entries: Iterable[dict]) -> int:
    """Append entries {at, tab, project_no, field, old, new, source, actor, run, note}.
    Returns how many were written; 0 (and no error) when there is no ledger DB yet."""
    rows = [(e.get("at"), e.get("tab"), str(e.get("project_no") or "").strip().upper(), e.get("field"),
             e.get("old"), e.get("new"), e.get("source"), e.get("actor"), e.get("run"), e.get("note"))
            for e in entries]
    rows = [r for r in rows if r[0] and r[2] and r[3]]
    if not rows or not LEDGER_DB.parent.exists():
        return 0
    con = sqlite3.connect(str(LEDGER_DB))
    try:
        _ensure(con)
        con.executemany(f"INSERT INTO wip_field_audit ({_COLS}) VALUES (?,?,?,?,?,?,?,?,?,?)", rows)
        con.commit()
    finally:
        con.close()
    return len(rows)


def read(project_no: Optional[str] = None, limit: int = 500, since: Optional[str] = None) -> List[dict]:
    """Entries newest first, for one job or for every job. [] when there is no DB / table."""
    if not LEDGER_DB.exists():
        return []
    q = f"SELECT id, {_COLS} FROM wip_field_audit"
    where, args = [], []
    if project_no:
        where.append("project_no = ?"); args.append(project_no.strip().upper())
    if since:
        where.append("at >= ?"); args.append(since)
    if where:
        q += " WHERE " + " AND ".join(where)
    q += " ORDER BY id DESC LIMIT ?"
    args.append(int(limit))
    try:
        con = sqlite3.connect(f"file:{LEDGER_DB}?mode=ro", uri=True)
        try:
            rows = con.execute(q, args).fetchall()
        finally:
            con.close()
    except sqlite3.OperationalError:
        return []
    keys = ["id"] + [c.strip() for c in _COLS.split(",")]
    return [dict(zip(keys, r)) for r in rows]


def diff_entries(prior: dict, rows, tab: str, run: str, at: str, actor: str = "sync",
                 owner_edits: Optional[dict] = None, source_of=None) -> List[dict]:
    """The entries for one tab write. `prior` = {PN: {contract, approved_cos, etc, co_costs,
    billed, costs}} as the tab stood; `rows` = the CpRow objects being written; `source_of(row,
    key)` -> a source string or None. A PN with no prior is 'line added'; a prior PN not written
    is 'line removed'."""
    owner_edits = owner_edits or {}
    key_attr = {"contract": "base_contract", "approved_cos": "co_revenue", "etc": "base_etc",
                "co_costs": "co_cost_override", "billed": "billed_to_date", "costs": "costs_to_date",
                "retainage": "retainage_held"}
    edit_field = {"base_contract": "contract", "co_revenue": "approved_cos", "base_etc": "etc",
                  "co_cost_estimate": "co_costs", "billed_to_date": "billed", "costs_to_date": "costs",
                  "retainage_held": "retainage"}
    out, seen = [], set()
    for row in rows:
        pn = str(getattr(row, "project_num", "") or "").strip().upper()
        if not pn:
            continue
        seen.add(pn)
        p = prior.get(pn)
        edits = {edit_field.get(k, k): v for k, v in (owner_edits.get(pn) or {}).items()}
        if p is None:
            k, e = getattr(row, "base_contract", None), getattr(row, "base_etc", None)
            out.append({"at": at, "tab": tab, "project_no": pn, "field": "line", "old": None, "new": 1,
                        "source": (source_of(row, "orig_contract") if source_of else None),
                        "actor": actor, "run": run,
                        "note": f"added · contract {k:,.2f}" if k is not None else "added · no contract"
                                + (f" · ETC {e:,.2f}" if e is not None else "")})
            continue
        for field, attr in key_attr.items():
            old, new = p.get(field), getattr(row, attr, None)
            if field == "co_costs" and new is None:
                new = getattr(row, "co_cost_estimate", None)
            if (old is None and new is None) or (old is not None and new is not None and abs(float(old) - float(new)) < 0.005):
                continue
            if field in edits:
                src = "typed on the tab (kept by the sync)"
            elif field in ("billed", "costs", "retainage"):
                src = "QuickBooks"
            else:
                src = (source_of(row, {"contract": "orig_contract", "approved_cos": "approved_cos", "etc": "orig_etc",
                                        "co_costs": "co_costs"}[field]) if source_of else None) or "the reader"
            out.append({"at": at, "tab": tab, "project_no": pn, "field": field, "old": old, "new": new,
                        "source": src, "actor": actor, "run": run, "note": None})
    for pn, p in prior.items():
        if pn not in seen:
            out.append({"at": at, "tab": tab, "project_no": pn, "field": "line", "old": 1, "new": None,
                        "source": None, "actor": actor, "run": run,
                        "note": "removed from the tab" + (f" · contract was {p['contract']:,.2f}" if p.get("contract") is not None else "")})
    return out
