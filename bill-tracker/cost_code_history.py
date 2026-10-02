"""Persistent cost-code miscode history across sync-ap runs.

The owner (2026-09-01) wanted the cost-code audit "logged in the system": to SEE
how often the bill clerk miscodes over time, and what got FIXED between refreshes.
There is one bill clerk doing data entry, so the aggregate rate IS the clerk's rate
(per-person attribution isn't in the QBO data we pull - same wall the per-super
report hit).

State lives OUTSIDE the repo at <companyhealth>/cost_code_history.json - business
data, next to the other audit configs, never committed. Each REAL run (dry-run
bails before the audit, so it never mutates state):
  * every current miscode opens a new entry, or updates an existing one
    (first_seen / last_seen / runs_seen);
  * any previously-OPEN entry absent from this run flips to FIXED (fixed_on) -
    that is "what got fixed after refreshing";
  * a FIXED entry that reappears re-opens.
`update` returns a recap the workbook shows: open / new / fixed-this-run counts
plus a rolling "mistakes per run" rate. Pure state - no QBO, no network.

Key = bill_id + '|' + cost_code : stable while a bill keeps a given bad code, so
recoding SL51->SL1 next run flips the SL51 entry to FIXED. Recoding to ANOTHER
wrong code reads as one fixed + one new (bill-LINE identity isn't tracked; the
clerk metric still counts a fresh mistake, which is the point).
"""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Dict, List, Optional

SCHEMA = 1
_RUN_KEEP = 120          # cap the run-stamp trail so the file can't grow forever


def _key(f: dict) -> str:
    return f"{f.get('bill_id', '') or ''}|{(f.get('cost_code') or '').strip()}"


def _iso(d) -> str:
    if isinstance(d, (dt.date, dt.datetime)):
        return d.isoformat()[:10]
    return str(d or "")[:10]


def _as_date(s) -> Optional[dt.date]:
    try:
        return dt.date.fromisoformat(str(s)[:10])
    except (ValueError, TypeError):
        return None


def load(path: Path) -> dict:
    """Read the history JSON. Missing / corrupt / first-run → a fresh shell."""
    try:
        data = json.loads(Path(path).read_text())
    except (FileNotFoundError, ValueError, OSError):
        data = None
    if not isinstance(data, dict) or not isinstance(data.get("entries"), dict):
        return {"schema": SCHEMA, "runs": [], "entries": {}}
    data.setdefault("schema", SCHEMA)
    data.setdefault("runs", [])
    return data


def save(path: Path, hist: dict) -> None:
    """Atomic-ish write (temp + replace) so a crash can't shred the log."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(hist, indent=2, sort_keys=True, default=str))
    tmp.replace(path)


def update(hist: dict, flags: List[dict], today: dt.date) -> dict:
    """Fold this run's cost-code flags into `hist` (mutates it). Returns a recap
    dict: {open, new, fixed, reopened, rate, runs}."""
    entries: Dict[str, dict] = hist.setdefault("entries", {})
    tstr = today.isoformat()

    # Collapse this run's flags to one record per key (sum amount across same-key
    # lines - two SL51 lines on one bill are one "this bill has an SL51 miscode").
    current: Dict[str, dict] = {}
    for f in flags:
        if not f.get("bill_id"):
            continue
        k = _key(f)
        amt = round(float(f.get("amount") or 0.0), 2)
        cur = current.get(k)
        if cur:
            cur["amount"] = round(cur["amount"] + amt, 2)
            continue
        current[k] = {
            "vendor": f.get("vendor", "") or "",
            "bill_id": f.get("bill_id", "") or "",
            "bill_doc": f.get("bill_doc", "") or "",
            "bill_date": _iso(f.get("date")),
            "cost_code": (f.get("cost_code") or "").strip(),
            "project": f.get("project", "") or "",
            "reason": f.get("reason", "") or "",
            "vtype": f.get("vtype", "") or "",
            "amount": amt,
        }

    new_ct = reopened_ct = 0
    for k, c in current.items():
        e = entries.get(k)
        if e is None:
            c.update({"first_seen": tstr, "last_seen": tstr, "runs_seen": 1,
                      "status": "open", "fixed_on": "", "reopened_on": ""})
            entries[k] = c
            new_ct += 1
            continue
        was_fixed = e.get("status") == "fixed"
        e.update({"vendor": c["vendor"], "bill_doc": c["bill_doc"],
                  "bill_date": c["bill_date"], "project": c["project"],
                  "reason": c["reason"], "vtype": c["vtype"],
                  "amount": c["amount"], "last_seen": tstr, "status": "open"})
        e["runs_seen"] = int(e.get("runs_seen", 0)) + 1
        if was_fixed:
            e["reopened_on"] = tstr
            e["fixed_on"] = ""
            reopened_ct += 1

    fixed_ct = 0
    for k, e in entries.items():
        if k not in current and e.get("status") == "open":
            e["status"] = "fixed"
            e["fixed_on"] = tstr
            fixed_ct += 1

    open_ct = sum(1 for e in entries.values() if e.get("status") == "open")
    runs = hist.setdefault("runs", [])
    runs.append({"date": tstr, "open": open_ct, "new": new_ct,
                 "fixed": fixed_ct, "reopened": reopened_ct})
    hist["runs"] = runs[-_RUN_KEEP:]
    rate = round(sum(r.get("new", 0) for r in hist["runs"]) / max(len(hist["runs"]), 1), 1)
    return {"open": open_ct, "new": new_ct, "fixed": fixed_ct,
            "reopened": reopened_ct, "rate": rate, "runs": len(hist["runs"])}


def to_rows(hist: dict, today: dt.date, fixed_window_days: int = 60
            ) -> List[list]:
    """Sheet rows: every OPEN entry + entries FIXED within the last
    `fixed_window_days` (older fixes stay in the JSON but drop off the sheet so it
    stays lean). Open first (oldest first_seen first - longest-standing on top),
    then most-recently-fixed. Last element is the bill_id for the QBO link."""
    cutoff = today - dt.timedelta(days=fixed_window_days)
    show = []
    for e in hist.get("entries", {}).values():
        st = e.get("status")
        if st == "open":
            show.append((0, e))
        elif st == "fixed":
            fx = _as_date(e.get("fixed_on"))
            if fx and fx >= cutoff:
                show.append((1, e))
    # sort: open (0) before fixed (1); open by first_seen asc; fixed by fixed_on desc
    show.sort(key=lambda t: (
        t[0],
        (_as_date(t[1].get("first_seen")) or today) if t[0] == 0
        else -(_as_date(t[1].get("fixed_on")) or today).toordinal()
    ))
    rows = []
    for _, e in show:
        status = "OPEN" if e.get("status") == "open" else "FIXED"
        rows.append([
            status, e.get("vendor", ""), e.get("bill_doc", ""),
            e.get("cost_code", ""), e.get("reason", ""), e.get("project", ""),
            _as_date(e.get("first_seen")), _as_date(e.get("last_seen")),
            int(e.get("runs_seen", 0) or 0), _as_date(e.get("fixed_on")),
            e.get("bill_id", ""),
        ])
    return rows


def recap_note(recap: dict, today: dt.date) -> str:
    """The italic caption shown on the sheet - owner-facing, mm/dd/yyyy dates."""
    d = today.strftime("%m/%d/%Y")
    tail = (f" · {recap['reopened']} reopened" if recap.get("reopened") else "")
    return (f"As of {d}: {recap['open']} open cost-code miscodes · "
            f"{recap['new']} new this run · {recap['fixed']} fixed this run{tail}  |  "
            f"bill-clerk coding-error rate: {recap['rate']} new/run over the last "
            f"{recap['runs']} run(s).  OPEN = still miscoded in QBO · FIXED = corrected "
            f"since a prior run (recent fixes shown; full log kept off-sheet).")
