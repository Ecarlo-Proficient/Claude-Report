#!/usr/bin/env python3
"""
draw_periods.py - Draw periods on the ledger (the owner 10/11/2026: "a way to change period draw dates by moving it in
a horizontal <> way and it changes the dates as it goes for all, and then a way to edit individual periods themselves
without it overlapping the other ones" - "never write" to QuickBooks - "a feature inside the page ... put save but
only on the ledger").

A draw's period is the "(Period:MM/DD/YYYY - MM/DD/YYYY)" on its invoice's memo in QuickBooks. The ledger's project
page lets the owner slide every period together or type one draw's dates between its neighbours, and SAVE them -
on the ledger only (`draw_period`, one row per invoice of the draw): the project page then uses the saved window
as the draw's period (the Draw overview, the subs matched into the draw) and shows "as invoiced" beside it.
QuickBooks, the Bill Tracker and Notion are never touched; a saved window that equals the invoice's is not kept.

    saved(con, project)               {invoice id: (start, end, saved_at)} - what the ledger holds for the job
    save(con, project, periods)       periods = [{id, start, end, inv_start, inv_end, label}] for EVERY draw of the
                                      job that carries a period; refused on an overlap; replaces the job's rows
                                      with the windows that differ from the invoice; returns how many are kept
    period_of / memo_with / check     the period vocabulary (read, rewrite in place, the overlap rule)

Command line:  ledger/draw_periods.py MFD325            every draw's period on the job, from the mirror (read-only)
"""
from __future__ import annotations

import datetime as dt
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_mirror  # noqa: E402

_PERIOD = re.compile(r"Period\s*:\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*-\s*(\d{1,2})/(\d{1,2})/(\d{4})", re.I)
MAX_SPAN_DAYS = 120          # a draw period longer than this is a typo, not a billing window


def _iso(m: str, d: str, y: str) -> str:
    return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"


def _us(iso: str) -> str:
    y, m, d = iso.split("-")
    return f"{int(m):02d}/{int(d):02d}/{y}"


def _day(iso: str) -> dt.date:
    return dt.date.fromisoformat(iso)


def period_of(memo: str) -> tuple:
    """('2026-07-02', '2026-08-01') from a memo, or (None, None)."""
    m = _PERIOD.search(memo or "")
    if not m:
        return (None, None)
    g = m.groups()
    return (_iso(g[0], g[1], g[2]), _iso(g[3], g[4], g[5]))


def memo_with(memo: str, start: str, end: str) -> str:
    """The memo with its period set to start - end: the dates swapped in place when it has one (every other
    word of the memo untouched), else "(Period:...)" added on its own line."""
    memo = memo or ""
    new = f"Period:{_us(start)} - {_us(end)}"
    if _PERIOD.search(memo):
        return _PERIOD.sub(new, memo, count=1)
    return (memo.rstrip() + "\n" if memo.strip() else "") + f"({new})"


def check(periods: list) -> str:
    """'' when every period is a real window and no two draws overlap; else the reason, in plain words.
    Several invoices on one draw share one period - identical windows are one draw, not an overlap."""
    seen = {}
    for p in periods:
        s, e = str(p.get("start") or ""), str(p.get("end") or "")
        try:
            ds, de = _day(s), _day(e)
        except ValueError:
            return f"{p.get('label') or p.get('id')}: a start and an end date are needed"
        if de < ds:
            return f"{p.get('label') or p.get('id')}: the end ({_us(e)}) is before the start ({_us(s)})"
        if (de - ds).days > MAX_SPAN_DAYS:
            return f"{p.get('label') or p.get('id')}: {_us(s)} - {_us(e)} is longer than {MAX_SPAN_DAYS} days"
        seen.setdefault((s, e), p.get("label") or p.get("id"))
    wins = sorted(seen.items())
    for (a, b) in zip(wins, wins[1:]):
        (s1, e1), l1 = a
        (s2, e2), l2 = b
        if s2 <= e1:
            return f"{l2} ({_us(s2)} - {_us(e2)}) overlaps {l1} ({_us(s1)} - {_us(e1)})"
    return ""


_TABLE = ("CREATE TABLE IF NOT EXISTS draw_period (invoice_id TEXT PRIMARY KEY, project_no TEXT NOT NULL, "
          "start TEXT NOT NULL, end TEXT NOT NULL, saved_at TEXT NOT NULL)")


def saved(con, project: str) -> dict:
    """{invoice id: (start, end, saved_at)} - the windows set on the ledger for this job ({} when none / no table)."""
    try:
        return {str(r[0]): (r[1], r[2], r[3]) for r in
                con.execute("SELECT invoice_id, start, end, saved_at FROM draw_period WHERE project_no=?", ((project or "").upper(),))}
    except Exception:            # noqa: BLE001 - an older ledger without the table
        return {}


def save(con, project: str, periods: list) -> dict:
    """Replace the job's saved windows with the ones that differ from the invoice. Refused on an overlap or a
    window that is not one. Ledger only."""
    project = (project or "").strip().upper()
    rows = [p for p in (periods or []) if isinstance(p, dict) and str(p.get("id") or "").strip()]
    if not project or not rows:
        return {"ok": False, "error": "project and periods required"}
    why = check(rows)
    if why:
        return {"ok": False, "error": why}
    keep = [p for p in rows if (p.get("start"), p.get("end")) != (p.get("inv_start"), p.get("inv_end"))]
    now = dt.datetime.now().isoformat(timespec="seconds")
    con.execute(_TABLE)
    con.execute("DELETE FROM draw_period WHERE project_no=?", (project,))
    for p in keep:
        con.execute("INSERT OR REPLACE INTO draw_period(invoice_id, project_no, start, end, saved_at) VALUES (?,?,?,?,?)",
                    (str(p["id"]).strip(), project, p["start"], p["end"], now))
    con.commit()
    return {"ok": True, "saved": len(keep), "saved_at": now}


def main() -> None:
    pn = (sys.argv[1] if len(sys.argv) > 1 else "").strip().upper()
    if not pn:
        sys.exit(__doc__)
    con = qbo_mirror.connect()
    try:
        rows = []
        for rec in qbo_mirror.load("Invoice", con=con):
            if pn in {j.upper() for j in qbo_mirror.jobs_of(rec)} or pn in str((rec.get("CustomerRef") or {}).get("name") or "").upper():
                s, e = period_of(rec.get("PrivateNote") or "")
                rows.append((s or "", rec.get("TxnDate"), rec.get("DocNumber"), s, e, (rec.get("PrivateNote") or "").split("\n")[0][:60]))
    finally:
        con.close()
    for _, d, doc, s, e, head in sorted(rows):
        print(f"{doc:>8}  {d}  {(_us(s) + ' - ' + _us(e)) if s else '(no period)':<25}  {head}")


if __name__ == "__main__":
    main()
