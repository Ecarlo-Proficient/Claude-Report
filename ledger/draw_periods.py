#!/usr/bin/env python3
"""
draw_periods.py - Draw periods, AS AN EXAMPLE ONLY (the owner 10/11/2026: "a way to change period draw dates by moving
it in a horizontal <> way and it changes the dates as it goes for all, and then a way to edit individual periods
themselves without it overlapping the other ones" - and then: "just as an example, never write").

A draw's period is the "(Period:MM/DD/YYYY - MM/DD/YYYY)" on its invoice's memo in QuickBooks. The ledger's project
page lets the owner slide every period together or type one draw's dates between its neighbours and see the result
in the Draw overview - and that is all it does. NOTHING IS WRITTEN: not to QuickBooks, not to the ledger, not to
Notion. The rules below (read a period, rewrite a memo's period in place, refuse overlaps) are the shared vocabulary
should a real change ever be asked for; today only the page uses the overlap rule, through its own date limits.

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
