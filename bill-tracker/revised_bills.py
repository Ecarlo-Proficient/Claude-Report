#!/usr/bin/env python3
"""revised_bills.py - the REVISED-bill follow-up as a process (AP-13), in the Bill Tracker, 10/12/2026.

The owner: "did the clerk send the bill to get revised? when? have they responded? with the new
revised bill? is the new revised bill uploaded to QBO? that's the part where we can get in a pickle."
And: "keep the process in the Excel, forget about Notion - it's only the Excel."

What it does, every AP run (excel_bill_sync, inside the tracker's own save) and by hand:
  1. FIND every revised bill in the QuickBooks copy (shared/qbo_mirror): a bill whose memo carries
     the PM's `REVISED - <reason>` line (bill_rows.revised_reason - line 2, or line 1 for a bill
     entered under the QBO approval workflow). Nothing is read from QuickBooks directly.
  2. For each: the vendor's ORIGINAL total and WHEN the PM changed it (the mirror's change log:
     the last 'amount changed' edit), the open balance, the approval state, the project, and what
     the vendor's latest STATEMENT says about it (shared/statement_record - the reconciler's record:
     matched at the revised amount / still the old amount / not on the statement).
  3. FLAGS the pickles:  entered twice (another open bill from the same vendor with the same #, or
     the # plus a suffix, entered after the revision) · amount changed again after the clerk sent it
     · paid before the vendor re-issued · bill gone from QuickBooks.
  4. REBUILDS the `Revised Bills` sheet of Bill Tracker.xlsx (shared/revised_sheet, an Excel Table,
     one row per bill keyed by the QuickBooks bill id). The clerk's columns - Stage (a dropdown),
     Sent on · Replied on · Received on · Entered on, Notes - are read back from the sheet before
     the rebuild and put back as typed (like the Lien tags and Notes). Stage moves forward only,
     and by this tool only to "Done - statement agrees" or "Dropped".

By hand: `python3 revised_bills.py` previews (the mirror + the statement records, no workbook touched);
`--xlsx PATH` rebuilds that workbook's sheet (default: the live Bill Tracker; `assert_clean` last);
`--since YYYY-MM-DD` widens the scan (default: bills created or dated in the last 120 days - the
REVISED follow-up began 09/16/2026).
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import bill_approval, paths, qbo_api, qbo_mirror, revised_sheet as rs, statement_record as sr  # noqa: E402
from shared.xlsx_guard import put as xl_text  # noqa: E402  (outside text never becomes a formula)
sys.path.insert(0, str(Path(__file__).resolve().parent))
from bill_rows import revised_reason  # noqa: E402  (tool-local: the ONE reading of the REVISED memo line)

DEFAULT_DAYS = 120
SUFFIX_RE = re.compile(r"^(.+?)[\s\-_/]*(R|REV|REVISED|A|B|C|1|2|CC|CORR|CORRECTED)$", re.I)


def _ref_key(ref: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", (ref or "").upper())


def _iso_date(s: Optional[str]) -> str:
    return (s or "")[:10]


# ─────────────────────────── 1 + 2 + 3: find ───────────────────────────

def _memo_project(memo: str) -> str:
    m = qbo_api.PROJ_RE.search(memo or "")
    return m.group(1).upper() if m else ""


def _line_project(bill: dict) -> str:
    for ln in bill.get("Line") or []:
        det = ln.get("ItemBasedExpenseLineDetail") or ln.get("AccountBasedExpenseLineDetail") or {}
        name = ((det.get("CustomerRef") or {}).get("name")) or ""
        m = qbo_api.PROJ_RE.search(name)
        if m:
            return m.group(1).upper()
    return ""


def statement_says(vendor: str, ref: str, total: float, records: Optional[dict] = None) -> dict:
    """What the vendor's latest statement record shows for this bill #.
    -> {text, matched: bool|None, as_of}. matched=None when no record / not on the statement."""
    rec = None
    try:
        rec = (records or {}).get(sr.key(vendor)) if records is not None else sr.find(vendor)
    except Exception:                                    # noqa: BLE001 - share off = no answer, not a failure
        rec = None
    if not rec:
        return {"text": "no statement record for this vendor", "matched": None, "as_of": ""}
    as_of = sr.us_date(rec.get("as_of") or "")
    want = _ref_key(ref)
    hit = next((ln for ln in rec.get("lines") or [] if ln.get("ref") and _ref_key(ln["ref"]) == want), None)
    if not hit:
        return {"text": f"not on the {as_of} statement", "matched": None, "as_of": rec.get("as_of") or ""}
    amt = float(hit.get("amount") or 0)
    bucket = hit.get("bucket") or ""
    if bucket in ("matched", "approve", "checkqbo") and abs(amt - float(total or 0)) <= 0.01:
        return {"text": f"matched at ${amt:,.2f} on the {as_of} statement", "matched": True, "as_of": rec.get("as_of") or ""}
    if bucket == "lag":
        return {"text": f"paid in QBO, vendor still shows ${amt:,.2f} open on the {as_of} statement", "matched": False, "as_of": rec.get("as_of") or ""}
    return {"text": f"still ${amt:,.2f} on the {as_of} statement (QBO ${float(total or 0):,.2f})", "matched": False, "as_of": rec.get("as_of") or ""}


def find(con, since: Optional[dt.date] = None, today: Optional[dt.date] = None, records: Optional[dict] = None) -> List[dict]:
    """Every revised bill in the mirror, enriched. `records` = {statement_record.key(vendor): record}
    for tests; None reads the share per vendor."""
    today = today or dt.date.today()
    since = since or (today - dt.timedelta(days=DEFAULT_DAYS))
    s = since.isoformat()
    bills = {}
    by_vendor_refs: Dict[str, List[dict]] = {}
    for r in con.execute("SELECT id, txn_date, doc_number, total, created, last_updated, deleted, json FROM qbo_bill "
                         "WHERE (txn_date >= ? OR created >= ?)", (s, s)):
        try:
            j = qbo_mirror._unpack(r[7]) or {}
        except Exception:                                # noqa: BLE001
            continue
        vendor = ((j.get("VendorRef") or {}).get("name")) or ""
        rec = {"bill_id": str(r[0]), "txn_date": r[1] or "", "ref": r[2] or "", "total": float(r[3] or 0),
               "created": r[4] or "", "last_updated": r[5] or "", "deleted": bool(r[6]), "vendor": vendor,
               "memo": j.get("PrivateNote") or "", "balance": float(j.get("Balance") or 0), "json": j}
        by_vendor_refs.setdefault(sr.key(vendor), []).append(rec)
        created_d = _iso_date(rec["created"])
        reason = revised_reason(rec["memo"], allow_first_line=bool(created_d and created_d >= bill_approval.APPROVAL_WORKFLOW_START.isoformat()))
        if reason:
            bills[rec["bill_id"]] = dict(rec, reason=reason)
    # the change log: the vendor's original total + when the PM changed it (the LAST amount change)
    changes: Dict[str, List[dict]] = {}
    if bills:
        ids = list(bills)
        for i in range(0, len(ids), 500):
            chunk = ids[i:i + 500]
            for c in con.execute(f"SELECT rec_id, kind, changed_at, total_before, total_after, flags FROM mirror_change "
                                 f"WHERE entity='Bill' AND rec_id IN ({','.join('?' * len(chunk))}) ORDER BY changed_at", chunk):
                changes.setdefault(str(c[0]), []).append({"kind": c[1], "at": c[2] or "", "before": c[3], "after": c[4],
                                                          "flags": [f for f in (c[5] or "").split(",") if f]})
    out = []
    if records is None:                                  # one read of the share per vendor, not per bill
        records = {}
        for v in sorted({b["vendor"] for b in bills.values()}):
            try:
                rec = sr.find(v)
            except Exception:                            # noqa: BLE001
                rec = None
            if rec:
                records[sr.key(v)] = rec
    for bid, b in bills.items():
        ch = changes.get(bid, [])
        amount_edits = [c for c in ch if c["kind"] == "edited" and "amount changed" in c["flags"]]
        first_total = next((c["after"] for c in ch if c["kind"] == "created"), None)
        original = amount_edits[0]["before"] if amount_edits else first_total
        revised_at = amount_edits[-1]["at"] if amount_edits else (b["last_updated"] or b["created"])
        approval = bill_approval.approval_state(b["memo"], _iso_date(b["created"]), b["balance"])
        project = _memo_project(b["memo"]) or _line_project(b["json"])
        st = statement_says(b["vendor"], b["ref"], b["total"], records)
        flags: List[str] = []
        if b["deleted"]:
            flags.append("bill gone from QuickBooks")
        # entered twice: another live bill from the same vendor, same # (or # + suffix), entered after the revision
        want = _ref_key(b["ref"])
        base = SUFFIX_RE.match(b["ref"] or "")
        for o in by_vendor_refs.get(sr.key(b["vendor"]), []):
            if o["bill_id"] == bid or o["deleted"]:
                continue
            ok = _ref_key(o["ref"])
            m2 = SUFFIX_RE.match(o["ref"] or "")
            same = want and (ok == want or (m2 and _ref_key(m2.group(1)) == want) or (base and _ref_key(base.group(1)) == ok))
            if same and _iso_date(o["created"]) >= _iso_date(revised_at):
                flags.append(f"entered twice? bill {o['ref']} ({o['txn_date']}, ${o['total']:,.2f}) entered {_iso_date(o['created'])}")
        if b["balance"] <= 0.005 and not b["deleted"] and st["matched"] is not True:
            flags.append("paid before the vendor re-issued")
        out.append({"bill_id": bid, "vendor": b["vendor"], "ref": b["ref"], "bill_date": b["txn_date"], "project": project,
                    "reason": b["reason"], "total": b["total"], "original": original, "open": b["balance"], "approval": approval,
                    "revised_at": _iso_date(revised_at), "revised_at_full": revised_at, "deleted": b["deleted"],
                    "amount_edits": [{"at": c["at"], "before": c["before"], "after": c["after"]} for c in amount_edits],
                    "statement_says": st["text"], "statement_matched": st["matched"], "statement_as_of": st["as_of"],
                    "flags": flags})
    out.sort(key=lambda r: (r["vendor"].upper(), r["bill_date"]))
    return out


# ─────────────────────────── 4: the sheet ───────────────────────────

def prior_rows(path: Path) -> List[dict]:
    """The sheet as it stands (the clerk's Stage / dates / Notes live here)."""
    return rs.read_sheet(Path(path))


def plan(found: List[dict], prior: List[dict], today: Optional[dt.date] = None) -> List[dict]:
    """Pure: the rows to write, given the sheet as it stands. The clerk's columns come back as
    typed; Stage: forward only - the tool sets "To send" on a blank, "Done - statement agrees"
    when the statement agrees, "Dropped" when the bill is gone; a stage the clerk set stays.
    'amount changed again after sent' is judged here (needs the clerk's date)."""
    by_key = {p.get("key"): p for p in prior if p.get("key")}
    out = []
    for r in found:
        p = by_key.get(r["bill_id"]) or {}
        typed = {rs.col_key(c): (p.get(rs.col_key(c)) or "") for c in rs.TYPED}
        flags = list(r["flags"])
        sent = typed.get("sent_on") or ""
        if sent and any(_iso_date(e["at"]) > sent for e in r["amount_edits"]):
            flags.append(f"amount changed again in QuickBooks after the clerk sent it ({sent[5:7]}/{sent[8:10]})")
        stage_now = rs.stage_of(typed.get("stage"))
        stage = stage_now or "To send"
        if r["deleted"] and stage_now not in rs.FINAL:
            stage = "Dropped"
        elif r["statement_matched"] is True and stage_now not in rs.FINAL:
            stage = "Done - statement agrees"
        if stage_now and rs.STAGE_RANK.get(stage, 0) < rs.STAGE_RANK.get(stage_now, 0):
            stage = stage_now                            # never backwards
        row = dict(r, flags=flags, flags_text=rs.flags_text(flags), stage=stage, **{k: v for k, v in typed.items() if k != "stage"})
        row["standing"] = rs.standing({**{k: v for k, v in typed.items() if k != "stage"}, "stage": stage, "revised_in_qbo": r["revised_at"]}, today)
        row["moved"] = stage != stage_now
        out.append(row)
    # a bill no longer found (older than the scan, or purged) keeps its row only while the clerk is mid-process
    seen = {r["bill_id"] for r in found}
    for p in prior:
        if p.get("key") in seen or not p.get("key"):
            continue
        st = rs.stage_of(p.get("stage")) or "To send"
        if st in rs.FINAL:
            continue
        kept = {"bill_id": p["key"], "vendor": p.get("vendor", ""), "ref": p.get("bill", ""), "bill_date": p.get("bill_date", ""),
                "project": p.get("project", ""), "reason": p.get("pm_reason", ""), "total": p.get("qbo_now"), "original": p.get("vendor_billed"),
                "open": p.get("open"), "approval": p.get("approval", ""), "revised_at": p.get("revised_in_qbo", ""),
                "statement_says": p.get("statement_says", ""), "flags": ["not in the last 120 days of the QuickBooks copy - carried as typed"],
                "stage": st, "moved": False, **{rs.col_key(c): p.get(rs.col_key(c), "") for c in rs.TYPED if c != "Stage"}}
        kept["flags_text"] = rs.flags_text(kept["flags"])
        kept["standing"] = rs.standing({**{rs.col_key(c): p.get(rs.col_key(c), "") for c in rs.TYPED if c != "Stage"}, "stage": st, "revised_in_qbo": kept["revised_at"]}, today)
        out.append(kept)
    out.sort(key=lambda r: (r["standing"]["final"], r["standing"]["rank"], -(r["standing"]["days"] or 0), r["vendor"].upper(), r["bill_date"] or ""))
    return out


def _date(v):
    try:
        return dt.date.fromisoformat(str(v)[:10]) if v else None
    except ValueError:
        return None


def write_sheet(wb, rows: List[dict]) -> int:
    """Replace the workbook's `Revised Bills` sheet with these rows (an Excel Table, the Stage dropdown)."""
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo
    if rs.SHEET in wb.sheetnames:
        idx = wb.sheetnames.index(rs.SHEET)
        del wb[rs.SHEET]
    else:
        idx = None
    ws = wb.create_sheet(rs.SHEET, idx)
    ws.sheet_view.showGridLines = False
    for c_i, name in enumerate(rs.HEADERS, 1):
        hc = ws.cell(row=1, column=c_i, value=name)
        hc.font = Font(name="Calibri", size=11, bold=True)
        hc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30
    for c_i, w in enumerate(rs.WIDTHS, 1):
        ws.column_dimensions[get_column_letter(c_i)].width = w
    values = {
        "Vendor": lambda r: r.get("vendor", ""), "Bill #": lambda r: r.get("ref", ""), "Bill date": lambda r: _date(r.get("bill_date")),
        "Project": lambda r: r.get("project", ""), "PM reason": lambda r: r.get("reason", ""),
        "Vendor billed": lambda r: r.get("original"), "QBO now": lambda r: r.get("total"), "Open": lambda r: r.get("open"),
        "Approval": lambda r: r.get("approval", ""), "Revised in QBO": lambda r: _date(r.get("revised_at")),
        "Statement says": lambda r: r.get("statement_says", ""), "Flags": lambda r: r.get("flags_text", ""),
        "Where it stands": lambda r: (r.get("standing") or {}).get("text", ""),
        "Stage": lambda r: r.get("stage", ""), "Sent on": lambda r: _date(r.get("sent_on")), "Replied on": lambda r: _date(r.get("replied_on")),
        "Received on": lambda r: _date(r.get("received_on")), "Entered on": lambda r: _date(r.get("entered_on")),
        "Notes": lambda r: r.get("notes", ""), "QBO": lambda r: rs.QBO_BILL_URL.format(bill_id=r["bill_id"]) if r.get("bill_id") else "",
        "_Key": lambda r: r.get("bill_id", ""),
    }
    r_i = 2
    for row in rows:
        for c_i, (h, kind) in enumerate(zip(rs.HEADERS, rs.KINDS), 1):
            v = values[h](row)
            if kind == "date":
                c = ws.cell(row=r_i, column=c_i, value=v)
                c.number_format = "mm/dd/yyyy"
            elif kind == "money":
                c = ws.cell(row=r_i, column=c_i, value=round(float(v), 2) if v not in (None, "") else None)
                c.number_format = '"$"#,##0.00'
            elif kind == "url":
                c = ws.cell(row=r_i, column=c_i, value=f'=HYPERLINK("{v}","↗")' if v else None)
                c.font = Font(name="Calibri", size=11, color="0563C1", underline="single")
                c.alignment = Alignment(horizontal="center")
            else:
                c = xl_text(ws, r_i, c_i, v if v is not None else "")
                if h in ("PM reason", "Statement says", "Flags", "Where it stands", "Notes"):
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        r_i += 1
    last = max(r_i - 1, 2)
    if not rows:
        ws.cell(row=2, column=1, value="(no revised bills in the last 120 days)")
    tbl = Table(displayName=rs.TABLE, ref=f"A1:{get_column_letter(len(rs.HEADERS))}{last}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
                                        showRowStripes=True, showColumnStripes=False)
    ws.add_table(tbl)
    # the Stage dropdown: an inline list (no reference operators, under 255 characters - the rules Excel keeps),
    # error alert off so the clerk may still type her own words
    dv = DataValidation(type="list", formula1=f'"{rs.STAGE_LIST}"', allow_blank=True, showErrorMessage=False)
    col = get_column_letter(rs.HEADERS.index("Stage") + 1)
    dv.add(f"{col}2:{col}{max(last, 200)}")
    ws.add_data_validation(dv)
    return len(rows)


def add_sheet(wb, prior: List[dict], today: Optional[dt.date] = None, since: Optional[dt.date] = None, log=print) -> int:
    """The AP run's hook: find in the mirror, plan against the prior rows, write the sheet."""
    con = qbo_mirror.connect()
    try:
        found = find(con, since=since, today=today)
    finally:
        con.close()
    rows = plan(found, prior, today)
    moved = [r for r in rows if r.get("moved")]
    if moved:
        log("  Revised Bills: stage moved on " + ", ".join(f"{r['vendor']} {r['ref']} -> {r['stage']}" for r in moved[:8])
            + (" …" if len(moved) > 8 else ""))
    return write_sheet(wb, rows)


def keep_sheet(wb, prior: List[dict]) -> int:
    """No fresh answer this run: put the sheet back exactly as read, so nothing typed is lost."""
    rows = plan([], prior) if prior else []
    return write_sheet(wb, rows)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--xlsx", nargs="?", const="LIVE", default=None, metavar="PATH",
                    help="rebuild the Revised Bills sheet of this workbook (no PATH = the live Bill Tracker); default: preview only")
    ap.add_argument("--since", default="", metavar="YYYY-MM-DD", help=f"scan bills created or dated since (default: {DEFAULT_DAYS} days back)")
    ap.add_argument("--json", action="store_true", help="print the found bills as JSON")
    args = ap.parse_args()
    since = dt.date.fromisoformat(args.since) if args.since else None
    target = Path(paths.bill_tracker_xlsx()) if args.xlsx == "LIVE" else (Path(args.xlsx) if args.xlsx else None)
    prior = prior_rows(target) if target else []
    con = qbo_mirror.connect()
    try:
        found = find(con, since=since)
    finally:
        con.close()
    rows = plan(found, prior)
    if args.json:
        print(json.dumps(rows, indent=1, default=str))
    print(f"  revised bills: {len(found)} in the QuickBooks copy · sheet had {len(prior)} row(s)" + ("" if target else " (preview only)"))
    for r in rows:
        print(f"    {r['vendor'][:28]:28s} {r['ref']:14s} ${float(r['total'] or 0):>11,.2f}"
              + (f" (was ${r['original']:,.2f})" if r.get("original") not in (None, r["total"]) else "")
              + f"  {r['stage']}{' (moved)' if r.get('moved') else ''} · {r['standing']['text']} · {r['statement_says']}"
              + (f"  !! {r['flags_text']}" if r["flags_text"] else ""))
    if not target:
        return 0
    from openpyxl import load_workbook
    from shared import xlsx_verify
    if not target.exists():
        print(f"  no workbook at {target}")
        return 1
    wb = load_workbook(target)
    write_sheet(wb, rows)
    wb.save(target)
    xlsx_verify.assert_clean(target)
    print(f"  wrote the {rs.SHEET} sheet of {target}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
