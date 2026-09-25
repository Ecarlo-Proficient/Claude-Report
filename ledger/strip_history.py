#!/usr/bin/env python3
"""
strip_history.py - the permanent record of every paid check QuickBooks left
with no bills on it ("stripped"), fixed or not, and the QBO support report
built from it. Beside check_drift.py: that page lists what is still WRONG;
this keeps what HAPPENED, so a fix never erases the evidence (the owner
2026-09-25: "don't just remove it once I fix it, it needs a history so that
we can pull that pdf at any time with the updated info").

Read-only on the raw QBO mirror (shared/qbo_mirror). No QBO call, no QBO write.

WHERE THE EVENTS COME FROM
  - the mirror's change log (mirror_change, kept from 09/23/2026 on, never
    pruned): every BillPayment change flagged "payment unapplied", with the
    check's before copy (which bills, how many saves). A void, or a check
    renumbered (a "To print" check re-issued as ACH), is a normal act - skipped.
  - the case file <companyhealth>/check_strip_case.json (local business data,
    never in the repo) for strips that happened BEFORE the change log existed:
        {"prelog": [{"payment_id": "1317436", "changed_at": "2026-09-17T07:52:13-07:00",
                     "bills": ["1317269"], "note": "bill confirmed in QBO audit history"}],
         "report": {"ruled_out": ["..."], "apps": "...", "asks": ["..."]}}
    `bills` is optional (the bill ids the check paid, when someone looked them
    up in QuickBooks' audit history). `report` holds the report's narrative -
    what was ruled out, the connected apps, what we ask QuickBooks.

PER EVENT: when (Central), check #, vendor, amount, bills before, saves, the
trigger (a paid bill deleted within a minute before / no bill edited or deleted
/ not recorded), status NOW (re-applied + when, or still unapplied), and
"paid again": a bill the check paid that a LATER payment now pays (a double
payment - check 48299 / 48379, 09/18/2026).

    python3 ledger/strip_history.py          the history, one line per event
    python3 ledger/strip_history.py --pdf    rebuild the QBO support PDF (CompanyHealth)
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from shared import paths                    # noqa: E402
from shared import qbo_mirror as mirror      # noqa: E402
import bill_payment_stub                     # noqa: E402  (same tool: its Chrome print recipe)

CASE_FILE = "check_strip_case.json"
PDF_NAME = "QBO Support - Unapplied Bill Payments.pdf"
BURST_GAP = dt.timedelta(minutes=30)       # strips closer than this are one window
TRIGGER_WINDOW = dt.timedelta(seconds=90)  # a paid bill deleted this close before the strip caused it
CENTRAL = dt.timezone(dt.timedelta(hours=-5))   # the owner's clock (CDT); QBO stamps are Pacific


def case_path() -> Path:
    return paths.companyhealth_dir() / CASE_FILE


def load_case() -> dict:
    p = case_path()
    try:
        return json.loads(p.read_text()) if p.exists() else {}
    except (OSError, ValueError):
        return {}


def _t(s) -> dt.datetime | None:
    try:
        return mirror.parse_iso(str(s))
    except (TypeError, ValueError):
        return None


def _ct(t: dt.datetime | None) -> str | None:
    return t.astimezone(CENTRAL).isoformat() if t else None


def _bill_links(pay: dict) -> list:
    out = []
    for ln in pay.get("Line") or []:
        for lt in ln.get("LinkedTxn") or []:
            if lt.get("TxnType") == "Bill":
                out.append((str(lt.get("TxnId")), round(float(ln.get("Amount") or 0), 2)))
    return out


def events(con=None) -> list:
    """Every strip on record, oldest first. JSON-safe."""
    own = con is None
    con = con or mirror.connect()
    try:
        return _events(con, load_case())
    finally:
        if own:
            con.close()


def _events(con, case: dict) -> list:
    chs = mirror.changes(con, limit=500000)
    bill_dels = [c for c in chs if c["entity"] == "Bill" and c["kind"] == "deleted"]
    pay_rows = [c for c in chs if c["entity"] == "BillPayment"]
    out = []
    for c in pay_rows:
        f = set(c["flags"])
        if "payment unapplied" not in f or f & {"voided", "number changed"}:
            continue
        t = _t(c["changed_at"])
        before = mirror.change_before(con, c["id"]) if c["has_before"] else None
        bills = [{"id": b, "amount": a} for b, a in _bill_links(before)] if before else []
        cause = [d for d in bill_dels if d.get("ref_id") == c.get("ref_id") and t and _t(d["changed_at"])
                 and dt.timedelta(0) <= t - _t(d["changed_at"]) <= TRIGGER_WINDOW]
        out.append({"payment_id": c["rec_id"], "at": t, "bills": bills, "bills_before": c.get("lines_before"),
                    "saves": (c.get("token_after") or 0) - (c.get("token_before") or 0) or None,
                    "trigger": (f"paid bill {cause[0]['doc_number'] or cause[0]['rec_id']} deleted" if cause
                                else "check saved, no bill edited or deleted"),
                    "recorded": True, "note": "", "time_known": True})
    # a deleted PAID bill strips the check that paid it (every line of it) even when the log holds no row for the
    # check itself (09/18: bill 1145 deleted -> check 48211 left empty) - the bill's before copy names the check
    near = lambda pid, t: any(e["payment_id"] == pid and abs(e["at"] - t) <= dt.timedelta(days=1) for e in out)  # noqa: E731
    for d in bill_dels:
        if "deleted paid bill" not in d["flags"] or not d["has_before"]:
            continue
        t, b = _t(d["changed_at"]), mirror.change_before(con, d["id"]) or {}
        for lt in b.get("LinkedTxn") or []:
            pid = str(lt.get("TxnId"))
            if "BillPayment" in str(lt.get("TxnType")) and t and not near(pid, t):
                out.append({"payment_id": pid, "at": t, "bills": [{"id": d["rec_id"], "amount": round(float(b.get("TotalAmt") or 0), 2)}],
                            "bills_before": None, "saves": None,
                            "trigger": f"paid bill {d['doc_number'] or d['rec_id']} deleted", "recorded": True, "note": "",
                            # a deletion found by an id sweep carries the sweep time, not QBO's: shown as "by" that time
                            "time_known": not str(d["changed_at"]).endswith("Z")})   # QBO's own stamps carry -07:00; a sweep's is UTC
    logged = {(e["payment_id"], e["at"].date()) for e in out if e["at"]}
    for p in case.get("prelog") or []:
        t = _t(p.get("changed_at"))
        if not t or (str(p.get("payment_id")), t.date()) in logged:
            continue
        out.append({"payment_id": str(p["payment_id"]), "at": t,
                    "bills": [{"id": str(b), "amount": None} for b in p.get("bills") or []],
                    "bills_before": len(p.get("bills") or []) or None, "saves": None,
                    "trigger": "not recorded (before the change log)", "recorded": False, "note": p.get("note") or "",
                    "time_known": True})

    for e in out:
        pay = mirror.get("BillPayment", e["payment_id"], con=con) or {}
        md = pay.get("MetaData") or {}
        e["check"] = pay.get("DocNumber") or e["payment_id"]
        e["vendor"] = (pay.get("VendorRef") or {}).get("name") or ""
        e["check_date"] = pay.get("TxnDate")
        e["total"] = round(float(pay.get("TotalAmt") or 0), 2)
        now = _bill_links(pay)
        e["bills_now"] = len(now)
        fixed = None
        if now:
            later = [c for c in pay_rows if c["rec_id"] == e["payment_id"] and _t(c["changed_at"])
                     and _t(c["changed_at"]) > e["at"] and (c.get("lines_after") or 0) > 0]
            fixed = min((_t(c["changed_at"]) for c in later), default=_t(md.get("LastUpdatedTime")))
        e["status"] = "re-applied" if now else "still unapplied"
        e["fixed_at"] = _ct(fixed)
        # paid again: a bill this check paid that a payment written AFTER the strip now pays
        again = []
        for b in e["bills"]:
            bill = mirror.get("Bill", b["id"], con=con) or {}
            b["doc_number"] = bill.get("DocNumber") or b["id"]
            b["txn_date"] = bill.get("TxnDate")
            if b["amount"] is None:
                b["amount"] = round(float(bill.get("TotalAmt") or 0), 2)
            for lt in bill.get("LinkedTxn") or []:
                if "BillPayment" not in str(lt.get("TxnType")) or str(lt.get("TxnId")) == e["payment_id"]:
                    continue
                other = mirror.get("BillPayment", str(lt.get("TxnId")), con=con) or {}
                made = _t((other.get("MetaData") or {}).get("CreateTime"))
                if made and made > e["at"]:
                    again.append({"bill": b["doc_number"], "bill_id": b["id"], "amount": b["amount"],
                                  "check": other.get("DocNumber") or other.get("Id"), "payment_id": other.get("Id"),
                                  "check_date": other.get("TxnDate"), "check_total": round(float(other.get("TotalAmt") or 0), 2)})
        e["paid_again"] = again
    out.sort(key=lambda e: e["at"])
    w, last = 0, None
    for e in out:
        if last is None or e["at"] - last > BURST_GAP:
            w += 1
        e["window"] = w
        last = e["at"]
    for e in out:
        e["at"] = _ct(e["at"])
    return out


def summary(ev: list) -> dict:
    wins = []
    for n in sorted({e["window"] for e in ev}):
        g = [e for e in ev if e["window"] == n]
        wins.append({"window": n, "start": g[0]["at"], "end": g[-1]["at"], "checks": len(g),
                     "total": round(sum(e["total"] for e in g), 2), "max_saves": max((e["saves"] or 0) for e in g),
                     "time_known": all(e["time_known"] for e in g)})
    return {"checks": len(ev), "total": round(sum(e["total"] for e in ev), 2),
            "by_delete": sum(e["trigger"].startswith("paid bill") for e in ev),
            "no_edit": sum(e["trigger"].startswith("check saved") for e in ev),
            "unrecorded": sum(not e["recorded"] for e in ev),
            "still": sum(e["status"] != "re-applied" for e in ev),
            "still_total": round(sum(e["total"] for e in ev if e["status"] != "re-applied"), 2),
            "fixed": sum(e["status"] == "re-applied" for e in ev),
            "paid_again": [dict(a, strip_check=e["check"]) for e in ev for a in e["paid_again"]],
            "windows": wins}


_DOUBLE_DDL = ("CREATE TABLE IF NOT EXISTS strip_double_payment (strip_payment_id TEXT NOT NULL, bill_id TEXT NOT NULL, "
               "later_payment_id TEXT NOT NULL, bill TEXT, amount REAL, later_check TEXT, later_check_date TEXT, "
               "first_seen TEXT NOT NULL, PRIMARY KEY (strip_payment_id, bill_id, later_payment_id))")


def _keep_doubles(ev: list, db_path) -> list:
    """A double payment is evidence: once seen it is kept in the ledger (strip_double_payment) and shown for good, even
    after the bills are put right in QuickBooks and the live check no longer shows it (48299 / 48379, 09/25/2026)."""
    if not db_path:
        return ev
    import sqlite3
    con = sqlite3.connect(str(db_path))
    try:
        con.execute(_DOUBLE_DDL)
        now = dt.datetime.now().isoformat(timespec="seconds")
        for e in ev:
            for a in e["paid_again"]:
                con.execute("INSERT OR IGNORE INTO strip_double_payment VALUES (?,?,?,?,?,?,?,?)",
                            (e["payment_id"], str(a["bill_id"]), str(a["payment_id"]), str(a["bill"]), a["amount"],
                             str(a["check"]), a["check_date"], now))
        con.commit()
        kept = {}
        for r in con.execute("SELECT strip_payment_id, bill_id, later_payment_id, bill, amount, later_check, later_check_date "
                             "FROM strip_double_payment"):
            kept.setdefault(r[0], []).append({"bill_id": r[1], "payment_id": r[2], "bill": r[3], "amount": r[4],
                                              "check": r[5], "check_date": r[6]})
    finally:
        con.close()
    for e in ev:
        have = {(str(a["bill_id"]), str(a["payment_id"])) for a in e["paid_again"]}
        e["paid_again"] += [k for k in kept.get(e["payment_id"], []) if (k["bill_id"], k["payment_id"]) not in have]
    return ev


def _apply_marks(ev: list, marks: dict | None) -> list:
    """A later check the owner kept as a credit settles the double payment it made (check_drift_mark in the ledger)."""
    for e in ev:
        for a in e["paid_again"]:
            m = (marks or {}).get(str(a.get("payment_id")))
            a["kept_as_credit"] = bool(m and m.get("kind") == "credit")
            a["credit_note"] = (m or {}).get("note") or ""
    return ev


def history(con=None, marks: dict | None = None, db_path=None) -> dict:
    """The ledger's /api/checkstrips payload."""
    try:
        ev = _apply_marks(_keep_doubles(events(con), db_path), marks)
    except Exception as ex:                  # noqa: BLE001 - a locked mirror / missing key is an answer
        return {"ok": False, "events": [], "error": f"strip history failed: {ex}"}
    return {"ok": True, "events": ev, "summary": summary(ev), "case_file": str(case_path()),
            "case_found": case_path().exists()}


# ── the report ──────────────────────────────────────────────────────────────
def _m(v) -> str:
    return f"${v:,.2f}"


def _d(iso) -> str:
    t = _t(iso)
    return t.astimezone(CENTRAL).strftime("%m/%d/%Y") if t else "–"


def _tm(iso) -> str:
    t = _t(iso)
    return t.astimezone(CENTRAL).strftime("%I:%M %p").lstrip("0") if t else "–"


def _ymd(s) -> str:
    try:
        return dt.date.fromisoformat(str(s)[:10]).strftime("%m/%d/%Y")
    except ValueError:
        return "–"


def report_html(ev: list, case: dict) -> str:
    s, rep, e_ = summary(ev), case.get("report") or {}, html.escape
    today = dt.date.today().strftime("%m/%d/%Y")
    wins = "".join(
        f"<li><b>{_d(w['start'])}, {'' if w['time_known'] else 'by '}{_tm(w['start'])}"
        + (f" to {_tm(w['end'])}" if _tm(w["end"]) != _tm(w["start"]) else "") + f":</b> "
        f"{w['checks']} check{'s' if w['checks'] != 1 else ''}, {_m(w['total'])}"
        + (f"; one check was saved <b>{w['max_saves']} times in a row</b>" if w["max_saves"] >= 3 else "") + ".</li>"
        for w in s["windows"])
    boxes = "".join(
        f"<div class=box><b>A double payment.</b> Check {e_(str(a['strip_check']))} paid bill {e_(str(a['bill']))} "
        f"({_m(a['amount'])}). After the check lost its bills, the bill showed as open and was paid again with check "
        f"{e_(str(a['check']))} dated {_ymd(a['check_date'])}."
        + (" We are keeping that second check as a credit with the vendor for their next bill." if a.get("kept_as_credit") else "")
        + "</div>" for a in s["paid_again"])
    boxes += (f"<div class=box><b>Books that no longer match the bank.</b> {s['still']} of these checks are still "
              f"unapplied today ({_m(s['still_total'])}): the bills they paid read as owed again.</div>")
    ul = lambda xs: "<ul>" + "".join(f"<li>{e_(x)}</li>" for x in xs) + "</ul>"   # noqa: E731
    rows = "".join(
        f"<tr><td class=n>{_d(e['at'])}</td><td class=n>{'' if e['time_known'] else 'by '}{_tm(e['at'])}</td><td>{e_(str(e['check']))}</td>"
        f"<td class=n>{_ymd(e['check_date'])}</td><td class='r n'>{_m(e['total'])}</td>"
        f"<td class=c>{e['bills_before'] if e['bills_before'] is not None else '–'}</td><td class=c>{e['saves'] or '–'}</td>"
        f"<td>{e_(e['trigger'][0].upper() + e['trigger'][1:])}</td>"
        f"<td class=n>{'Re-applied ' + _d(e['fixed_at']) if e['status'] == 're-applied' else 'Still unapplied'}"
        + "".join(f"<br><b>Paid again</b> ({a['check']})" + (" - kept as credit" if a.get("kept_as_credit") else "") for a in e["paid_again"])
        + "</td></tr>" for e in ev)
    return f"""<!doctype html><html><head><meta charset=utf-8><title>Unapplied bill payments</title><style>
@page {{ size: letter; margin: 0.6in; }}
body {{ font: 10.5pt -apple-system, Helvetica, Arial, sans-serif; color:#1f2430; }}
h1 {{ font-size: 17pt; margin:0 0 2px; }} h2 {{ font-size: 12pt; margin: 18px 0 6px; border-bottom:1px solid #d5d9df; padding-bottom:3px; }}
.sub {{ color:#5b6472; margin-bottom: 12px; }} p, li {{ line-height:1.4; }} ul {{ margin: 4px 0 4px 18px; padding:0; }}
table {{ border-collapse: collapse; width:100%; font-size: 8.5pt; }} th, td {{ border-bottom: 1px solid #e3e6ea; padding: 4px 5px; text-align:left; vertical-align:top; }}
th {{ font-size: 7.5pt; text-transform: uppercase; color:#5b6472; letter-spacing:.03em; }} .r {{ text-align:right; }} .c {{ text-align:center; }} .n {{ white-space: nowrap; }}
.box {{ border:1px solid #d5d9df; border-left: 3px solid #9b2c2c; border-radius:6px; padding:8px 12px; margin: 8px 0; }}
tr {{ page-break-inside: avoid; }}
</style></head><body>
<h1>Bill payments unapplied without anyone editing them</h1>
<div class=sub>QuickBooks Online · updated {today} · all times US Central</div>
<h2>What is happening</h2>
<p><b>{s['checks']} paid checks (bill payments) totalling {_m(s['total'])}</b> have been left in our company file with
all of their bill lines removed. Each check keeps its number, date, amount, bank account and print status, but is left
with $0.00 applied, and every bill it paid shows as open again.</p>
<ul><li><b>{s['no_edit']}</b> were re-saved with no bill edited or deleted first. Our Audit Log shows these under the company
owner's login, with no visible difference between entries. The owner did not make them.</li>
<li><b>{s['by_delete']}</b> followed the deletion of one paid bill: QuickBooks then removed <i>every</i> bill from the check
that paid it, not just the deleted one, with no warning.</li>
<li><b>{s['unrecorded']}</b> happened before we kept a copy of the data from before each change, so the trigger is not recorded.</li></ul>
<h2>When</h2><ul>{wins}</ul>
<h2>What it caused</h2>{boxes}
{('<h2>What we have already ruled out</h2>' + ul(rep['ruled_out'])) if rep.get('ruled_out') else ''}
{('<p>' + e_(rep['apps']) + '</p>') if rep.get('apps') else ''}
{('<h2>What we need from QuickBooks</h2>' + ul(rep['asks'])) if rep.get('asks') else ''}
<h2>Every affected check</h2>
<table><thead><tr><th>Date</th><th>Time</th><th>Check #</th><th>Check date</th><th class=r>Amount</th><th class=c>Bills before</th>
<th class=c>Saves</th><th>Trigger</th><th>Status today</th></tr></thead><tbody>{rows}</tbody></table>
<p class=sub style="margin-top:8px">"Bills before" and "Saves" come from a copy of our data taken before each change,
kept since 09/23/2026. "Saves" = how many times QuickBooks recorded the check being saved in that event.
{s['fixed']} of {s['checks']} checks have since been re-applied by us; {s['still']} are still unapplied as of {today}.</p>
</body></html>"""


def build_pdf(out: Path | None = None, con=None, marks: dict | None = None, db_path=None) -> dict:
    """Rebuild the report from the live history -> <companyhealth>/QBO Support - Unapplied Bill Payments.pdf."""
    out = out or paths.companyhealth_dir() / PDF_NAME
    ev = _apply_marks(_keep_doubles(events(con), db_path), marks)
    with tempfile.TemporaryDirectory() as tmp:
        h = Path(tmp) / "report.html"
        h.write_text(report_html(ev, load_case()))
        ok = bill_payment_stub.render_pdf(h, out)
    return {"ok": ok, "path": str(out), "checks": len(ev),
            **({} if ok else {"error": "Chrome not found or the print failed"})}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pdf", action="store_true", help="rebuild the QBO support PDF in CompanyHealth")
    a = ap.parse_args()
    if a.pdf:
        r = build_pdf()
        print(("wrote " + r["path"]) if r["ok"] else r["error"], f"({r['checks']} checks)")
        return 0 if r["ok"] else 1
    ev = events()
    for e in ev:
        print(f"{_d(e['at'])} {('' if e['time_known'] else 'by ') + _tm(e['at']):>11}  w{e['window']}  {str(e['check']):>9}  {_m(e['total']):>13}  "
              f"{e['status']:<15} {e['trigger']}" + (f"  PAID AGAIN by {[a['check'] for a in e['paid_again']]}" if e["paid_again"] else ""))
    s = summary(ev)
    print(f"{s['checks']} strips, {_m(s['total'])}; {s['fixed']} re-applied, {s['still']} still unapplied ({_m(s['still_total'])})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
