#!/usr/bin/env python3
"""
check_drift.py - checks QBO quietly rewrote after they were paid. The ledger's
"Checks QBO changed" audit (Company, beside QBO changes) - /api/checkdrift.

Read-only on the raw QBO mirror (shared/qbo_mirror). No QBO call, no QBO write.

When a PAID bill is edited or deleted, QBO rewrites the check that paid it on
its own - no prompt, no warning, the audit log says "Indirect edit by System":
  - the check goes (partly or fully) UNAPPLIED - money that left the bank now
    floats on the vendor as a credit, and every bill it paid reopens;
  - the re-entered copy of the bill shows OPEN - it looks owed a second time;
  - the loan-repayment / vendor credit that rode on the check is freed;
  - later, QBO AUTO-APPLIES the floating money to the next bill entered for
    that vendor - a check dated 08/28 ends up "paying" a bill from 09/15, so
    the next real check goes out short by that amount.
Subs get hit hardest (weekly bills + loan credits on every check). The worked
example is check 48314 (the owner 2026-09-24): paid UC41 on 08/28, UC41 was
deleted 09/17 and re-entered 09/18, the check floated $4,635.70 and QBO then
put $135.40 of it on UC044 (09/15).

For every bill payment since --since this flags:
  FLOATING   check total > what it applies to bills (net of credits)
  LATE LINK  check applied to a bill dated AND entered after the check was
             written - QBO put floating money on the next bill
  COPY OPEN  an open bill for the same vendor entered after the check but
             dated on/before it (the re-entered copy of a bill the check paid)
  REOPENED   the original bill, EDITED after it was paid, now open again - QBO
             took the check off it on the edit (check 47436: UC015 edited
             09/22 7:23 am, the check stripped at 7:26). Matched to the penny,
             or any open bill in the 60 days before a sub's check
  CREDIT     an open vendor credit dated on/before the check (freed from it)
and names the cause when the mirror still holds it (a deleted bill that linked
to the check; the change log's pre-strip copy of the check). The copy match is
tight for subs (one bill a week), loose for high-volume suppliers.

    python3 ledger/check_drift.py                 summary, every year on file
    python3 ledger/check_drift.py --check 48314   one check, printed in full
    python3 ledger/check_drift.py --subs          subs only

Refresh the mirror first (python3 ledger/refresh_mirror.py) so it reads today.
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_mirror as mirror    # noqa: E402
from shared.qbo_costs import sub_evidence  # noqa: E402

EPS = 0.005
MIN_FLOAT = 1.00           # cents of rounding on old checks are not drift
SINCE = ""                 # every year on file - the audit runs all year round (owner 2026-09-24)
IMPORT_DONE = dt.datetime(2025, 6, 1, tzinfo=dt.timezone.utc)   # the 05/2025 import re-stamped CreateTime: entry order before it is unknowable
REOPEN_DAYS = 60           # a bill the check paid is dated within this many days before the check


def _t(s) -> dt.datetime | None:
    try:
        return mirror.parse_iso(str(s))
    except (TypeError, ValueError):
        return None


def _d(s) -> dt.date | None:
    try:
        return dt.date.fromisoformat(str(s)[:10])
    except (TypeError, ValueError):
        return None


def _fmt(d) -> str:
    return d.strftime("%m/%d/%Y") if d else ""


def _money(x) -> str:
    return f"${x:,.2f}"


def _created(rec: dict) -> dt.datetime | None:
    return _t((rec.get("MetaData") or {}).get("CreateTime"))


def _vendor(rec: dict) -> tuple:
    r = rec.get("VendorRef") or {}
    return str(r.get("value") or ""), str(r.get("name") or "")


def _links(pay: dict) -> list:
    """[(TxnType, TxnId, amount)] - what the payment applies, line by line."""
    out = []
    for ln in pay.get("Line") or []:
        for lt in ln.get("LinkedTxn") or []:
            out.append((lt.get("TxnType"), str(lt.get("TxnId")), round(float(ln.get("Amount") or 0), 2)))
    return out


def _bill_ref(b: dict) -> dict:
    return {"id": b.get("Id"), "doc_number": b.get("DocNumber") or "", "txn_date": b.get("TxnDate"),
            "entered": (b.get("MetaData") or {}).get("CreateTime"),
            "total": float(b.get("TotalAmt") or 0), "balance": float(b.get("Balance") or 0)}


def audit(con=None, since: str = SINCE, check: str | None = None, subs_only: bool = False) -> dict:
    """Every check QBO rewrote after payment, subs first. JSON-safe (the ledger serves it as is)."""
    own = con is None
    con = con or mirror.connect()
    try:
        return _audit(con, since, check, subs_only)
    finally:
        if own:
            con.close()


def _audit(con, since: str, check: str | None, subs_only: bool) -> dict:
    bills = {r["Id"]: r for r in mirror.load("Bill", include_deleted=True, con=con)}
    deleted_bills = {r[0] for r in con.execute("SELECT id FROM qbo_bill WHERE deleted=1")}
    credits = mirror.load("VendorCredit", con=con)
    pays = [p for p in mirror.load("BillPayment", con=con) if str(p.get("TxnDate") or "") >= since]

    sub_vendor = set()
    open_bills = defaultdict(list)
    for b in bills.values():
        vid = _vendor(b)[0]
        if sub_evidence(b.get("PrivateNote")):
            sub_vendor.add(vid)
        if b["Id"] not in deleted_bills and float(b.get("Balance") or 0) > EPS:
            open_bills[vid].append(b)
    open_credits = defaultdict(list)
    for c in credits:
        if float(c.get("Balance") or 0) > EPS:
            open_credits[_vendor(c)[0]].append(c)
    # deleted bills that still name the payment that paid them -> the cause
    killed_by = defaultdict(list)
    for bid in deleted_bills:
        b = bills.get(bid) or {}
        for lt in b.get("LinkedTxn") or []:
            if "BillPayment" in str(lt.get("TxnType")):
                killed_by[str(lt.get("TxnId"))].append(b)
    # the change log's strip events (kept from 09/23 on with the pre-strip copy)
    strips = defaultdict(list)
    for c in mirror.changes(con, entity="BillPayment", limit=100000):
        if "payment unapplied" in c["flags"] or (c.get("lines_before") or 0) > (c.get("lines_after") or 0):
            strips[c["rec_id"]].append(c)
    payers = defaultdict(list)                     # bill id -> [(check #, payment id, amount)] across every payment
    for q in pays:
        for typ, bid, amt in _links(q):
            if typ == "Bill":
                payers[bid].append((q.get("DocNumber") or q["Id"], q["Id"], amt))

    rows = []
    for p in pays:
        vid, vname = _vendor(p)
        is_sub = vid in sub_vendor
        if subs_only and not is_sub:
            continue
        if check and str(p.get("DocNumber") or "") != check:
            continue
        total = round(float(p.get("TotalAmt") or 0), 2)
        lk = _links(p)
        billed = sum(x[2] for x in lk if x[0] in ("Bill", "JournalEntry"))   # a JE to AP pays like a bill
        credited = sum(x[2] for x in lk if x[0] == "VendorCredit")
        floating = round(total - (billed - credited), 2)
        p_made, p_date = _created(p), _d(p.get("TxnDate"))

        late = []
        card = p.get("PayType") == "CreditCard"      # a card-fee bill added after an Amex payment is habit, not drift
        for typ, bid, amt in lk:
            if typ != "Bill" or card or not p_made or p_made < IMPORT_DONE:
                continue
            b = bills.get(bid) or {}
            bm, bd = _created(b), _d(b.get("TxnDate"))
            # a check cannot pay for work billed after it was written: the bill is dated AFTER the
            # check and entered after it. A re-entered copy (dated on/before the check) re-applied
            # by hand is a repair, not drift; the 05/2025 import re-stamped CreateTime on both sides.
            if p_made and bm and bm > p_made + dt.timedelta(minutes=1) and bd and p_date and bd > p_date:
                late.append({**_bill_ref(b), "amount": amt,
                             "others": [{"check": c, "payment_id": i, "amount": a}
                                        for c, i, a in payers.get(bid, []) if i != p["Id"]]})
        cands, freed = [], []
        if floating > MIN_FLOAT or late:
            for c in open_credits.get(vid, []):
                cd = _d(c.get("TxnDate"))
                if cd and p_date and cd <= p_date and (p_date - cd).days <= 31:
                    freed.append({"id": c.get("Id"), "doc_number": c.get("DocNumber") or "", "txn_date": c.get("TxnDate"),
                                  "balance": float(c.get("Balance") or 0), "memo": (c.get("PrivateNote") or "")[:120]})
            freed_amt = sum(c["balance"] for c in freed)
            for b in open_bills.get(vid, []):
                bm, bd = _created(b), _d(b.get("TxnDate"))
                if not (p_made and bm and bd and p_date and bd <= p_date):
                    continue
                copy = bm > p_made                     # entered after the check: the re-entered copy
                if not copy and (p_date - bd).days > REOPEN_DAYS:
                    continue
                bal = float(b.get("Balance") or 0)
                exact = abs(bal - floating) < 0.01 or abs(bal - freed_amt - floating) < 0.01
                if copy or exact or is_sub:
                    ref = _bill_ref(b)
                    ref.update(exact=exact, kind="copy" if copy else "reopened",
                               edited=(b.get("MetaData") or {}).get("LastUpdatedTime"))
                    cands.append(ref)
        if floating <= MIN_FLOAT and not late and not check:
            continue

        cause = [f"bill {b.get('DocNumber') or b['Id']} deleted" for b in killed_by.get(p["Id"], [])]
        for s in strips.get(p["Id"], []):
            cause.append(f"stripped {_fmt(_d(s['changed_at']))} (before copy kept)")
        applied = []
        for typ, bid, amt in lk:
            r = (bills.get(bid) or {}) if typ == "Bill" else {}
            applied.append({"type": typ, "id": bid, "amount": amt, "doc_number": r.get("DocNumber") or "",
                            "txn_date": r.get("TxnDate")})
        md = p.get("MetaData") or {}
        rows.append({
            "payment_id": p["Id"], "vendor": vname, "vendor_id": vid, "sub": is_sub,
            "check": p.get("DocNumber") or "", "txn_date": p.get("TxnDate"), "pay_type": p.get("PayType") or "",
            "written": md.get("CreateTime"), "last_changed": md.get("LastUpdatedTime"),
            "total": total, "applied": round(billed - credited, 2), "floating": floating,
            "late": late, "late_amt": round(sum(L["amount"] for L in late), 2),
            "freed": freed, "cause": cause, "applied_lines": applied, "_cands": cands,
        })
    _assign(rows)
    for r in rows:
        _finish(r)
    rows.sort(key=lambda r: (not r["sub"], -(r["floating"] + r["late_amt"])))
    fl = [r for r in rows if r["floating"] > MIN_FLOAT]
    summary = {
        "checks": len(rows), "subs": sum(1 for r in rows if r["sub"]),
        "floating": round(sum(r["floating"] for r in fl), 2), "floating_n": len(fl),
        "floating_subs": round(sum(r["floating"] for r in fl if r["sub"]), 2),
        "late_n": sum(1 for r in rows if r["late"]), "late": round(sum(r["late_amt"] for r in rows), 2),
        "copies_n": sum(len(r["copies"]) for r in rows), "copies": round(sum(r["copies_amt"] for r in rows), 2),
        "reopened_n": sum(len(r["reopened"]) for r in rows), "reopened": round(sum(r["reopened_amt"] for r in rows), 2),
    }
    return {"ok": True, "since": since, "rows": rows, "summary": summary,
            "last_refresh": mirror._meta_get(con, "last_refresh")}


def _assign(rows: list) -> None:
    """Each open bill belongs to ONE check. Exact-amount matches claim their bill first (oldest check
    first, the nearest-dated bill wins); a check with an exact match keeps only it; a check without one
    keeps the unclaimed candidates as 'possible'. Without this a sub's weekly $2,200 bill landed under
    every $2,200 check (Escobar, 2026-09-24)."""
    claimed = set()
    for r in sorted(rows, key=lambda r: r["txn_date"] or ""):
        ex = sorted((c for c in r["_cands"] if c["exact"] and c["id"] not in claimed),
                    key=lambda c: c["txn_date"] or "", reverse=True)
        if ex:
            r["_pick"] = [ex[0]]
            claimed.add(ex[0]["id"])
    for r in rows:
        if "_pick" not in r:
            r["_pick"] = [c for c in r["_cands"] if c["id"] not in claimed]
    for r in rows:
        pick = r.pop("_pick")
        r.pop("_cands")
        r["copies"] = [c for c in pick if c["kind"] == "copy"]
        r["reopened"] = [c for c in pick if c["kind"] == "reopened"]
        r["copies_amt"] = round(sum(b["balance"] for b in r["copies"]), 2)
        r["reopened_amt"] = round(sum(b["balance"] for b in r["reopened"]), 2)


def _finish(r: dict) -> None:
    """The cause an edit left behind + the plain-words fix list."""
    p_upd = _t(r["last_changed"])
    for b in r["reopened"]:                            # the edit that stripped it: minutes apart
        be = _t(b["edited"])
        if be and p_upd and abs((p_upd - be).total_seconds()) <= 1800:
            r["cause"].append(f"bill {b['doc_number']} edited {_fmt(be)} {be.strftime('%I:%M %p').lstrip('0')}, "
                              f"check changed {p_upd.strftime('%I:%M %p').lstrip('0')}")
    todo = []
    maybe = lambda b: "" if b["exact"] else " - possible, check it is the one"   # noqa: E731
    if r["reopened"]:
        todo.append("Re-apply this check to " + ", ".join(
            f"{b['doc_number']} ({_fmt(_d(b['txn_date']))}, open {_money(b['balance'])}){maybe(b)}" for b in r["reopened"])
            + " - the bill was edited after it was paid; do NOT pay it again")
    if r["copies"]:
        todo.append("Re-apply this check to " + ", ".join(
            f"{b['doc_number']} ({_fmt(_d(b['txn_date']))}, open {_money(b['balance'])}){maybe(b)}" for b in r["copies"])
            + " - the re-entered copy; do NOT pay it again")
    if r["freed"]:
        todo.append("Put back credit " + ", ".join(f"{c['doc_number'] or c['id']} {_money(c['balance'])}" for c in r["freed"]))
    for L in r["late"]:
        todo.append(f"Take {_money(L['amount'])} off {L['doc_number']} (dated {_fmt(_d(L['txn_date']))}, after this "
                    f"check) - {L['doc_number']} then shows {_money(L['amount'])} open = what the vendor was short-paid")
    if r["floating"] > MIN_FLOAT and not (r["copies"] or r["reopened"] or r["late"]):
        todo.append("Find the bill(s) this check paid and re-apply it")
    r["todo"] = todo


def _story(r: dict) -> None:
    print(f"\ncheck {r['check']} · {r['vendor']}{' (sub)' if r['sub'] else ''} · dated {_fmt(_d(r['txn_date']))} "
          f"· written {_fmt(_t(r['written']))} · last touched {_fmt(_t(r['last_changed']))}")
    print(f"  total {_money(r['total'])} · applied now {_money(r['applied'])} · FLOATING {_money(r['floating'])}")
    for a in r["applied_lines"]:
        print(f"    applied {_money(a['amount']):>11} -> {a['type']} {a['doc_number'] or a['id']} dated {_fmt(_d(a['txn_date']))}")
    for L in r["late"]:
        others = ", ".join(f"check {o['check']} {_money(o['amount'])}" for o in L["others"]) or "nothing else"
        print(f"  LATE LINK: {L['doc_number']} dated {_fmt(_d(L['txn_date']))}, entered {_fmt(_t(L['entered']))} - QBO put "
              f"{_money(L['amount'])} of the floating money on it; the rest of it was paid by {others}")
    for b in r["copies"]:
        print(f"  COPY OPEN: {b['doc_number']} dated {_fmt(_d(b['txn_date']))} entered {_fmt(_t(b['entered']))} "
              f"open {_money(b['balance'])} - reads as owed again")
    for b in r["reopened"]:
        print(f"  REOPENED: {b['doc_number']} dated {_fmt(_d(b['txn_date']))} open {_money(b['balance'])} "
              f"edited {_fmt(_t(b['edited']))}{'' if b['exact'] else ' (not an exact match)'}")
    for c in r["freed"]:
        print(f"  CREDIT FREED: {c['doc_number'] or c['id']} {_money(c['balance'])} ({c['memo'][:60]})")
    if r["cause"]:
        print(f"  cause: {'; '.join(r['cause'])}")
    print("  to fix: " + ("; ".join(r["todo"]) or "-"))


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--since", default=SINCE, help=f"check date floor (default {SINCE})")
    ap.add_argument("--check", help="one check number - print its full story")
    ap.add_argument("--subs", action="store_true", help="subs only")
    a = ap.parse_args()
    out = audit(since=a.since, check=a.check, subs_only=a.subs)
    print(f"mirror as of {out['last_refresh'] or '?'} (refresh first for today)")
    if a.check:
        for r in out["rows"]:
            _story(r)
        if not out["rows"]:
            print(f"check {a.check}: not found since {a.since}")
        return 0
    s = out["summary"]
    print(f"\nchecks QBO rewrote after payment{' since ' + a.since if a.since else ', every year on file'}: {s['checks']} ({s['subs']} subs)")
    print(f"  floating (paid, applied to nothing): {_money(s['floating'])} on {s['floating_n']} checks "
          f"(subs {_money(s['floating_subs'])})")
    print(f"  auto-applied to a bill dated after the check: {s['late_n']} checks, {_money(s['late'])}")
    print(f"  re-entered copy sitting OPEN (double-pay risk): {s['copies_n']} bills, {_money(s['copies'])}")
    print(f"  paid bill edited and reopened (double-pay risk): {s['reopened_n']} bills, {_money(s['reopened'])}")
    for r in out["rows"][:25]:
        print(f"  {'Sub' if r['sub'] else '':<3} {r['check']:<8} {_fmt(_d(r['txn_date']))} {r['vendor'][:28]:<28} "
              f"total {_money(r['total']):>12}  floating {_money(r['floating']):>11}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
