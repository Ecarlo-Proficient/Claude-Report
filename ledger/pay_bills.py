#!/usr/bin/env python3
"""
pay_bills.py - pay the saved pay run in QuickBooks: one bill payment per vendor.

The ledger's Pay run view holds the check run (which bills, how much - `pay_mark`). This module
turns that saved run into real QBO BillPayments, the same thing QuickBooks' Pay Bills screen does:

    plan(opts)          dry run, read LIVE from QBO: per vendor, the bills and amounts that would go
                        on the payment, the vendor's open credits (applied only when ticked), and
                        anything skipped (paid elsewhere, deleted, balance smaller now). Writes nothing.
    commit(body)        write it - only what the plan still says, re-checked live bill by bill.
    queued()            checks sitting "To print" in QBO (any, not only the ledger's).
    assign_number(...)  after a queued check is printed: give it its check # (marks it printed).

How it's paid (owner 2026-09-28):
    print   check, print later  -> PrintStatus NeedToPrint, number "To print" (QBO's own marker)
    check   check already written -> the check # the owner types
    ach     ACH / wire          -> the reference the owner types, kept as typed
    card    a credit-card account was picked -> the reference the owner types (optional)
QBO has no payment-method field on a bill payment (only check vs card), so the reference IS the
record of how it went out. QBO cuts a reference at 21 characters (the owner hit it with
"25783-123456831813153", dash included) - REF_MAX, enforced here and in the page.

Vendor credits go on only when the owner ticked them. Every write carries a QBO `requestid`
derived from the dry run it came from, so a double click or a retried request can never post
the same payment twice. The live bills are saved to ~/Library/Logs/Proficient/pay-bills/ before
any write. After each payment: its bills leave the pay run, the payment and bills go into the
mirror, and - unless it waits to be printed - its stub is filed in the vendor's stub folder
(bill_payment_stub.print_stub). A "To print" check files its stub when its number is assigned.

    ledger/pay_bills.py --plan              the dry run for the saved pay run, on the terminal
    ledger/pay_bills.py --queued            checks waiting to be printed
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import sys
import time
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import requests  # noqa: E402

from shared import bill_marks, qbo_api, qbo_mirror as mirror  # noqa: E402

EPS = 0.005
REF_MAX = 21                      # QBO's DocNumber limit on a bill payment
TO_PRINT = "To print"             # what QBO itself writes on a queued check
METHODS = ("print", "check", "ach", "card")
APPROVAL_SINCE = "2026-09-16"     # bill approval moved into a QBO workflow the API cannot see
LOG_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "pay-bills"


def _r2(x) -> float:
    return round(float(x or 0) + 0.0, 2)


def _ref(v) -> str:
    return str((v or {}).get("value") or "")


# ───────────────────────── reads ─────────────────────────

def accounts() -> list:
    """Active bank + credit-card accounts to pay from, the most used (last 180 days) first."""
    con = mirror.connect()
    try:
        accts = [a for a in mirror.load("Account", con=con)
                 if a.get("AccountType") in ("Bank", "Credit Card") and a.get("Active", True)]
        since = (dt.date.today() - dt.timedelta(days=180)).isoformat()
        uses: dict = {}
        for b in mirror.load("BillPayment", "txn_date >= ?", (since,), con=con):
            aid = _ref((b.get("CheckPayment") or {}).get("BankAccountRef")) or \
                _ref((b.get("CreditCardPayment") or {}).get("CCAccountRef"))
            if aid:
                uses[aid] = uses.get(aid, 0) + 1
    finally:
        con.close()
    out = [{"id": str(a["Id"]), "name": a.get("FullyQualifiedName") or a.get("Name") or "",
            "type": "card" if a.get("AccountType") == "Credit Card" else "bank", "uses": uses.get(str(a["Id"]), 0)}
           for a in accts]
    out.sort(key=lambda a: (-a["uses"], a["name"].lower()))
    return out


def _live(access: str, cid: str, typ: str, where: str) -> list:
    q = f"SELECT * FROM {typ} WHERE {where} MAXRESULTS 1000"
    return qbo_api._api_get(f"/v3/company/{cid}/query", access, {"query": q}).get("QueryResponse", {}).get(typ, [])


def _live_ids(access: str, cid: str, typ: str, ids) -> dict:
    ids, out = sorted(set(ids)), {}
    for i in range(0, len(ids), 40):
        chunk = ",".join(f"'{x}'" for x in ids[i:i + 40])
        for r in _live(access, cid, typ, f"Id IN ({chunk})"):
            out[str(r["Id"])] = r
    return out


def _credits(access: str, cid: str, vendor_id: str) -> list:
    rows = _live(access, cid, "VendorCredit", f"VendorRef = '{vendor_id}'")
    return [{"id": str(r["Id"]), "doc": r.get("DocNumber") or "", "date": r.get("TxnDate"),
             "balance": _r2(r.get("Balance")), "memo": (r.get("PrivateNote") or "").split("\n")[0][:80],
             "sync": r.get("SyncToken")}
            for r in rows if _r2(r.get("Balance")) > EPS]


def _plan_live(access: str, cid: str) -> dict:
    """The saved pay run against QBO right now: {vendor_id: {...}} + skipped rows."""
    marks = bill_marks.read_pay_marks()
    bills = _live_ids(access, cid, "Bill", marks) if marks else {}
    vendors, skipped = {}, []
    for bid, m in marks.items():
        b = bills.get(bid)
        if not b:
            skipped.append({"bill_id": bid, "why": "not in QuickBooks any more (deleted?)"})
            continue
        bal = _r2(b.get("Balance"))
        want = bal if m.get("amount") is None else _r2(m["amount"])
        row = {"bill_id": bid, "doc": b.get("DocNumber") or "", "date": b.get("TxnDate"), "balance": bal,
               "amount": want, "sync": b.get("SyncToken"), "partial": want < bal - EPS,
               "memo": (b.get("PrivateNote") or "").split("\n")[0][:80],
               "check_approval": str((b.get("MetaData") or {}).get("CreateTime") or "")[:10] >= APPROVAL_SINCE}
        vid, vname = _ref(b.get("VendorRef")), (b.get("VendorRef") or {}).get("name") or ""
        if bal <= EPS:
            skipped.append({**row, "vendor": vname, "why": "already paid in QuickBooks"})
            continue
        if want <= EPS:
            skipped.append({**row, "vendor": vname, "why": "pay amount is 0"})
            continue
        if want > bal + EPS:
            skipped.append({**row, "vendor": vname, "why": f"open balance is {bal:,.2f} now, less than the {want:,.2f} on the run"})
            continue
        v = vendors.setdefault(vid, {"vendor_id": vid, "vendor": vname, "bills": []})
        v["bills"].append(row)
    for v in vendors.values():
        v["bills"].sort(key=lambda r: (str(r["date"] or ""), r["doc"]))
        v["total"] = _r2(sum(r["amount"] for r in v["bills"]))
    return {"vendors": vendors, "skipped": skipped}


def plan() -> dict:
    """Dry run for the saved pay run. Writes nothing. The run_token keys the write's idempotency."""
    access, cid = qbo_api.load_credentials()
    p = _plan_live(access, cid)
    for v in p["vendors"].values():
        v["credits"] = _credits(access, cid, v["vendor_id"])
    vs = sorted(p["vendors"].values(), key=lambda v: v["vendor"].lower())
    return {"ok": True, "run_token": uuid.uuid4().hex, "ref_max": REF_MAX, "accounts": accounts(),
            "today": dt.date.today().isoformat(), "vendors": vs, "skipped": p["skipped"],
            "total": _r2(sum(v["total"] for v in vs)), "bills": sum(len(v["bills"]) for v in vs)}


# ───────────────────────── writes ─────────────────────────

def _post(cid: str, body: dict, request_id: str) -> dict:
    """POST a BillPayment create/update. The requestid makes a resend return the SAME payment
    (QBO de-duplicates on it), so a timeout or 5xx is retried safely; a 4xx is QBO saying no."""
    last = ""
    for attempt in range(4):
        try:
            r = requests.post(f"{qbo_api.API_BASE}/v3/company/{cid}/billpayment",
                              headers={"Authorization": f"Bearer {qbo_api._CURRENT_ACCESS}",
                                       "Content-Type": "application/json", "Accept": "application/json"},
                              params={"minorversion": qbo_api.MINOR_VERSION, "requestid": request_id},
                              json=body, timeout=120)
        except requests.exceptions.RequestException as e:
            last = type(e).__name__
            time.sleep(2 ** attempt)
            continue
        if r.status_code == 200:
            return r.json()["BillPayment"]
        if r.status_code == 401 and attempt < 2:
            qbo_api.refresh_access()
            continue
        if r.status_code in (429, 500, 502, 503, 504):
            last = f"status {r.status_code}"
            time.sleep(2 ** attempt)
            continue
        try:
            e = r.json()["Fault"]["Error"][0]
            msg = f"{e.get('Message')}: {e.get('Detail')}"
        except Exception:                                   # noqa: BLE001
            msg = r.text[:400]
        raise RuntimeError(f"QuickBooks refused it ({r.status_code}): {msg}")
    raise RuntimeError(f"QuickBooks did not answer ({last}) - check QuickBooks before trying again")


def _to_mirror(entity: str, recs: list) -> None:
    """Hold what QBO just returned, so the ledger and the stub see it before the next refresh."""
    if not recs:
        return
    con = mirror.connect()
    try:
        mirror.upsert_many(con, entity, recs, mirror.iso_z(mirror.now_utc()), track=True)
        con.commit()
    finally:
        con.close()


def _stub(pid: str) -> dict:
    """File the stub in the vendor's folder. A failure never undoes the payment - it is reported."""
    try:
        import bill_payment_stub as stubs          # noqa: PLC0415  (ledger-local module)
        from shared import paths                   # noqa: PLC0415
        r = stubs.print_stub(pid, company=paths.get("ACB_COMPANY_NAME", ""))
        return {"filed": bool(r.get("ok")), "file": r.get("file"), "error": r.get("error")}
    except Exception as e:                         # noqa: BLE001
        return {"filed": False, "error": str(e)}


def clean_ref(method: str, ref) -> str:
    """The reference as it will be written, or ValueError with the reason."""
    ref = str(ref or "").strip()
    if method == "print":
        return TO_PRINT
    if len(ref) > REF_MAX:
        raise ValueError(f"reference '{ref}' is {len(ref)} characters - QuickBooks keeps {REF_MAX} at most")
    if method in ("check", "ach") and not ref:
        raise ValueError("a check # / ACH reference is required")
    return ref


def commit(body: dict) -> dict:
    """body = {run_token, account_id, date, method, vendors: [{vendor_id, ref, bills: [{bill_id, amount}],
    credits: [credit ids ticked]}]}. Each vendor's payment is re-checked live against QBO and written
    only if it still matches the saved pay run exactly; one vendor failing never blocks the others."""
    token = str(body.get("run_token") or "")
    method = body.get("method")
    acct_id = str(body.get("account_id") or "")
    try:
        date = dt.date.fromisoformat(str(body.get("date") or "")).isoformat()
    except ValueError:
        return {"ok": False, "error": "a payment date is required"}
    if len(token) < 16 or method not in METHODS or not acct_id:
        return {"ok": False, "error": "bad request - run the review again"}
    acct = next((a for a in accounts() if a["id"] == acct_id), None)
    if not acct:
        return {"ok": False, "error": "that account is not an active bank or card account"}
    if (acct["type"] == "card") != (method == "card"):
        return {"ok": False, "error": "a card account pays by card; a bank account pays by check or ACH"}
    asked = body.get("vendors") or []
    try:
        refs = {str(v["vendor_id"]): clean_ref(method, v.get("ref")) for v in asked}
    except (ValueError, KeyError) as e:
        return {"ok": False, "error": str(e)}

    access, cid = qbo_api.load_credentials()
    live = _plan_live(access, cid)["vendors"]
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    results = []
    for v in asked:
        vid = str(v.get("vendor_id") or "")
        res = {"vendor_id": vid, "vendor": (live.get(vid) or {}).get("vendor") or vid}
        results.append(res)
        lv = live.get(vid)
        want = {str(b["bill_id"]): _r2(b["amount"]) for b in v.get("bills") or []}
        have = {r["bill_id"]: r["amount"] for r in (lv or {}).get("bills", [])}
        if not want or want != have:
            res["error"] = "the bills or amounts changed since the review (paid elsewhere, or the run was edited) - review again"
            continue
        credits = []
        if v.get("credits"):
            open_c = {c["id"]: c for c in _credits(access, cid, vid)}
            missing = [c for c in v["credits"] if str(c) not in open_c]
            if missing:
                res["error"] = "a ticked credit is no longer open in QuickBooks - review again"
                continue
            credits = [open_c[str(c)] for c in v["credits"]]
        bills_amt = _r2(sum(want.values()))
        cred_amt = _r2(sum(c["balance"] for c in credits))
        if cred_amt > bills_amt - EPS:      # a credit bigger than the payment: use only what the bills absorb
            res["error"] = f"the ticked credits ({cred_amt:,.2f}) cover the whole {bills_amt:,.2f} - apply those in QuickBooks"
            continue
        lines = [{"Amount": a, "LinkedTxn": [{"TxnId": bid, "TxnType": "Bill"}]} for bid, a in want.items()]
        lines += [{"Amount": c["balance"], "LinkedTxn": [{"TxnId": c["id"], "TxnType": "VendorCredit"}]} for c in credits]
        pay = {"VendorRef": {"value": vid}, "TxnDate": date, "TotalAmt": _r2(bills_amt - cred_amt),
               "DocNumber": refs[vid], "Line": lines}
        if method == "card":
            pay["PayType"] = "CreditCard"
            pay["CreditCardPayment"] = {"CCAccountRef": {"value": acct_id}}
            if not refs[vid]:
                pay.pop("DocNumber")
        else:
            pay["PayType"] = "Check"
            pay["CheckPayment"] = {"BankAccountRef": {"value": acct_id},
                                   "PrintStatus": "NeedToPrint" if method == "print" else "NotSet"}
        rid = hashlib.sha1(f"{token}|{vid}".encode()).hexdigest()[:36]
        stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        (LOG_DIR / f"{stamp}-{vid}.json").write_text(json.dumps(
            {"request_id": rid, "payment": pay, "account": acct, "method": method, "credits": credits,
             "bills_live": (lv or {}).get("bills")}, indent=1, default=str))
        try:
            done = _post(cid, pay, rid)
        except RuntimeError as e:
            res["error"] = str(e)
            continue
        pid = str(done["Id"])
        res.update({"ok": True, "payment_id": pid, "ref": done.get("DocNumber") or "", "total": _r2(done.get("TotalAmt")),
                    "bills": len(want), "credits": len(credits), "to_print": method == "print"})
        bill_marks.set_pay_marks([{"bill_id": b, "selected": False} for b in want], dt.datetime.now().isoformat(timespec="seconds"))
        try:
            _to_mirror("BillPayment", [done])
            _to_mirror("Bill", list(_live_ids(access, cid, "Bill", want).values()))
            if credits:
                _to_mirror("VendorCredit", list(_live_ids(access, cid, "VendorCredit", [c["id"] for c in credits]).values()))
        except Exception as e:                     # noqa: BLE001  (the payment is in QBO; the next refresh catches up)
            res["mirror_error"] = str(e)
        res["stub"] = {"filed": False, "waiting": "files when the check # is assigned"} if method == "print" else _stub(pid)
    posted = [r for r in results if r.get("ok")]
    return {"ok": bool(posted), "posted": len(posted), "failed": len(results) - len(posted),
            "total": _r2(sum(r["total"] for r in posted)), "results": results,
            **({} if posted else {"error": "; ".join(f"{r['vendor']}: {r.get('error')}" for r in results) or "nothing to pay"})}


# ───────────────────────── checks waiting to be printed ─────────────────────────

def queued() -> list:
    """Every bill payment QBO holds as a check "To print" (the mirror), newest first."""
    con = mirror.connect()
    try:
        out = []
        for b in mirror.load("BillPayment", con=con):
            cp = b.get("CheckPayment") or {}
            if cp.get("PrintStatus") != "NeedToPrint":
                continue
            out.append({"payment_id": str(b["Id"]), "vendor": (b.get("VendorRef") or {}).get("name") or "",
                        "date": b.get("TxnDate"), "total": _r2(b.get("TotalAmt")), "sync": b.get("SyncToken"),
                        "account": (cp.get("BankAccountRef") or {}).get("name") or "",
                        "bills": sum(1 for ln in b.get("Line") or [] for lt in ln.get("LinkedTxn") or []
                                     if lt.get("TxnType") == "Bill")})
    finally:
        con.close()
    out.sort(key=lambda r: (str(r["date"] or ""), r["payment_id"]), reverse=True)
    return out


def assign_number(payment_id: str, number: str) -> dict:
    """A queued check was printed: give it its check # and mark it printed, then file its stub."""
    pid, number = str(payment_id or "").strip(), str(number or "").strip()
    if not pid.isdigit():
        return {"ok": False, "error": "bad payment"}
    try:
        number = clean_ref("check", number)
    except ValueError as e:
        return {"ok": False, "error": str(e)}
    access, cid = qbo_api.load_credentials()
    live = qbo_api._api_get(f"/v3/company/{cid}/billpayment/{pid}", access)["BillPayment"]
    cp = live.get("CheckPayment") or {}
    if cp.get("PrintStatus") != "NeedToPrint":
        return {"ok": False, "error": f"that check is not waiting to be printed any more (it shows #{live.get('DocNumber') or '-'})"}
    acct = _ref(cp.get("BankAccountRef"))
    con = mirror.connect()
    try:
        dup = [b for b in mirror.load("BillPayment", "doc_number = ?", (number,), con=con)
               if str(b["Id"]) != pid and _ref((b.get("CheckPayment") or {}).get("BankAccountRef")) == acct]
    finally:
        con.close()
    if dup:
        d = dup[0]
        return {"ok": False, "error": f"check #{number} is already used on this account "
                                      f"({(d.get('VendorRef') or {}).get('name')}, {d.get('TxnDate')})"}
    body = {k: val for k, val in live.items() if k not in ("MetaData", "domain", "sparse")}
    body["DocNumber"] = number
    body["CheckPayment"] = {**cp, "PrintStatus": "PrintComplete"}
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    (LOG_DIR / f"{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}-number-{pid}.json").write_text(
        json.dumps({"live_before": live, "number": number}, indent=1, default=str))
    rid = hashlib.sha1(f"number|{pid}|{live.get('SyncToken')}|{number}".encode()).hexdigest()[:36]
    try:
        done = _post(cid, body, rid)
    except RuntimeError as e:
        return {"ok": False, "error": str(e)}
    try:
        _to_mirror("BillPayment", [done])
    except Exception as e:                         # noqa: BLE001
        return {"ok": True, "payment_id": pid, "number": number, "mirror_error": str(e), "stub": {"filed": False}}
    return {"ok": True, "payment_id": pid, "number": done.get("DocNumber"), "stub": _stub(pid)}


def _selftest() -> int:
    ok = True
    for m, r, want in (("print", "", TO_PRINT), ("ach", " 25783-123456831813153 ", "25783-123456831813153"),
                       ("check", "25801", "25801"), ("card", "", "")):
        got = clean_ref(m, r)
        ok &= got == want
        print(("ok  " if got == want else "FAIL"), m, repr(r), "->", repr(got))
    for m, r in (("ach", "25783-1234568318131534"), ("ach", ""), ("check", "")):
        try:
            clean_ref(m, r)
            print("FAIL", m, repr(r), "accepted")
            ok = False
        except ValueError as e:
            print("ok  ", m, repr(r), "refused:", e)
    return 0 if ok else 1


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--plan", action="store_true", help="dry run the saved pay run (live QBO read, no write)")
    ap.add_argument("--queued", action="store_true", help="checks waiting to be printed")
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()
    if a.selftest:
        return _selftest()
    if a.queued:
        for q in queued():
            print(f"{q['payment_id']:>10}  {q['date']}  {q['vendor'][:40]:<40} {q['total']:>12,.2f}  {q['bills']} bills  {q['account']}")
        return 0
    if a.plan:
        p = plan()
        for v in p["vendors"]:
            print(f"{v['vendor']}  {len(v['bills'])} bills  {v['total']:,.2f}  ({len(v['credits'])} open credits)")
            for b in v["bills"]:
                print(f"    {b['doc']:<16} {b['date']}  {b['amount']:>12,.2f}{'  partial' if b['partial'] else ''}"
                      f"{'  check approval in QBO' if b['check_approval'] else ''}")
        for s in p["skipped"]:
            print(f"  SKIP {s.get('vendor', '')} {s.get('doc', s['bill_id'])}: {s['why']}")
        print(f"total {p['total']:,.2f} on {p['bills']} bills - dry run, nothing written")
        return 0
    ap.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
