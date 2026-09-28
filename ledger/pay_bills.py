#!/usr/bin/env python3
"""
pay_bills.py - pay the saved pay run in QuickBooks: one bill payment per vendor.

The ledger's Pay run view holds the check run (which bills, how much - `pay_mark`). This module
turns that saved run into real QBO BillPayments, the same thing QuickBooks' Pay Bills screen does:

    plan(opts)          dry run, read LIVE from QBO: per vendor, the bills and amounts that would go
                        on the payment, the vendor's open credits (applied only when ticked), and
                        anything skipped (paid elsewhere, deleted, balance smaller now). Writes nothing.
    commit(body)        write it - only what the plan still says, re-checked live bill by bill.
    queue() / watch()   every payment pushed from here, until it MATCHES: QBO shows its number (a
                        "To print" check gets it when printed in QBO - nobody types it) and its stub
                        is filed. watch() runs after every mirror refresh; a payment QBO deleted,
                        voided or changed is flagged with what happened to its bills, and stays
                        until the owner marks it Resolved or Keeps it with a reason (mark()).
    assign_number(...)  fallback: a check printed OUTSIDE QBO gets its # typed here.

How it's paid (owner 2026-09-28):
    print   check, print later  -> PrintStatus NeedToPrint, number "To print" (QBO's own marker);
                                   print it in QBO, the next mirror refresh picks up the number
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
    ledger/pay_bills.py --queue             the pushed payments not matched yet, or flagged
    ledger/pay_bills.py --watch             judge them against the mirror, file matched stubs
"""
from __future__ import annotations

import argparse
import datetime as dt
import getpass
import hashlib
import json
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import requests  # noqa: E402

from shared import bill_marks, qbo_api, qbo_mirror as mirror  # noqa: E402

EPS = 0.005
REF_MAX = 21                      # QBO's DocNumber limit on a bill payment
TO_PRINT = "To print"             # what QBO itself writes on a queued check
METHODS = ("print", "check", "ach", "card")
APPROVAL_SINCE = "2026-09-16"     # bill approval moved into a QBO workflow the API cannot see
LOG_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "pay-bills"
TOKEN_TTL = 20 * 60               # a review is good for 20 minutes, once
FUTURE_DAYS = 30                  # a payment date may run this far ahead of today, no further
ACCOUNT_DAYS = 365                # only accounts that paid bills in this window are offered


# ───────────────────────── guardrails (owner 2026-09-28: "level 3 high access only") ─────────────────────────
# Ruled: ONLY the owner, ONLY this Mac, Touch ID on EVERY write, no dollar cap.
#  - authority: a register written once by `--authorize-this-mac` (itself behind Touch ID) holds a hash of this
#    Mac's hardware id + the macOS account. Absent or different -> every write is refused (fails closed), so a
#    clone on another machine can review but never pay.
#  - presence: macOS' own Touch ID / password dialog, naming the vendor count, total and bank account.
#  - one review, one write: the run_token is issued by plan() in THIS process, lives 20 minutes, is spent on use,
#    and the write must be a subset of exactly what that review showed.
#  - one writer at a time: a process lock + a file lock around every QBO write.
#  - every attempt, refused or written, is a line in ~/Library/Logs/Proficient/pay-bills/audit.log.
# What code cannot stop: someone with the code AND a QuickBooks connection of their own. QuickBooks' user roles
# and who holds an Intuit connection are the real boundary around the books.
_TOKENS: dict = {}
_WRITE_LOCK = threading.Lock()


def _machine_hash() -> str:
    out = subprocess.run(["/usr/sbin/ioreg", "-rd1", "-c", "IOPlatformExpertDevice"], capture_output=True, text=True).stdout
    uid = next((ln.split('"')[3] for ln in out.splitlines() if "IOPlatformUUID" in ln), "")
    if not uid:
        raise RuntimeError("could not read this Mac's hardware id")
    return hashlib.sha256(f"acb-pay|{uid}".encode()).hexdigest()


def _user_hash() -> str:
    return hashlib.sha256(f"acb-pay|{getpass.getuser()}".encode()).hexdigest()


def _auth_file() -> Path:
    from shared import paths                        # noqa: PLC0415
    return paths.register_file("pay_authority.json")


def authorized() -> tuple:
    """(True, "") when THIS Mac + macOS account is the one authorized to pay; (False, why) otherwise."""
    try:
        f = _auth_file()
        if not f.exists():
            return False, "payments from the ledger are locked on this Mac (not authorized)"
        a = json.loads(f.read_text())
        if a.get("machine") != _machine_hash() or a.get("user") != _user_hash():
            return False, "payments from the ledger are authorized on a different Mac or account"
        return True, ""
    except Exception as e:                          # noqa: BLE001
        return False, f"authority check failed: {e}"


def authorize_this_mac() -> int:
    import presence                                 # noqa: PLC0415  (ledger-local)
    ok, why = presence.confirm("authorize THIS Mac to pay bills in QuickBooks from the Project Ledger")
    if not ok:
        _audit("authorize", "refused", why)
        print(f"not authorized: {why}")
        return 1
    f = _auth_file()
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps({"machine": _machine_hash(), "user": _user_hash(), "authorized_at": _now_s()}, indent=1))
    f.chmod(0o600)
    _audit("authorize", "ok", "this Mac + account authorized (replaces any other)")
    print("this Mac + macOS account can now pay bills from the ledger (any other is no longer authorized)")
    return 0


def _now_s() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def _audit(action: str, result: str, detail="") -> None:
    """One line per attempt - never raises."""
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        with open(LOG_DIR / "audit.log", "a") as fh:
            fh.write(json.dumps({"at": _now_s(), "mac_user": getpass.getuser(), "action": action,
                                 "result": result, "detail": detail}, default=str) + "\n")
    except Exception:                               # noqa: BLE001
        pass


class _Locked:
    """Process lock + file lock: one QBO money write at a time, across tabs and processes."""
    def __enter__(self):
        import fcntl                                # noqa: PLC0415
        if not _WRITE_LOCK.acquire(blocking=False):
            raise BlockingIOError
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        self.fh = open(LOG_DIR / ".write.lock", "w")
        try:
            fcntl.flock(self.fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self.fh.close()
            _WRITE_LOCK.release()
            raise BlockingIOError from None
        return self

    def __exit__(self, *a):
        import fcntl                                # noqa: PLC0415
        fcntl.flock(self.fh, fcntl.LOCK_UN)
        self.fh.close()
        _WRITE_LOCK.release()


def _close_date(access: str, cid: str) -> str:
    """QBO's books-closed date ('' when none) - no payment is dated on or before it."""
    try:
        pref = qbo_api._api_get(f"/v3/company/{cid}/preferences", access).get("Preferences", {})
        return str((pref.get("AccountingInfoPrefs") or {}).get("BookCloseDate") or "")[:10]
    except Exception:                               # noqa: BLE001
        return ""


def _r2(x) -> float:
    return round(float(x or 0) + 0.0, 2)


def _ref(v) -> str:
    return str((v or {}).get("value") or "")


# ───────────────────────── reads ─────────────────────────

def accounts() -> list:
    """Active bank + card accounts that paid bills in the last year (ACCOUNT_DAYS), the most used first.
    An account never used to pay bills (money market, investments, clearing) is not offered."""
    con = mirror.connect()
    try:
        accts = [a for a in mirror.load("Account", con=con)
                 if a.get("AccountType") in ("Bank", "Credit Card") and a.get("Active", True)]
        since = (dt.date.today() - dt.timedelta(days=ACCOUNT_DAYS)).isoformat()
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
           for a in accts if uses.get(str(a["Id"]))]
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
        if "NOT APPROVED" in (b.get("PrivateNote") or "").upper():
            skipped.append({**row, "vendor": vname, "why": "its memo says NOT APPROVED - approve it in QuickBooks first"})
            continue
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
    can, why = authorized()
    token = uuid.uuid4().hex
    now = time.time()
    for t in [t for t, x in _TOKENS.items() if now - x["at"] > TOKEN_TTL]:
        _TOKENS.pop(t, None)
    _TOKENS[token] = {"at": now, "vendors": {v["vendor_id"]: {b["bill_id"]: b["amount"] for b in v["bills"]} for v in vs},
                      "credits": {v["vendor_id"]: {c["id"]: c["balance"] for c in v["credits"]} for v in vs},
                      "names": {v["vendor_id"]: v["vendor"] for v in vs}}
    today = dt.date.today()
    return {"ok": True, "run_token": token, "ref_max": REF_MAX, "accounts": accounts(), "can_pay": can, "why_not": why,
            "today": today.isoformat(), "min_date": _close_date(access, cid),
            "max_date": (today + dt.timedelta(days=FUTURE_DAYS)).isoformat(), "vendors": vs, "skipped": p["skipped"],
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


def _find_posted(access: str, cid: str, pay: dict) -> Optional[dict]:
    """QBO never answered: did it create the payment anyway? The vendor's payment on that date that links
    exactly these bills and amounts - then it is ours (a resend with the same requestid would return it too)."""
    want = {(lt["TxnType"], str(lt["TxnId"])): _r2(ln["Amount"]) for ln in pay["Line"] for lt in ln["LinkedTxn"][:1]}
    try:
        rows = _live(access, cid, "BillPayment", f"VendorRef = '{pay['VendorRef']['value']}' AND TxnDate = '{pay['TxnDate']}'")
    except Exception:                               # noqa: BLE001
        return None
    hits = [r for r in rows if _links_of(r) == want and abs(_r2(r.get("TotalAmt")) - _r2(pay["TotalAmt"])) < EPS]
    return hits[0] if len(hits) == 1 else None


def _after_post(res: dict, done: dict, want: dict, lv: dict, credits: list, method: str, acct: dict, date: str,
                access: str, cid: str) -> None:
    """Everything after QBO accepted a payment: off the pay run, into the mirror + the queue, the stub."""
    pid = str(done["Id"])
    bill_marks.set_pay_marks([{"bill_id": b, "selected": False} for b in want], _now_s())
    try:
        _to_mirror("BillPayment", [done])
        _to_mirror("Bill", list(_live_ids(access, cid, "Bill", want).values()))
        if credits:
            _to_mirror("VendorCredit", list(_live_ids(access, cid, "VendorCredit", [c["id"] for c in credits]).values()))
    except Exception as e:                          # noqa: BLE001  (the next refresh catches up)
        res["mirror_error"] = str(e)
    _record_push(res, want, lv, credits, method, acct, date)
    if method == "print":
        res["stub"] = {"filed": False, "waiting": "files itself once QuickBooks shows the printed check #"}
        return
    watch(only=pid)
    con = _pcon()
    try:
        r = con.execute("SELECT stub_file, stub_error FROM pay_push WHERE payment_id=?", (pid,)).fetchone()
    finally:
        con.close()
    res["stub"] = {"filed": bool(r and r["stub_file"] and not r["stub_error"]), "file": r and r["stub_file"],
                   "error": r and r["stub_error"]}


def commit(body: dict) -> dict:
    """The guarded write: authority -> the review's own token (single use) -> one writer -> _commit()."""
    ok, why = authorized()
    if not ok:
        _audit("pay", "refused", why)
        return {"ok": False, "error": why}
    t = _TOKENS.pop(str(body.get("run_token") or ""), None)      # spent now, whatever happens next
    if not t or time.time() - t["at"] > TOKEN_TTL:
        _audit("pay", "refused", "no live review token")
        return {"ok": False, "error": "this review has expired or was already used - run Pay in QuickBooks again"}
    for v in body.get("vendors") or []:
        vid = str(v.get("vendor_id") or "")
        shown = t["vendors"].get(vid)
        want = {str(b.get("bill_id")): _r2(b.get("amount")) for b in v.get("bills") or []}
        if shown is None or want != shown or not {str(c) for c in v.get("credits") or []} <= set(t["credits"].get(vid, {})):
            _audit("pay", "refused", f"vendor {vid} is not what the review showed")
            return {"ok": False, "error": "the payment asked for is not what the review showed - run Pay in QuickBooks again"}
    try:
        with _Locked():
            return _commit(body, t)
    except BlockingIOError:
        return {"ok": False, "error": "another payment write is running - wait for it to finish"}


def _mdy(iso: str) -> str:
    try:
        return dt.date.fromisoformat(str(iso)[:10]).strftime("%m/%d/%Y")
    except ValueError:
        return str(iso)


def approval_text(lines: list, acct: str, how: str) -> str:
    """What the macOS dialog says (owner 2026-09-28 layout): the title, then per vendor
    "vendor · amount · date" over "account · ref # / print status"."""
    out = ["approve these vendor payments:"]
    for vendor, amount, date, ref in lines[:6]:
        out.append(f"{vendor} · ${amount:,.2f} · {_mdy(date)}")
        out.append(f"{acct} · {ref or how}")
    if len(lines) > 6:
        out.append(f"… and {len(lines) - 6} more (${sum(x[1] for x in lines[6:]):,.2f})")
    if len(lines) > 1:
        out.append(f"Total ${sum(x[1] for x in lines):,.2f}")
    return "\n".join(out)


def _commit(body: dict, t: Optional[dict] = None) -> dict:
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

    today = dt.date.today()
    if date > (today + dt.timedelta(days=FUTURE_DAYS)).isoformat():
        return {"ok": False, "error": f"the payment date is more than {FUTURE_DAYS} days ahead"}
    access, cid = qbo_api.load_credentials()
    closed = _close_date(access, cid)
    if closed and date <= closed:
        return {"ok": False, "error": f"the books are closed through {closed} - pick a later payment date"}
    live = _plan_live(access, cid)["vendors"]

    # Touch ID, naming exactly what leaves: macOS' own dialog, which a web page cannot draw or click
    import presence                                 # noqa: PLC0415  (ledger-local)
    how = {"print": "to print", "check": "check", "ach": "ACH", "card": "card"}[method]
    plines = []
    for v in asked:
        vid = str(v.get("vendor_id") or "")
        net = _r2(sum(_r2(b["amount"]) for b in v.get("bills") or [])
                  - sum((t or {}).get("credits", {}).get(vid, {}).get(str(c), 0) for c in v.get("credits") or []))
        ref = "to print" if method == "print" else (("check #" if method == "check" else f"{how} ") + refs[vid]).strip()
        plines.append(((live.get(vid) or {}).get("vendor") or (t or {}).get("names", {}).get(vid) or vid, net, date, ref))
    ok, why = presence.confirm(approval_text(plines, acct["name"], how))
    if not ok:
        _audit("pay", "refused", f"Touch ID: {why}")
        return {"ok": False, "error": f"not confirmed on this Mac ({why}) - nothing was written"}
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
            done = _find_posted(access, cid, pay) if "did not answer" in str(e) else None
            if not done:
                res["error"] = str(e)
                _audit("pay", "failed", {"vendor": vid, "error": str(e), "request_id": rid})
                continue
        pid = str(done["Id"])
        _audit("pay", "written", {"vendor": vid, "payment_id": pid, "total": _r2(done.get("TotalAmt")),
                                  "ref": done.get("DocNumber"), "bills": list(want), "credits": [c["id"] for c in credits]})
        res.update({"ok": True, "payment_id": pid, "ref": done.get("DocNumber") or "", "total": _r2(done.get("TotalAmt")),
                    "bills": len(want), "credits": len(credits), "to_print": method == "print"})
        try:                                            # the payment IS in QBO from here: nothing below may lose it
            _after_post(res, done, want, lv, credits, method, acct, date, access, cid)
        except Exception as e:                          # noqa: BLE001
            res["after_error"] = f"written to QuickBooks, but the ledger's follow-up failed: {e}"
            _audit("pay", "after_error", {"payment_id": pid, "error": str(e)})
    posted = [r for r in results if r.get("ok")]
    return {"ok": bool(posted), "posted": len(posted), "failed": len(results) - len(posted),
            "total": _r2(sum(r["total"] for r in posted)), "results": results,
            **({} if posted else {"error": "; ".join(f"{r['vendor']}: {r.get('error')}" for r in results) or "nothing to pay"})}


# ───────────────────────── the pushed-payments queue + its watcher ─────────────────────────
# Every payment pushed from here is a row in `pay_push` (ledger DB) until it is MATCHED: QBO shows its
# number (a "To print" check gets one when it is printed in QBO - nobody types it) and its stub is filed
# under that number. `watch()` runs after every mirror refresh (refresh_mirror.py) and whenever the queue
# is opened; it reads the mirror only. A payment QBO deleted, voided or changed (bills / amounts not what
# was pushed) is FLAGGED with what happened to its bills; the owner marks it Resolved (optional note) or
# Keeps it in the queue with a reason (owner 2026-09-28).
_PUSH_TABLE = """CREATE TABLE IF NOT EXISTS pay_push (
    payment_id TEXT PRIMARY KEY, vendor_id TEXT, vendor TEXT, pushed_at TEXT NOT NULL, txn_date TEXT,
    method TEXT, account TEXT, total REAL, ref TEXT, bills TEXT, credits TEXT,
    qbo_state TEXT, qbo_ref TEXT, checked_at TEXT, flagged_at TEXT,
    stub_file TEXT, stub_ref TEXT, stub_at TEXT, stub_error TEXT,
    owner_state TEXT NOT NULL DEFAULT 'open', owner_note TEXT, noted_at TEXT)"""
FLAGS = ("deleted", "voided", "changed", "missing")
STATE_WORDS = {"to_print": "waiting to print", "matched": "matched", "deleted": "deleted in QuickBooks",
               "voided": "voided in QuickBooks", "changed": "changed in QuickBooks", "missing": "not in the mirror"}


def _pcon():
    import sqlite3                                  # noqa: PLC0415
    con = sqlite3.connect(str(bill_marks.LEDGER_DB))
    con.row_factory = sqlite3.Row
    con.execute(_PUSH_TABLE)
    return con


def _now() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def _record_push(res: dict, v: dict, lv: dict, credits: list, method: str, acct: dict, date: str) -> None:
    docs = {r["bill_id"]: r["doc"] for r in (lv or {}).get("bills", [])}
    bills = [{"bill_id": b, "doc": docs.get(b, ""), "amount": a} for b, a in v.items()]
    con = _pcon()
    try:
        con.execute("INSERT OR REPLACE INTO pay_push (payment_id, vendor_id, vendor, pushed_at, txn_date, method, account,"
                    " total, ref, bills, credits, qbo_state, qbo_ref, checked_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (res["payment_id"], res["vendor_id"], res["vendor"], _now(), date, method, acct["name"], res["total"],
                     res["ref"], json.dumps(bills), json.dumps([{"id": c["id"], "doc": c["doc"], "amount": c["balance"]} for c in credits]),
                     "to_print" if method == "print" else "matched", "" if method == "print" else res["ref"], _now()))
        con.commit()
    finally:
        con.close()


def _links_of(rec: dict) -> dict:
    return {(lt.get("TxnType"), str(lt.get("TxnId"))): _r2(ln.get("Amount"))
            for ln in rec.get("Line") or [] for lt in (ln.get("LinkedTxn") or [])[:1]}


def _judge(row, rec: Optional[dict], mst: dict) -> tuple:
    """(state, qbo ref) of one pushed payment against the mirror."""
    if not mst.get("present") or not rec:
        return "missing", ""
    if mst.get("deleted"):
        return "deleted", rec.get("DocNumber") or ""
    if mst.get("voided"):
        return "voided", rec.get("DocNumber") or ""
    pushed = {("Bill", b["bill_id"]): _r2(b["amount"]) for b in json.loads(row["bills"] or "[]")}
    pushed.update({("VendorCredit", c["id"]): _r2(c["amount"]) for c in json.loads(row["credits"] or "[]")})
    if _links_of(rec) != pushed or abs(_r2(rec.get("TotalAmt")) - _r2(row["total"])) > EPS:
        return "changed", rec.get("DocNumber") or ""
    if (rec.get("CheckPayment") or {}).get("PrintStatus") == "NeedToPrint":
        return "to_print", ""
    return "matched", rec.get("DocNumber") or ""


def watch(file_stubs: bool = True, only: str = "") -> dict:
    """Judge every open pushed payment against the mirror; file the stub of a newly matched one."""
    import bill_payment_stub as stubs               # noqa: PLC0415  (ledger-local module)
    con, mcon = _pcon(), mirror.connect()
    out = {"checked": 0, "matched": 0, "flagged": 0, "stubs": 0}
    try:
        rows = con.execute("SELECT * FROM pay_push WHERE owner_state != 'resolved'" + (" AND payment_id = ?" if only else ""),
                           ((only,) if only else ())).fetchall()
        for row in rows:
            pid = row["payment_id"]
            state, ref = _judge(row, mirror.get("BillPayment", pid, con=mcon), stubs.mirror_state(pid, con=mcon))
            out["checked"] += 1
            flagged_at = row["flagged_at"] if state in FLAGS else None
            if state in FLAGS and not flagged_at:
                flagged_at = _now()
                out["flagged"] += 1
            con.execute("UPDATE pay_push SET qbo_state=?, qbo_ref=?, checked_at=?, flagged_at=? WHERE payment_id=?",
                        (state, ref, _now(), flagged_at, pid))
            con.commit()                            # BEFORE the stub: print_stub records its print in this same DB
            if state == "matched" and file_stubs and (not row["stub_at"] or (row["stub_ref"] or "") != ref):
                st = _stub(pid)
                con.execute("UPDATE pay_push SET stub_file=?, stub_ref=?, stub_at=?, stub_error=? WHERE payment_id=?",
                            (st.get("file"), ref if st.get("filed") else row["stub_ref"], _now() if st.get("filed") else row["stub_at"],
                             None if st.get("filed") else st.get("error"), pid))
                out["stubs"] += bool(st.get("filed"))
                out["matched"] += 1
            con.commit()
    finally:
        con.close()
        mcon.close()
    return out


def _investigate(row, mcon) -> list:
    """What happened to the bills a flagged payment paid: open again, paid by another payment, or deleted."""
    bills = json.loads(row["bills"] or "[]")
    ids = {b["bill_id"] for b in bills}
    paid_by: dict = {}
    since = (dt.date.fromisoformat(row["txn_date"]) - dt.timedelta(days=60)).isoformat() if row["txn_date"] else "2000-01-01"
    for bp in mirror.load("BillPayment", "txn_date >= ?", (since,), con=mcon):
        if str(bp["Id"]) == row["payment_id"]:
            continue
        for (t, i), a in _links_of(bp).items():
            if t == "Bill" and i in ids:
                paid_by.setdefault(i, []).append(f"{bp.get('DocNumber') or 'payment ' + str(bp['Id'])} ({bp.get('TxnDate')}, {a:,.2f})")
    out = []
    for b in bills:
        r = mcon.execute("SELECT deleted, json FROM qbo_bill WHERE id=?", (b["bill_id"],)).fetchone()
        rec = mirror._unpack(r[1]) if r else None
        if not r:
            what = "not in the mirror"
        elif r[0]:
            what = "bill deleted in QuickBooks"
        elif paid_by.get(b["bill_id"]):
            what = "paid by " + "; ".join(paid_by[b["bill_id"]])
        elif _r2((rec or {}).get("Balance")) > EPS:
            what = f"owes {_r2(rec.get('Balance')):,.2f} now"
        else:
            what = "shows paid, no other payment found"
        out.append({"bill_id": b["bill_id"], "doc": b.get("doc") or b["bill_id"], "amount": b["amount"], "now": what})
    return out


def queue(show: str = "active") -> dict:
    """The pushed payments. active = not matched-and-stubbed yet, or flagged, or kept; all = history too."""
    watch()
    con, mcon = _pcon(), mirror.connect()
    try:
        rows = con.execute("SELECT * FROM pay_push ORDER BY pushed_at DESC").fetchall()
        out = []
        for r in rows:
            d = dict(r)
            d["bills"] = json.loads(d["bills"] or "[]")
            d["credits"] = json.loads(d["credits"] or "[]")
            d["state_words"] = STATE_WORDS.get(d["qbo_state"] or "", d["qbo_state"] or "")
            d["flagged"] = d["qbo_state"] in FLAGS
            d["done"] = d["qbo_state"] == "matched" and bool(d["stub_at"]) and (d["stub_ref"] or "") == (d["qbo_ref"] or "")
            if d["flagged"] and d["owner_state"] != "resolved":
                d["investigation"] = _investigate(r, mcon)
            active = d["owner_state"] != "resolved" and (not d["done"] or d["flagged"] or d["owner_state"] == "kept")
            if show == "all" or active:
                out.append(d)
    finally:
        con.close()
        mcon.close()
    return {"ok": True, "rows": out, "flagged": sum(1 for d in out if d["flagged"] and d["owner_state"] == "open")}


def mark(payment_id: str, state: str, note: str = "") -> dict:
    """The owner's ruling on a pushed payment: resolved (note optional) · kept (reason required) · open."""
    note = str(note or "").strip()
    if state not in ("resolved", "kept", "open"):
        return {"ok": False, "error": "bad state"}
    if state == "kept" and not note:
        return {"ok": False, "error": "say why it stays in the queue"}
    con = _pcon()
    try:
        n = con.execute("UPDATE pay_push SET owner_state=?, owner_note=?, noted_at=? WHERE payment_id=?",
                        (state, note or None, _now(), str(payment_id))).rowcount
        con.commit()
    finally:
        con.close()
    _audit("mark", state if n else "not found", {"payment_id": str(payment_id), "note": note})
    return {"ok": bool(n), **({} if n else {"error": "not in the queue"})}


def assign_number(payment_id: str, number: str) -> dict:
    """The guarded fallback write: authority -> one writer -> Touch ID (inside) -> _assign_number()."""
    ok, why = authorized()
    if not ok:
        _audit("number", "refused", why)
        return {"ok": False, "error": why}
    con = _pcon()
    try:
        row = con.execute("SELECT vendor, total FROM pay_push WHERE payment_id=?", (str(payment_id),)).fetchone()
    finally:
        con.close()
    if not row:                                     # only payments pushed from here are numbered here
        return {"ok": False, "error": "that payment was not pushed from the ledger"}
    try:
        with _Locked():
            return _assign_number(payment_id, number, row)
    except BlockingIOError:
        return {"ok": False, "error": "another payment write is running - wait for it to finish"}


def _assign_number(payment_id: str, number: str, row) -> dict:
    """Fallback only: a check printed OUTSIDE QuickBooks gets its # typed here (QBO numbers the ones it prints
    itself, and the watcher picks that up). Marks it printed, then the watcher files its stub."""
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
        dup += [x for x in mirror.load("Purchase", "doc_number = ?", (number,), con=con)   # a written check (Expense) too
                if _ref(x.get("AccountRef")) == acct]
    finally:
        con.close()
    if dup:
        d = dup[0]
        who = (d.get("VendorRef") or d.get("EntityRef") or {}).get("name")
        return {"ok": False, "error": f"check #{number} is already used on this account ({who}, {d.get('TxnDate')})"}
    import presence                                 # noqa: PLC0415  (ledger-local)
    ok, why = presence.confirm(approval_text([(row["vendor"], _r2(row["total"]), live.get("TxnDate"), f"printed as check #{number}")],
                                             (cp.get("BankAccountRef") or {}).get("name") or "", "check"))
    if not ok:
        _audit("number", "refused", f"Touch ID: {why}")
        return {"ok": False, "error": f"not confirmed on this Mac ({why}) - nothing was written"}
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
        _audit("number", "failed", {"payment_id": pid, "number": number, "error": str(e)})
        return {"ok": False, "error": str(e)}
    _audit("number", "written", {"payment_id": pid, "number": number})
    try:
        _to_mirror("BillPayment", [done])
    except Exception as e:                         # noqa: BLE001
        return {"ok": True, "payment_id": pid, "number": number, "mirror_error": str(e)}
    watch(only=pid)
    con = _pcon()
    try:
        r = con.execute("SELECT stub_ref, stub_error FROM pay_push WHERE payment_id=?", (pid,)).fetchone()
    finally:
        con.close()
    return {"ok": True, "payment_id": pid, "number": done.get("DocNumber"),
            "stub": {"filed": bool(r and r["stub_ref"] == number), "error": r and r["stub_error"]}}


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
    ap.add_argument("--queue", action="store_true", help="the pushed payments still waiting to match, or flagged")
    ap.add_argument("--watch", action="store_true", help="judge the pushed payments against the mirror, file matched stubs")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--authorize-this-mac", action="store_true",
                    help="make THIS Mac + macOS account the one that can pay from the ledger (Touch ID; replaces any other)")
    ap.add_argument("--authority", action="store_true", help="is this Mac authorized to pay?")
    a = ap.parse_args()
    if a.selftest:
        return _selftest()
    if a.authorize_this_mac:
        return authorize_this_mac()
    if a.authority:
        ok, why = authorized()
        print("authorized: this Mac + account can pay from the ledger" if ok else f"NOT authorized: {why}")
        return 0 if ok else 1
    if a.watch:
        w = watch()
        print(f"pay-bills watch: {w['checked']} pushed payments checked · {w['matched']} newly matched "
              f"({w['stubs']} stubs filed) · {w['flagged']} newly flagged")
        return 0
    if a.queue:
        for q in queue()["rows"]:
            print(f"{q['payment_id']:>10}  {q['txn_date']}  {q['vendor'][:36]:<36} {q['total']:>12,.2f}  "
                  f"{q['state_words']}{' #' + q['qbo_ref'] if q['qbo_ref'] else ''}  [{q['owner_state']}]")
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
