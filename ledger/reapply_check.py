#!/usr/bin/env python3
"""
reapply_check.py - put a stripped check back on the bills it paid.

QBO sometimes strips a paid check: an edit or delete on one bill unapplies the
WHOLE payment, and the money floats (Company > Checks QBO changed in the ledger).
The mirror's change log kept the check as it was BEFORE the strip, listing every
bill it paid and how much. This script reads that before copy and writes those
lines back onto the check, so nobody has to pick 129 bills out of Pay Bills by hand.

    ledger/reapply_check.py 25745              dry run: what would go back on
    ledger/reapply_check.py 25745 --commit     write it to QBO

The ledger's Checks QBO changed page runs the same two steps from a check's card
(Re-apply in QuickBooks -> dry run -> "Are you sure?" -> write): dry_run() / commit().

Guards (a line is restored only when ALL hold, live from QBO at run time):
  - the bill (or vendor credit) still exists, same vendor as the check
  - its open balance still covers the amount the check paid on it
    (paid elsewhere since -> skipped and listed, never partially applied)
  - lines already back on the check are kept as they are, never doubled
  - the restored total never exceeds the check amount
  - the check is re-read right before the write (fresh SyncToken)
The live check is saved to ~/Library/Logs/Proficient/reapply-check/ before any write.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import requests  # noqa: E402

from shared import qbo_api, qbo_mirror as mirror  # noqa: E402

EPS = 0.005
LOG_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "reapply-check"
ENTITY = {"Bill": "Bill", "VendorCredit": "VendorCredit"}


def _money(x: float) -> str:
    return f"{x:,.2f}"


def _fmt(d: str | None) -> str:
    try:
        return dt.date.fromisoformat(str(d)[:10]).strftime("%m/%d/%Y")
    except ValueError:
        return str(d or "")


def _links(pay: dict) -> list:
    """[(TxnType, TxnId, amount)] - what the payment applies, line by line."""
    out = []
    for ln in pay.get("Line") or []:
        for lt in ln.get("LinkedTxn") or []:
            out.append((lt.get("TxnType"), str(lt.get("TxnId")), round(float(ln.get("Amount") or 0), 2)))
    return out


def _find_before(con, check: str, payment_id: str | None, change_id: int | None):
    """(payment id, before copy, change row): the fullest pre-strip copy the change log holds."""
    if payment_id is None:
        hits = [p for p in mirror.load("BillPayment", con=con) if str(p.get("DocNumber") or "") == check]
        if not hits:
            raise LookupError(f"check {check}: no bill payment with that number in the mirror")
        logged = {c["rec_id"] for c in mirror.changes(con, entity="BillPayment", limit=100000) if c.get("has_before")}
        if len(hits) > 1 and sum(p["Id"] in logged for p in hits) == 1:
            hits = [p for p in hits if p["Id"] in logged]   # check numbers repeat across years and vendors
        if len(hits) > 1:
            listing = "; ".join(f"id {p['Id']} {_fmt(p.get('TxnDate'))} {(p.get('VendorRef') or {}).get('name')}"
                                f" {_money(float(p.get('TotalAmt') or 0))}" for p in hits)
            raise LookupError(f"check {check} matches {len(hits)} payments ({listing}) - rerun with --payment-id")
        payment_id = hits[0]["Id"]
    rows = [c for c in mirror.changes(con, entity="BillPayment", limit=100000)
            if c["rec_id"] == payment_id and c.get("has_before")]
    if change_id is not None:
        rows = [c for c in rows if c["id"] == change_id]
    if not rows:
        snap = _from_snapshot(con, payment_id)
        if snap:
            return payment_id, snap[0], snap[1]
        raise LookupError(f"payment {payment_id}: no before copy in the change log (stripped before it started "
                          "09/23/2026) and no ledger snapshot whose bills add up to the check - re-apply by hand")
    best = max(rows, key=lambda c: (c.get("lines_before") or 0, c["changed_at"]))
    return payment_id, mirror.change_before(con, best["id"]), best


def _from_snapshot(con, payment_id: str):
    """A check stripped BEFORE the change log (09/23/2026): the bills it paid from a saved copy of the ledger
    (its `bill_payment_line` table, loaded from QBO while the check was still whole). Used only when that
    copy's lines add up to the check amount to the cent. Newest copy first."""
    pay = next((p for p in mirror.load("BillPayment", con=con) if p["Id"] == payment_id), None)
    if not pay:
        return None
    total = round(float(pay.get("TotalAmt") or 0), 2)
    for f in sorted(mirror.db_path().parent.glob("ledger*.sqlite3*"), key=lambda f: f.stat().st_mtime, reverse=True):
        try:
            lc = sqlite3.connect(f"file:{f}?mode=ro", uri=True)
            try:
                got = lc.execute("SELECT bill_id, amount FROM bill_payment_line WHERE payment_id = ?", (payment_id,)).fetchall()
            finally:
                lc.close()
        except sqlite3.Error:
            continue
        if got and abs(round(sum(a for _, a in got), 2) - total) < EPS:
            when = dt.datetime.fromtimestamp(f.stat().st_mtime).isoformat(timespec="seconds")
            before = {"Line": [{"Amount": round(a, 2), "LinkedTxn": [{"TxnId": str(b), "TxnType": "Bill"}]} for b, a in got]}
            return before, {"changed_at": when, "source": f"ledger snapshot {f.name}"}
    return None


def _live(access: str, cid: str, typ: str, ids: list) -> dict:
    out = {}
    for i in range(0, len(ids), 40):
        chunk = ",".join(f"'{x}'" for x in ids[i:i + 40])
        q = f"SELECT * FROM {typ} WHERE Id IN ({chunk}) MAXRESULTS 1000"
        for r in qbo_api._api_get(f"/v3/company/{cid}/query", access, {"query": q}).get("QueryResponse", {}).get(typ, []):
            out[str(r["Id"])] = r
    return out


def _get_payment(access: str, cid: str, pid: str) -> dict:
    return qbo_api._api_get(f"/v3/company/{cid}/billpayment/{pid}", access)["BillPayment"]


def plan(access: str, cid: str, pid: str, before: dict) -> dict:
    live = _get_payment(access, cid, pid)
    vendor = str((live.get("VendorRef") or {}).get("value") or "")
    total = round(float(live.get("TotalAmt") or 0), 2)
    now = {(t, i): a for t, i, a in _links(live)}
    want = [(t, i, a) for t, i, a in _links(before)]

    recs = {}
    for typ in ENTITY:
        ids = sorted({i for t, i, _ in want if t == typ})
        if ids:
            recs[typ] = _live(access, cid, typ, ids)

    restore, skipped, already = [], [], []
    for typ, tid, amt in want:
        if now.get((typ, tid), 0) >= amt - EPS:
            already.append((typ, tid, amt))
            continue
        if typ not in ENTITY:
            skipped.append((typ, tid, amt, f"{typ} line - re-apply by hand"))
            continue
        r = recs.get(typ, {}).get(tid)
        if not r:
            skipped.append((typ, tid, amt, "deleted in QBO"))
            continue
        doc = r.get("DocNumber") or tid
        if str((r.get("VendorRef") or {}).get("value") or "") != vendor:
            skipped.append((typ, doc, amt, "different vendor now"))
            continue
        need = round(amt - now.get((typ, tid), 0), 2)
        bal = float(r.get("Balance") or 0)
        if bal < need - EPS:
            skipped.append((typ, doc, amt, f"open balance {_money(bal)} < {_money(need)} (paid elsewhere since?)"))
            continue
        restore.append({"type": typ, "id": tid, "doc": doc, "date": r.get("TxnDate"), "amount": amt,
                        "balance": bal, "memo": (r.get("PrivateNote") or "").split("\n")[0][:60]})

    lines = [dict(ln) for ln in live.get("Line") or []]
    have = {(t, i) for t, i, _ in _links(live)}
    for x in restore:
        if (x["type"], x["id"]) in have:        # partially on already: raise that line to the before amount
            for ln in lines:
                if any(lt.get("TxnType") == x["type"] and str(lt.get("TxnId")) == x["id"] for lt in ln.get("LinkedTxn") or []):
                    ln["Amount"] = x["amount"]
        else:
            lines.append({"Amount": x["amount"], "LinkedTxn": [{"TxnId": x["id"], "TxnType": x["type"]}]})

    def net(ls):
        b = sum(float(ln.get("Amount") or 0) for ln in ls
                for lt in (ln.get("LinkedTxn") or [])[:1] if lt.get("TxnType") != "VendorCredit")
        c = sum(float(ln.get("Amount") or 0) for ln in ls
                for lt in (ln.get("LinkedTxn") or [])[:1] if lt.get("TxnType") == "VendorCredit")
        return round(b - c, 2)

    return {"live": live, "total": total, "applied_now": net(live.get("Line") or []), "applied_after": net(lines),
            "lines": lines, "restore": restore, "skipped": skipped, "already": already}


def post(access: str, cid: str, live: dict, lines: list) -> dict:
    body = {k: v for k, v in live.items() if k not in ("MetaData", "domain", "sparse")}
    body["Line"] = lines
    r = requests.post(f"{qbo_api.API_BASE}/v3/company/{cid}/billpayment",
                      headers={"Authorization": f"Bearer {qbo_api._CURRENT_ACCESS or access}",
                               "Content-Type": "application/json", "Accept": "application/json"},
                      params={"minorversion": qbo_api.MINOR_VERSION}, json=body, timeout=120)
    if r.status_code != 200:
        try:
            e = r.json()["Fault"]["Error"][0]
            msg = f"{e.get('Message')}: {e.get('Detail')}"
        except Exception:                                   # noqa: BLE001
            msg = r.text[:400]
        raise RuntimeError(f"QBO refused the update ({r.status_code}): {msg} - nothing was changed")
    return r.json()["BillPayment"]


def _json_plan(pid: str, before: dict, ch: dict, p: dict) -> dict:
    live = p["live"]
    return {"ok": True, "payment_id": pid, "check": live.get("DocNumber") or "",
            "vendor": (live.get("VendorRef") or {}).get("name") or "", "txn_date": live.get("TxnDate"),
            "total": p["total"], "applied_now": p["applied_now"], "applied_after": p["applied_after"],
            "floating_after": round(p["total"] - p["applied_after"], 2), "sync_token": live.get("SyncToken"),
            "before_from": ch["changed_at"], "before_source": ch.get("source") or "change log", "before_lines": len(_links(before)), "already": len(p["already"]),
            "restore": p["restore"], "restore_amt": round(sum(x["amount"] for x in p["restore"]), 2),
            "skipped": [{"type": t, "doc": d, "amount": a, "why": w} for t, d, a, w in p["skipped"]],
            "over": p["applied_after"] > p["total"] + EPS}


def dry_run(payment_id: str) -> dict:
    """What a re-apply would write, read live from QBO. Writes nothing."""
    con = mirror.connect()
    try:
        pid, before, ch = _find_before(con, "", str(payment_id), None)
    finally:
        con.close()
    access, cid = qbo_api.load_credentials()
    return _json_plan(pid, before, ch, plan(access, cid, pid, before))


def commit(payment_id: str, sync_token: str, n: int, amount: float) -> dict:
    """Write the re-apply - only if the plan is still exactly the one the owner confirmed
    (same check version, same number of bills, same amount)."""
    con = mirror.connect()
    try:
        pid, before, ch = _find_before(con, "", str(payment_id), None)
    finally:
        con.close()
    access, cid = qbo_api.load_credentials()
    p = plan(access, cid, pid, before)
    live = p["live"]
    amt = round(sum(x["amount"] for x in p["restore"]), 2)
    if live.get("SyncToken") != str(sync_token) or len(p["restore"]) != int(n) or abs(amt - float(amount)) > EPS:
        return {"ok": False, "error": "The check or its bills changed in QuickBooks since the dry run - run the dry run again."}
    if not p["restore"]:
        return {"ok": False, "error": "Nothing to put back."}
    if p["applied_after"] > p["total"] + EPS:
        return {"ok": False, "error": "Restoring would apply more than the check amount - nothing written."}
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    backup = LOG_DIR / f"{pid}-{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}.json"
    backup.write_text(json.dumps({"live_before": live, "before_copy": before, "restore": p["restore"],
                                  "skipped": p["skipped"]}, indent=1, default=str))
    done = post(access, cid, live, p["lines"])
    lk = _links(done)
    applied = round(sum(a for t, _, a in lk if t != "VendorCredit") - sum(a for t, _, a in lk if t == "VendorCredit"), 2)
    return {"ok": True, "payment_id": pid, "check": live.get("DocNumber") or "", "applied": applied,
            "lines": len(lk), "total": p["total"], "backup": str(backup)}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("check", help="check number (the payment's DocNumber), e.g. 25745")
    ap.add_argument("--payment-id", help="QBO BillPayment Id, when two payments share the check number")
    ap.add_argument("--change-id", type=int, help="use this change-log entry's before copy (default: the fullest one)")
    ap.add_argument("--list", action="store_true", help="print every bill that goes back on, not just the totals")
    ap.add_argument("--commit", action="store_true", help="write to QBO (default is a dry run)")
    a = ap.parse_args()

    con = mirror.connect()
    try:
        pid, before, ch = _find_before(con, a.check, a.payment_id, a.change_id)
    except LookupError as e:
        sys.exit(str(e))
    finally:
        con.close()
    access, cid = qbo_api.load_credentials()
    p = plan(access, cid, pid, before)
    live = p["live"]

    print(f"check {a.check}  {(live.get('VendorRef') or {}).get('name')}  {_fmt(live.get('TxnDate'))}  "
          f"total {_money(p['total'])}")
    src = ch.get("source") or "the change log"
    print(f"  what it paid, from {src} ({_fmt(ch['changed_at'])}): {len(_links(before))} lines")
    print(f"  applied now {_money(p['applied_now'])}  ->  after {_money(p['applied_after'])}"
          f"  (floating after: {_money(p['total'] - p['applied_after'])})")
    print(f"  already on the check: {len(p['already'])}   to put back: {len(p['restore'])} "
          f"({_money(sum(x['amount'] for x in p['restore']))})   skipped: {len(p['skipped'])}")
    if a.list:
        for x in sorted(p["restore"], key=lambda x: x["date"] or ""):
            print(f"    + {x['doc']:<14} {_fmt(x['date'])}  {_money(x['amount']):>12}  {x['memo']}")
    for t, doc, amt, why in p["skipped"]:
        print(f"    SKIP {t} {doc}  {_money(amt)}  - {why}")

    if p["applied_after"] > p["total"] + EPS:
        sys.exit("STOP: restoring would apply more than the check amount - nothing written")
    if not p["restore"]:
        print("nothing to put back")
        return 0
    if not a.commit:
        print("\ndry run - nothing written. Add --commit to write it to QBO.")
        return 0

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = LOG_DIR / f"{pid}-{stamp}.json"
    fresh = _get_payment(access, cid, pid)                  # fresh SyncToken, and nobody changed it meanwhile
    if fresh.get("SyncToken") != live.get("SyncToken"):
        sys.exit("STOP: the check changed in QBO while this ran - run it again")
    backup.write_text(json.dumps({"live_before": fresh, "before_copy": before, "restore": p["restore"],
                                  "skipped": p["skipped"]}, indent=1, default=str))
    try:
        done = post(access, cid, fresh, p["lines"])
    except RuntimeError as e:
        sys.exit(str(e))
    applied = sum(x[2] for x in _links(done) if x[0] != "VendorCredit") - \
        sum(x[2] for x in _links(done) if x[0] == "VendorCredit")
    print(f"\nwritten. QBO now shows {_money(applied)} applied on check {a.check} "
          f"({len(_links(done))} lines). Backup: {backup}")
    print("Refresh the mirror (ledger gear > sync, or ledger/refresh_mirror.py) so the ledger clears it.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
