"""
shared/payment_reapply.py - put a client payment back on the invoices and credits QBO took it off.

The customer-side twin of ledger/reapply_check.py (owner 2026-10-16: "if any client payment gets
offed we need to fix JUST like the bill payments getting unapplied"). Changing a credit memo's
customer makes QBO drop the credit off the payment that used it and take the same amount off an
invoice line instead; the payment's SyncToken does not move, so the raw mirror still holds the
payment as it was. That copy is the target: re-save the payment with exactly its lines.

Used by the ledger (Payments QBO changed -> Re-apply in QuickBooks: dry_run -> "Are you sure?" ->
ledger/reapply_payment.commit) and by one-offs/credit_memo_reclass.py (move a credit, then put its payment back).
Reads and checks only: the write itself is the caller's `post` (shared/ ships to the office server, read-only).

Guards, live from QBO at run time (plan() is pure, so tests drive it without QuickBooks):
  - same payment: total, date and customer as the copy on file
  - nothing new on the payment that the copy does not have (somebody re-applied it differently)
  - every document still exists and its open balance covers what goes back on it
  - the payment is re-read right before the write (fresh SyncToken); after the write every
    line, the payment total / unapplied and every document's open balance are checked to the cent
The live payment and documents are saved to ~/Library/Logs/Proficient/payment-reapply/ before any write.
"""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from shared import qbo_api, qbo_mirror as mirror

EPS = 0.005
DOCS = ("Invoice", "CreditMemo")
NOTHING = "nothing to put back - the payment already matches the copy on file"
LOG_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "payment-reapply"


def links(pay: dict) -> list:
    """[(TxnType, TxnId, amount)] - what the payment applies, line by line."""
    return [(lt["TxnType"], str(lt["TxnId"]), round(float(ln.get("Amount") or 0), 2))
            for ln in (pay or {}).get("Line") or [] for lt in ln.get("LinkedTxn") or []]


def _by_doc(pay: dict) -> dict:
    out: dict = {}
    for t, i, a in links(pay):
        out[(t, i)] = round(out.get((t, i), 0) + a, 2)
    return out


def open_of(rec: dict) -> float:
    v = rec.get("RemainingCredit") if "RemainingCredit" in rec else rec.get("Balance")
    return round(float(v or 0), 2)


def plan(target: dict, live: dict, docs: dict) -> dict:
    """What re-saving `live` with `target`'s lines puts back. docs: {(type, id): live record}.
    -> {ok, back: [{type, id, doc, amount, open_now, open_after}], why: [reasons it must not be written]}"""
    why = []
    for k, label in (("TotalAmt", "amount"), ("TxnDate", "date")):
        if str(target.get(k)) != str(live.get(k)) and not (
                k == "TotalAmt" and abs(float(target.get(k) or 0) - float(live.get(k) or 0)) < EPS):
            why.append(f"the payment's {label} changed ({target.get(k)} -> {live.get(k)})")
    if (target.get("CustomerRef") or {}).get("value") != (live.get("CustomerRef") or {}).get("value"):
        why.append("the payment's customer changed")
    want, now = _by_doc(target), _by_doc(live)
    for k, a in now.items():
        if a > want.get(k, 0) + EPS:
            why.append(f"the payment now applies {a:,.2f} to {k[0]} {k[1]} - more than the copy on file")
    back = []
    for (t, i), a in want.items():
        missing = round(a - now.get((t, i), 0), 2)
        if missing < EPS:
            continue
        rec = docs.get((t, i))
        if t not in DOCS:
            why.append(f"{t} {i} came off - only invoices and credit memos are put back")
            continue
        if not rec:
            why.append(f"{t} {i} no longer exists")
            continue
        op = open_of(rec)
        if op + EPS < missing:
            why.append(f"{t} {rec.get('DocNumber') or i} is open {op:,.2f}, less than the {missing:,.2f} to put back "
                       "(used somewhere else since)")
        back.append({"type": t, "id": i, "doc": rec.get("DocNumber") or i,
                     "customer": (rec.get("CustomerRef") or {}).get("name") or "",
                     "amount": missing, "open_now": op, "open_after": round(op - missing, 2)})
    if not back and not why:
        why.append(NOTHING)
    return {"ok": not why, "back": back, "why": why}


def body_for(live: dict, target: dict) -> dict:
    body = {k: v for k, v in live.items() if k not in ("MetaData", "domain", "sparse")}
    body["Line"] = [{"Amount": ln["Amount"], "LinkedTxn": [{"TxnId": lt["TxnId"], "TxnType": lt["TxnType"]}
                                                           for lt in ln.get("LinkedTxn") or []]}
                    for ln in target.get("Line") or [] if ln.get("LinkedTxn")]
    return body


# ── QuickBooks ──────────────────────────────────────────────────────────────────────────────────────────

def _get(access: str, cid: str, ent: str, rid: str) -> dict | None:
    try:
        return qbo_api._api_get(f"/v3/company/{cid}/{ent.lower()}/{rid}", access)[ent]
    except Exception as e:                                  # noqa: BLE001
        if "6240" in str(e) or "not found" in str(e).lower() or "Object Not Found" in str(e):
            return None
        raise


def _target(payment_id: str, target: dict | None) -> dict:
    if target is not None:
        return target
    rec = mirror.get("Payment", str(payment_id))
    if not rec:
        raise RuntimeError(f"payment {payment_id} is not in the mirror")
    return rec


def _live(access: str, cid: str, payment_id: str, target: dict) -> tuple:
    live = _get(access, cid, "Payment", str(payment_id))
    if not live:
        raise RuntimeError(f"payment {payment_id} no longer exists in QuickBooks")
    keys = set(_by_doc(target)) | set(_by_doc(live))
    docs = {k: d for k in keys if (d := _get(access, cid, k[0], k[1])) is not None}
    return live, docs


def _summary(payment_id: str, target: dict, live: dict, p: dict) -> dict:
    return {"ok": p["ok"], "payment_id": str(payment_id), "ref": target.get("PaymentRefNum") or "",
            "date": target.get("TxnDate"), "customer": (target.get("CustomerRef") or {}).get("name") or "",
            "total": round(float(target.get("TotalAmt") or 0), 2), "sync_token": live.get("SyncToken"),
            "back": p["back"], "why": p["why"], "n": len(p["back"]),
            "amount": round(sum(b["amount"] for b in p["back"]), 2)}


def dry_run(payment_id: str, target: dict | None = None, creds: tuple | None = None) -> dict:
    """What the re-apply would write - live read, no write."""
    target = _target(payment_id, target)
    access, cid = creds or qbo_api.load_credentials()
    live, docs = _live(access, cid, payment_id, target)
    return _summary(payment_id, target, live, plan(target, live, docs))


def commit(post, payment_id: str, sync_token: str | None = None, n: int | None = None, amount: float | None = None,
           target: dict | None = None, creds: tuple | None = None, expect_open: dict | None = None) -> dict:
    """Write the re-apply - only if the plan is still the one confirmed (same payment version, same count, same
    amount; None skips that check for a caller that plans and writes in one step). expect_open: {(type, id): open}
    overrides what each document must show afterwards (a caller that knows the balances from before its own change).
    post(access, cid, entity, body) -> saved record: the WRITE, handed in by the caller (ledger/reapply_payment.py,
    the cleanup one-off). This module never writes itself - shared/ is copied to the office server, which must not
    be able to write to QuickBooks (tests/test_office_server_image.py)."""
    target = _target(payment_id, target)
    access, cid = creds or qbo_api.load_credentials()
    live, docs = _live(access, cid, payment_id, target)
    p = plan(target, live, docs)
    s = _summary(payment_id, target, live, p)
    if sync_token is not None and (live.get("SyncToken") != str(sync_token) or s["n"] != int(n or 0)
                                   or abs(s["amount"] - float(amount or 0)) > EPS):
        return {"ok": False, "error": "The payment or its invoices changed in QuickBooks since the dry run - "
                                      "run the dry run again."}
    if not p["back"] and p["why"] == [NOTHING]:             # QBO kept it on: still prove every balance
        bad = verify(access, cid, payment_id, target, dict(expect_open or {}))
        return {"ok": not bad, "payment_id": str(payment_id), "n": 0, "amount": 0.0, "nothing": True,
                **({"error": "; ".join(bad)} if bad else {})}
    if not p["ok"]:
        return {"ok": False, "error": "; ".join(p["why"])}
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    backup = LOG_DIR / f"{payment_id}-{dt.datetime.now():%Y%m%d-%H%M%S}.json"
    backup.write_text(json.dumps({"live_before": live, "target": target, "docs_before": {f"{k[0]}:{k[1]}": v for k, v in docs.items()},
                                  "plan": p}, indent=1, default=str))
    post(access, cid, "Payment", body_for(live, target))
    want_open = {(b["type"], b["id"]): b["open_after"] for b in p["back"]}
    want_open.update(expect_open or {})
    bad = verify(access, cid, payment_id, target, want_open)
    return {"ok": not bad, "payment_id": str(payment_id), "ref": target.get("PaymentRefNum") or "",
            "n": s["n"], "amount": s["amount"], "backup": str(backup),
            **({"error": "Written, but not back exactly: " + "; ".join(bad)} if bad else {})}


def verify(access: str, cid: str, payment_id: str, target: dict, want_open: dict) -> list:
    """Every way the payment is still off - empty = back exactly as the copy on file."""
    now = _get(access, cid, "Payment", str(payment_id)) or {}
    bad = []
    if sorted(links(now)) != sorted(links(target)):
        bad.append("the payment's lines differ from the copy on file")
    for k in ("TotalAmt", "UnappliedAmt"):
        if abs(float(now.get(k) or 0) - float(target.get(k) or 0)) > EPS:
            bad.append(f"{k} {now.get(k)} vs {target.get(k)}")
    for (t, i), want in want_open.items():
        rec = _get(access, cid, t, i) or {}
        if abs(open_of(rec) - float(want or 0)) > EPS:
            bad.append(f"{t} {rec.get('DocNumber') or i} open {open_of(rec):,.2f}, expected {float(want or 0):,.2f}")
    return bad
