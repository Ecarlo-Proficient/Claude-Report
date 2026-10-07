"""
qbo_payment_notify.py - one Teams card per payment a client made through QuickBooks Payments
(the user 2026-10-07, the payment intake process).

A client who pays from the e-invoice link (bank or card) never hands anyone a check, so nothing in
the office says it came in. This posts a card to the payments channel the first time the mirror
shows one; the payroll manager prints it for the invoice file and the deposit sheet.

What counts: a QBO `Payment` whose `TxnSource` is EInvoice / INTUITMASPAYMENT or whose payment
method is "QuickBooks Payments-..." (checked 10/07/2026: every online payment carries both; checks,
ACH and wires the invoice clerk types in carry neither).

Rules:
  * Reads the QBO mirror only (sync-all and the office server refresh it right before AR).
  * One card per payment, ever: posted ids are kept in STATE_DIR/qbo_payments_posted.json.
  * The first run with a webhook marks every payment already there as seen and posts nothing, and
    only payments created in the last LOOKBACK_HOURS are ever posted - so a new box, a lost state
    file or the writer moving to the office server never replays old payments.
  * No webhook (TEAMS_WEBHOOK_PAYMENTS) = off. Dry run = log what would post, state untouched.
  * A failed post is not marked seen; the next run tries again.
"""
from __future__ import annotations

import datetime as dt
import json
import logging
import sys
from pathlib import Path
from typing import Dict, List, Optional

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_mirror  # noqa: E402
from shared import teams_notify as teams  # noqa: E402

log = logging.getLogger(__name__)

STATE_FILE = "qbo_payments_posted.json"
LOOKBACK_HOURS = 72
KEEP_DAYS = 30
ONLINE_SOURCES = {"EINVOICE", "INTUITMASPAYMENT"}
QBO_LINK = "https://qbo.intuit.com/app/recvpayment?txnId={}"


def is_online(pay: dict, method_names: Dict[str, str]) -> bool:
    """Paid through QuickBooks Payments (the e-invoice link), not typed in from a check."""
    if str(pay.get("TxnSource") or "").upper() in ONLINE_SOURCES:
        return True
    method = method_names.get(str((pay.get("PaymentMethodRef") or {}).get("value") or ""), "")
    return method.lower().startswith("quickbooks payments")


def _created(pay: dict) -> Optional[dt.datetime]:
    s = (pay.get("MetaData") or {}).get("CreateTime") or ""
    try:
        return dt.datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None


def _money(v) -> str:
    return f"${float(v or 0):,.2f}"


def _date(s: str) -> str:
    try:
        return dt.date.fromisoformat(s).strftime("%m/%d/%Y")
    except (TypeError, ValueError):
        return s or ""


def _client(cust_id: str, customers: Dict[str, dict]) -> str:
    """The top parent (the GC) of the paying customer; stops at the deepest known ancestor."""
    seen, cur, name = set(), customers.get(cust_id), ""
    while cur and cur.get("Id") not in seen:
        seen.add(cur.get("Id"))
        name = cur.get("DisplayName") or name
        cur = customers.get(str((cur.get("ParentRef") or {}).get("value") or ""))
    return name


def card(pay: dict, method: str, client: str, project: str, invoices: List[dict]) -> dict:
    """The Teams card: who paid, how much, which invoice(s) and what is still open on each."""
    inv_lines = []
    for inv in invoices:
        bal = float(inv.get("Balance") or 0)
        state = "paid in full" if bal <= 0.005 else f"{_money(bal)} still open"
        inv_lines.append(f"#{inv.get('DocNumber') or '?'} - {state}")
    facts = [
        {"title": "Client", "value": client or "-"},
        {"title": "Project", "value": project or "-"},
        {"title": "Amount", "value": _money(pay.get("TotalAmt"))},
        {"title": "Invoice", "value": "; ".join(inv_lines) or "not applied to an invoice"},
        {"title": "Paid by", "value": method.replace("QuickBooks Payments-", "") or "QuickBooks Payments"},
        {"title": "Date", "value": _date(pay.get("TxnDate") or "")},
    ]
    body = [
        {"type": "TextBlock", "size": "Medium", "weight": "Bolder", "color": "Good", "wrap": True,
         "text": "QuickBooks payment received"},
        {"type": "FactSet", "facts": facts},
        {"type": "ActionSet", "actions": [
            {"type": "Action.OpenUrl", "title": "Open in QuickBooks", "url": QBO_LINK.format(pay.get("Id"))}]},
    ]
    return teams._envelope(body)


def _load_state(path: Path) -> Optional[dict]:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError):
        return None


def _save_state(path: Path, state: dict, now: dt.datetime) -> None:
    cutoff = (now - dt.timedelta(days=KEEP_DAYS)).isoformat()
    state["posted"] = {k: v for k, v in state.get("posted", {}).items() if v >= cutoff}
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, indent=2))
    tmp.replace(path)


def run(webhook: str, state_dir: Path, dry_run: bool = False,
        now: Optional[dt.datetime] = None) -> dict:
    """Post every new online payment once. Returns counts for the run log; never raises on a post."""
    now = now or dt.datetime.now(dt.timezone.utc)
    out = {"posted": 0, "failed": 0, "would_post": 0, "seeded": 0}
    if not webhook and not dry_run:
        return out
    if not qbo_mirror.serves("Payment"):
        log.warning("qbo payments: the mirror holds no payments - run ledger/refresh_mirror.py")
        return out
    methods = {r["Id"]: r.get("Name") or "" for r in qbo_mirror.load("PaymentMethod")}
    online = [p for p in qbo_mirror.load("Payment") if is_online(p, methods)]

    path = Path(state_dir) / STATE_FILE
    state = _load_state(path)
    if state is None:
        if dry_run:
            log.info("qbo payments (dry run): first run would mark %d payments seen, post none", len(online))
            return out
        state = {"seeded": now.isoformat(), "posted": {p["Id"]: now.isoformat() for p in online}}
        _save_state(path, state, now)
        out["seeded"] = len(online)
        log.info("qbo payments: first run - %d existing payments marked seen, none posted", len(online))
        return out

    since = now - dt.timedelta(hours=LOOKBACK_HOURS)
    fresh = [p for p in online if p["Id"] not in state.get("posted", {})
             and (_created(p) or now) >= since]
    if not fresh:
        return out
    customers = {c["Id"]: c for c in qbo_mirror.load("Customer")}
    for p in sorted(fresh, key=lambda r: (_created(r) or now)):
        cust_id = str((p.get("CustomerRef") or {}).get("value") or "")
        project = (customers.get(cust_id) or {}).get("DisplayName") or (p.get("CustomerRef") or {}).get("name") or ""
        inv_ids = [lt.get("TxnId") for ln in p.get("Line") or [] for lt in ln.get("LinkedTxn") or []
                   if lt.get("TxnType") == "Invoice"]
        invoices = [i for i in (qbo_mirror.get("Invoice", x) for x in inv_ids) if i]
        method = methods.get(str((p.get("PaymentMethodRef") or {}).get("value") or ""), "")
        payload = card(p, method, _client(cust_id, customers), project, invoices)
        label = f"payment {p['Id']} {_money(p.get('TotalAmt'))} invoice(s) {[i.get('DocNumber') for i in invoices]}"
        if dry_run:
            out["would_post"] += 1
            log.info("qbo payments (dry run): would post %s", label)
            continue
        if teams.post(webhook, payload):
            state.setdefault("posted", {})[p["Id"]] = now.isoformat()
            out["posted"] += 1
            log.info("qbo payments: posted %s", label)
        else:
            out["failed"] += 1
            log.warning("qbo payments: Teams rejected %s - retried next run", label)
    if not dry_run:
        _save_state(path, state, now)
    return out
