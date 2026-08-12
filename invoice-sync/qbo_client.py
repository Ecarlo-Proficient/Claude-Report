"""
qbo_client.py — Minimal QuickBooks Online client for the automation worker.

Auth via the qbo_vault Keychain blob at project root (Phase 1 of QBO Export v2).
One Touch ID prompt per process unlocks all QBO keys.

Public API:
    creds = load_qbo_credentials()
    invoices = query_all(creds, "Invoice", "Balance > '0'")
"""
from __future__ import annotations

import base64
import time
import logging
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests


# Repo root on sys.path so `shared/` (vault etc.) is importable.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_vault as kc  # noqa: E402


log = logging.getLogger("automation_worker.qbo")


API_BASE = "https://quickbooks.api.intuit.com"
MINOR_VERSION = "70"


class QBOError(Exception):
    """Raised on non-retryable QBO API errors."""


@dataclass(frozen=True)
class QBOCredentials:
    access_token: str
    company_id: str


def invoice_deep_link(company_id: str, txn_id: str) -> str:
    """Company-scoped deep link to a QBO invoice — Intuit's own 'copy link' form.

    The bare ``app.qbo.intuit.com/app/invoice?txnId=<id>`` link carries no company:
    the browser resolves that txnId inside WHATEVER Intuit company the session is
    currently on. With more than one Intuit company on the login, that silently
    opens a different company's transaction — the "random invoice" symptom. Routing
    through ``/app/login`` with ``deeplinkcompanyid`` pins the company first, then
    opens the invoice. This is exactly the URL the QBO API returns in an invoice's
    ``link`` field. ``company_id`` is passed in (from the loaded creds) so the realm
    never lands in source or logs.
    """
    from urllib.parse import quote

    pagereq = quote(f"invoice?txnId={txn_id}", safe="")
    return (
        f"https://qbo.intuit.com/app/login?pagereq={pagereq}"
        f"&deeplinkcompanyid={company_id}"
    )


def load_qbo_credentials() -> QBOCredentials:
    """
    Refresh the QBO access token using the stored refresh token.

    Reads QBO_CLIENT_ID / QBO_CLIENT_SECRET / QBO_COMPANY_ID / QBO_REFRESH_TOKEN
    from the qbo_vault Keychain blob (single Touch ID prompt). Persists the
    rotated refresh token if QBO returns a new one.

    Raises QBOError if the blob is missing or the refresh fails.
    """
    if not kc.has_credentials():
        raise QBOError(
            "No QBO credentials in Keychain. Run `python3 setup_qbo.py` "
            "from project root once."
        )
    creds = kc.get_all()
    required = ("QBO_CLIENT_ID", "QBO_CLIENT_SECRET", "QBO_COMPANY_ID", "QBO_REFRESH_TOKEN")
    missing = [k for k in required if not creds.get(k)]
    if missing:
        raise QBOError(
            f"QBO blob is incomplete (missing: {missing}). "
            f"Run `python3 setup_qbo.py` from project root."
        )

    basic = base64.b64encode(
        f"{creds['QBO_CLIENT_ID']}:{creds['QBO_CLIENT_SECRET']}".encode()
    ).decode()
    # Retry the bearer refresh on transient network/Intuit blips — the OAuth
    # endpoint occasionally times out its TLS handshake and a single POST would
    # crash the sync (Ted 2026-07-15). Retry timeouts/connection errors + 5xx;
    # a real 4xx (e.g. an expired refresh token) fails fast.
    r = None
    last = ""
    for attempt in range(4):                    # 1 try + 3 retries
        try:
            r = requests.post(
                "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer",
                headers={
                    "Authorization": f"Basic {basic}",
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Accept": "application/json",
                },
                data={"grant_type": "refresh_token",
                      "refresh_token": creds["QBO_REFRESH_TOKEN"]},
                timeout=30,
            )
            if r.status_code < 500:
                break
            last = f"status={r.status_code}"
        except requests.exceptions.RequestException as e:
            last = type(e).__name__
            r = None
        if attempt < 3:
            time.sleep((attempt + 1) * 3)       # 3s, 6s, 9s
    if r is None:
        raise QBOError(
            f"QBO token refresh failed after retries — {last} "
            "(network/Intuit timeout; retry the sync)"
        )
    if r.status_code != 200:
        raise QBOError(
            f"QBO token refresh failed status={r.status_code} body={r.text[:300]}"
        )
    body = r.json()
    new_rt = body.get("refresh_token")
    if new_rt and new_rt != creds["QBO_REFRESH_TOKEN"]:
        try:
            kc.put("QBO_REFRESH_TOKEN", new_rt)
            log.debug("Rotated QBO refresh token persisted to Keychain")
        except kc.SecretsError as e:
            log.warning("Could not persist rotated refresh token: %s", e)

    return QBOCredentials(
        access_token=body["access_token"],
        company_id=creds["QBO_COMPANY_ID"],
    )


def _api_get(creds: QBOCredentials, path: str, params: Optional[dict] = None) -> dict:
    """Single GET against the QBO REST API. Adds minorversion. Raises QBOError."""
    p = dict(params or {})
    p["minorversion"] = MINOR_VERSION
    r = requests.get(
        f"{API_BASE}{path}",
        headers={
            "Authorization": f"Bearer {creds.access_token}",
            "Accept": "application/json",
        },
        params=p,
        timeout=60,
    )
    if r.status_code != 200:
        raise QBOError(f"GET {path} → {r.status_code}: {r.text[:300]}")
    return r.json()


def query(creds: QBOCredentials, q: str) -> dict:
    """Single QBO QBO-SQL query."""
    return _api_get(creds, f"/v3/company/{creds.company_id}/query", {"query": q})


def invoice_exists(creds: QBOCredentials, invoice_id: str) -> bool:
    """
    Return True if an Invoice with this Id still exists in QBO, False if it
    has been DELETED.

    QBO's standard query endpoint does not return deleted entities — a deleted
    invoice yields an empty QueryResponse (no fault). A voided invoice, by
    contrast, still exists (Balance 0) and is returned here as True. So this
    distinguishes "deleted/removed entirely" from "still on file" (paid, voided,
    zero-balance, written-off, etc.).

    Raises on a real API error (network, auth, 5xx) so the caller can tell
    "couldn't determine" apart from "confirmed gone" and avoid acting on
    uncertainty. `invoice_id` is the QBO Id (numeric string), not the DocNumber.
    """
    safe_id = str(invoice_id).replace("'", "")  # Id is numeric; strip quotes defensively
    data = query(creds, f"SELECT Id FROM Invoice WHERE Id = '{safe_id}'")
    rows = data.get("QueryResponse", {}).get("Invoice", [])
    return bool(rows)


def fetch_deleted_invoice_ids(creds: QBOCredentials, changed_since: str) -> set:
    """
    Return the set of QBO Invoice Ids DELETED since `changed_since` (ISO 8601,
    e.g. '2026-06-01T00:00:00Z'), using QBO Change Data Capture (/cdc).

    CDC reports entities created/updated/deleted since the timestamp. Deleted
    entities come back carrying a top-level `status` of "Deleted" with minimal
    fields (Id + MetaData). We extract just those Ids — this is how we catch an
    invoice that was deleted AFTER it was already marked Paid in Notion, which
    the open-set sweep can't see.

    QBO caps `changedSince` at 30 days in the past; the caller clamps to that.
    Raises QBOError on API failure (caller skips the pass and leaves state alone).
    """
    data = _api_get(
        creds,
        f"/v3/company/{creds.company_id}/cdc",
        {"entities": "Invoice", "changedSince": changed_since},
    )
    deleted: set = set()
    for block in data.get("CDCResponse", []) or []:
        for qr in block.get("QueryResponse", []) or []:
            for inv in qr.get("Invoice", []) or []:
                if str(inv.get("status", "")).lower() == "deleted":
                    iid = inv.get("Id")
                    if iid:
                        deleted.add(str(iid))
    return deleted


def fetch_invoice(creds: QBOCredentials, invoice_id: str) -> Optional[Dict[str, Any]]:
    """
    Return the full QBO Invoice dict (including its Line[] array) for one Id,
    or None if it doesn't exist. Used to read line items for paid invoices that
    are no longer in the open-invoice fetch. Raises QBOError on API failure.
    """
    safe_id = str(invoice_id).replace("'", "")
    data = query(creds, f"SELECT * FROM Invoice WHERE Id = '{safe_id}'")
    rows = data.get("QueryResponse", {}).get("Invoice", [])
    return rows[0] if rows else None


def query_all(creds: QBOCredentials, entity: str, where: str = "") -> List[Dict[str, Any]]:
    """
    Run a paginated SELECT * on `entity` with optional WHERE clause.
    Returns the full list of records, draining all pages.
    """
    out: List[Dict[str, Any]] = []
    start = 1
    page_size = 500
    while True:
        q = f"SELECT * FROM {entity}"
        if where:
            q += f" WHERE {where}"
        q += f" STARTPOSITION {start} MAXRESULTS {page_size}"
        data = query(creds, q)
        batch = data.get("QueryResponse", {}).get(entity, [])
        if not batch:
            break
        out.extend(batch)
        if len(batch) < page_size:
            break
        start += page_size
    return out


def fetch_payment_dates(
    creds: QBOCredentials,
    lookback_months: int = 6,
) -> Dict[str, str]:
    """
    Build {invoice_id → most_recent_payment_TxnDate} for invoices that have
    received payments in the last `lookback_months`. Used by the flip-to-paid
    sweep to stamp the actual payment date on Notion's Paid Date field instead
    of "today" (which is wrong when the sync missed the payment by N days).

    Returns an ISO date string per invoice. Caller falls back to today if an
    invoice doesn't appear (voided, credit-memo applied, write-off, etc.).
    """
    import datetime as _dt
    cutoff = (_dt.date.today() - _dt.timedelta(days=lookback_months * 30)).isoformat()
    payments = query_all(creds, "Payment", where=f"TxnDate >= '{cutoff}'")

    inv_to_date: Dict[str, str] = {}
    for p in payments:
        pdate = (p.get("TxnDate") or "")[:10]
        if not pdate:
            continue
        for line in p.get("Line") or []:
            for lt in line.get("LinkedTxn") or []:
                if lt.get("TxnType") != "Invoice":
                    continue
                inv_id = str(lt.get("TxnId") or "")
                if not inv_id:
                    continue
                # Keep the LATEST payment date per invoice (handles partial payments
                # followed by a final one — final payment's date wins).
                prev = inv_to_date.get(inv_id)
                if prev is None or pdate > prev:
                    inv_to_date[inv_id] = pdate
    return inv_to_date


def fetch_term_map(creds: QBOCredentials) -> Dict[str, str]:
    """
    Build {term_id → term_name} for every QBO Term (active + inactive).
    Used to resolve Invoice.SalesTermRef.value (an ID) to a human-readable
    name like "Net 30" for the Net Terms select on Notion invoice rows.
    """
    terms = query_all(creds, "Term", where="Active IN (true, false)")
    return {
        str(t.get("Id")): (t.get("Name") or "").strip()
        for t in terms
        if t.get("Id")
    }


def fetch_customer_hierarchy(creds: QBOCredentials) -> Dict[str, str]:
    """
    Build {customer_id → root_parent_name} for every QBO customer (active +
    inactive). Walks ParentRef chains so a sub-customer (project) resolves
    to its top-level parent customer (the actual GC / builder name).

    Why: QBO Invoice.CustomerRef.name is the customer's local display name —
    for sub-customers (projects like 'RP7038-FTW') it's just the project
    string, not the parent. To match invoices to a customer list, we need
    the resolved root name.
    """
    # Pull active and inactive both — sub-customers can be inactive while
    # their parents are still around, and we need the lookup either way.
    customers = query_all(creds, "Customer", where="Active IN (true, false)")

    # First, index every customer by Id for fast parent lookup.
    by_id: Dict[str, dict] = {str(c.get("Id")): c for c in customers if c.get("Id")}

    def resolve_root(cust: dict, depth: int = 0) -> str:
        """Walk up ParentRef until null. Returns the root display name."""
        if depth > 20:
            # Defensive: ParentRef cycle. Shouldn't happen in real QBO data,
            # but if it does we don't want infinite recursion.
            return cust.get("DisplayName") or cust.get("Name") or ""
        parent_ref = cust.get("ParentRef") or {}
        parent_id = parent_ref.get("value")
        if not parent_id:
            return cust.get("DisplayName") or cust.get("Name") or ""
        parent = by_id.get(str(parent_id))
        if not parent:
            # Parent not in result set (deleted? permission?). Fall back to
            # the parent ref's name as a last-resort label.
            return parent_ref.get("name") or cust.get("DisplayName") or ""
        return resolve_root(parent, depth + 1)

    return {cid: resolve_root(c) for cid, c in by_id.items()}
