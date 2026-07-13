"""
shared/qbo_api.py — QBO auth + API helpers shared by report-style tools.

Extracted verbatim from project-pnl/project_pnl_export.py (2026-07-13
restructure) so the WIP readers no longer have to load that file by raw
path with importlib. project_pnl_export and cp_wip_reader both import
from here now.

Auth model: token-style — `access, company_id = load_credentials()` does
one refresh-token exchange against the qbo_vault Keychain blob (single
Touch ID per run) and persists a rotated refresh token if QBO returns one.

Retry model: `_api_get` retries transient 5xx/429 and network timeouts
with exponential backoff (8 attempts, capped at 30s + jitter) — absorbs a
QBO outage of ~2–3 minutes. 4xx errors raise immediately.
"""
from __future__ import annotations

import base64
import re
import sys
from typing import Dict, List, Optional, Tuple

import requests

try:
    from . import qbo_vault as kc
except ImportError:  # run outside the package (script dir on sys.path)
    import qbo_vault as kc  # type: ignore


API_BASE = "https://quickbooks.api.intuit.com"
MINOR_VERSION = "70"

# Project numbers: RP#### (opt. suffix like -FTW), CP###(#), MFD###(#).
# Strict match — RP7186 and RP7186-FTW are DIFFERENT projects (no family
# rollup, ever).
PROJ_RE = re.compile(
    r"\b(RP\d{4}(?:-[A-Za-z]{2,6})?|CP\d{3,4}(?:-[A-Za-z0-9]{1,6})?|MFD\d{3,4})\b",
    re.IGNORECASE,
)


def load_credentials() -> Tuple[str, str]:
    if not kc.has_credentials():
        print("✗  No QBO credentials in Keychain. Run: python3 shared/setup_qbo.py")
        sys.exit(1)
    creds = kc.get_all()
    required = ["QBO_CLIENT_ID", "QBO_CLIENT_SECRET", "QBO_COMPANY_ID", "QBO_REFRESH_TOKEN"]
    missing = [k for k in required if not creds.get(k)]
    if missing:
        print(f"✗  Missing credentials: {', '.join(missing)}")
        sys.exit(1)
    basic = base64.b64encode(
        f"{creds['QBO_CLIENT_ID']}:{creds['QBO_CLIENT_SECRET']}".encode()
    ).decode()
    r = requests.post(
        "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer",
        headers={
            "Authorization": f"Basic {basic}",
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        },
        data={"grant_type": "refresh_token", "refresh_token": creds["QBO_REFRESH_TOKEN"]},
        timeout=30,
    )
    if r.status_code != 200:
        print(f"✗  Token refresh failed ({r.status_code}): {r.text[:300]}")
        sys.exit(1)
    body = r.json()
    new_rt = body.get("refresh_token")
    if new_rt and new_rt != creds["QBO_REFRESH_TOKEN"]:
        try:
            kc.put("QBO_REFRESH_TOKEN", new_rt)
        except kc.SecretsError:
            pass
    return body["access_token"], creds["QBO_COMPANY_ID"]


# ────────────────────────── api helpers ──────────────────────────

def _api_get(path: str, access: str, params: Optional[dict] = None) -> dict:
    """GET with patient retry on read/connect timeouts and transient QBO
    5xx/429 (incl. Intuit's 503 SystemFailureError, code 10000, which is a
    server-side blip — not our bug). 8 attempts, exponential backoff capped
    at 30s + small jitter → absorbs a QBO outage up to ~2–3 minutes before
    giving up. See [[reference_qbo_api_resilience]]."""
    import time as _time
    import random as _random
    p = dict(params or {})
    p["minorversion"] = MINOR_VERSION
    MAX_ATTEMPTS = 8

    def _sleep(attempt: int) -> None:
        _time.sleep(min(2 ** attempt, 30) + _random.uniform(0, 1.0))

    for attempt in range(MAX_ATTEMPTS):
        last = attempt == MAX_ATTEMPTS - 1
        try:
            r = requests.get(
                f"{API_BASE}{path}",
                headers={"Authorization": f"Bearer {access}", "Accept": "application/json"},
                params=p, timeout=120,
            )
        except (requests.exceptions.ReadTimeout,
                requests.exceptions.ConnectTimeout,
                requests.exceptions.ConnectionError) as e:
            if last:
                raise RuntimeError(f"{path} → network error after "
                                   f"{MAX_ATTEMPTS} tries: {e}")
            _sleep(attempt)
            continue
        if r.status_code == 200:
            return r.json()
        if r.status_code in (429, 500, 502, 503, 504) and not last:
            print(f"      QBO {r.status_code} (transient) — retry "
                  f"{attempt + 1}/{MAX_ATTEMPTS - 1}...")
            _sleep(attempt)
            continue
        raise RuntimeError(f"{path} → {r.status_code}: {r.text[:300]}")
    raise RuntimeError(f"{path} → unreachable")


def query_all(access: str, company_id: str, entity: str, where: str = "") -> List[dict]:
    # MAXRESULTS 1000 is QBO's max page size — fewer round-trips than 500.
    PAGE = 1000
    out: List[dict] = []
    start = 1
    while True:
        q = f"SELECT * FROM {entity}"
        if where:
            q += f" WHERE {where}"
        q += f" STARTPOSITION {start} MAXRESULTS {PAGE}"
        batch = (
            _api_get(f"/v3/company/{company_id}/query", access, {"query": q})
            .get("QueryResponse", {})
            .get(entity, [])
        )
        if not batch:
            break
        out.extend(batch)
        if len(batch) < PAGE:
            break
        start += PAGE
    return out


def report(access: str, company_id: str, name: str, params: Optional[dict] = None) -> dict:
    return _api_get(f"/v3/company/{company_id}/reports/{name}", access, params=params)


# ────────────────────────── customer lookup ──────────────────────────

def extract_proj(text: str) -> Optional[str]:
    if not text:
        return None
    m = PROJ_RE.search(str(text))
    return m.group(1).upper() if m else None


def build_project_customer_map(access: str, company_id: str) -> Dict[str, dict]:
    """
    Returns { 'MFD177': {id, name, parent, balance, ...}, ... }
    Strict project # match per [[feedback_ftw_separate_project]] — no family rollup.
    """
    customers = query_all(access, company_id, "Customer")
    by_proj: Dict[str, dict] = {}
    for c in customers:
        name = c.get("DisplayName") or c.get("CompanyName") or ""
        proj = extract_proj(name)
        if proj and proj not in by_proj:
            by_proj[proj] = {
                "id": c["Id"],
                "name": name,
                "fully_qualified_name": c.get("FullyQualifiedName", name),
                "balance": float(c.get("Balance", 0) or 0),
                "parent_id": (c.get("ParentRef") or {}).get("value"),
            }
    return by_proj


# ────────────────────────── P&L report ──────────────────────────

def fetch_project_pl(
    access: str,
    company_id: str,
    customer_id: str,
    start_date: str,
    end_date: str,
) -> dict:
    """Returns raw QBO ProfitAndLoss report filtered to one customer/project."""
    return report(access, company_id, "ProfitAndLoss", params={
        "start_date": start_date,
        "end_date": end_date,
        "accounting_method": "Accrual",
        "customer": customer_id,
    })


def _walk_pl_rows(report_data: dict) -> List[Tuple[str, Optional[float], int, str]]:
    """
    Flatten the nested QBO report row structure.
    Returns list of (label, amount, depth, kind) tuples in display order.
    kind ∈ {'section_header', 'section_total', 'data'}
    """
    rows: List[Tuple[str, Optional[float], int, str]] = []

    def walk(node: dict, depth: int) -> None:
        rtype = node.get("type")
        if rtype == "Section":
            header = node.get("Header", {}).get("ColData", [])
            if header:
                label = header[0].get("value", "")
                if label:
                    rows.append((label, None, depth, "section_header"))
            inner = (node.get("Rows") or {}).get("Row") or []
            for child in inner:
                walk(child, depth + 1)
            summary = node.get("Summary", {}).get("ColData", [])
            if summary and len(summary) > 1:
                label = summary[0].get("value", "") or ""
                amt_s = summary[-1].get("value", "") or "0"
                try:
                    amt = float(amt_s.replace(",", ""))
                except ValueError:
                    amt = 0.0
                rows.append((label, amt, depth, "section_total"))
        elif rtype == "Data":
            cols = node.get("ColData", [])
            if len(cols) >= 2:
                label = cols[0].get("value", "") or ""
                amt_s = cols[-1].get("value", "") or "0"
                try:
                    amt = float(amt_s.replace(",", "")) if amt_s else 0.0
                except ValueError:
                    amt = 0.0
                rows.append((label, amt, depth, "data"))
        else:
            inner = (node.get("Rows") or {}).get("Row") or []
            for child in inner:
                walk(child, depth)

    root = (report_data.get("Rows") or {}).get("Row") or []
    for n in root:
        walk(n, 0)
    return rows


def extract_pl_totals(report_data: dict) -> Dict[str, float]:
    """
    Walk the report and pull the canonical roll-up numbers we'll display
    at the top of Sheet 1 even if the body shows the full account tree.
    """
    out = {"income": 0.0, "cogs": 0.0, "gross_profit": 0.0,
           "expenses": 0.0, "net_ordinary_income": 0.0, "net_income": 0.0}
    rows = _walk_pl_rows(report_data)
    for label, amt, _depth, _kind in rows:
        if amt is None:
            continue
        l = label.lower().strip()
        if l in ("total income", "total revenue", "total sales"):
            out["income"] = amt
        elif l.startswith("total cost of goods") or l == "total cogs" or l.startswith("total job"):
            out["cogs"] = amt
        elif l == "gross profit":
            out["gross_profit"] = amt
        elif l in ("total expenses", "total expense", "total operating expenses"):
            out["expenses"] = amt
        elif l in ("net ordinary income", "net operating income"):
            out["net_ordinary_income"] = amt
        elif l == "net income":
            out["net_income"] = amt
    # Fallbacks
    if out["gross_profit"] == 0 and out["income"]:
        out["gross_profit"] = out["income"] - out["cogs"]
    if out["net_ordinary_income"] == 0:
        out["net_ordinary_income"] = out["gross_profit"] - out["expenses"]
    if out["net_income"] == 0:
        out["net_income"] = out["net_ordinary_income"]
    return out


# ────────────────────────── invoice pulls ──────────────────────────

def fetch_customer_invoices(access: str, company_id: str, customer_id: str) -> List[dict]:
    return query_all(
        access, company_id, "Invoice",
        where=f"CustomerRef = '{customer_id}'",
    )
