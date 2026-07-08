#!/usr/bin/env python3
"""
qbo_bulk_close.py — Mark a list of QBO customers Inactive via the API.

INPUT:
    wip/qbo_close_list.json  (produced by qbo_close_list.py)

DEFAULT BEHAVIOR: dry-run. Lists what would be closed, writes nothing.

To actually mark customers inactive in QBO:
    python3 wip/qbo_bulk_close.py --execute
    (then type CLOSE-1227 at the prompt to confirm)

To process only the first N customers for a small test batch:
    python3 wip/qbo_bulk_close.py --execute --limit 10

SAFETY:
- Default is dry-run. --execute required for any writes.
- Typed-confirmation prompt with exact count must match.
- QBO refuses to deactivate customers with open balance — those are
  logged as 'skipped: open balance' and the script continues.
- Throttled to 5 requests/sec to stay well under QBO's rate limit.
- Every action (success, skip, error) is logged to
  wip/qbo_bulk_close_results.json for audit/undo.
"""
from __future__ import annotations

import argparse
import base64
import datetime as dt
import json
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import qbo_vault as kc

API_BASE = "https://quickbooks.api.intuit.com"
RATE_LIMIT_DELAY_S = 0.2  # 5 req/sec — well under QBO's 500/min/realm limit.


def load_credentials() -> Tuple[str, str]:
    if not kc.has_credentials():
        print("[ERR] No QBO credentials in Keychain.")
        sys.exit(1)
    creds = kc.get_all()
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
    r.raise_for_status()
    body = r.json()
    new_rt = body.get("refresh_token")
    if new_rt and new_rt != creds["QBO_REFRESH_TOKEN"]:
        kc.put("QBO_REFRESH_TOKEN", new_rt)
    return body["access_token"], creds["QBO_COMPANY_ID"]


def _api_get(access: str, company_id: str, path: str, params: Optional[dict] = None) -> dict:
    p = dict(params or {})
    p["minorversion"] = "70"
    r = requests.get(
        f"{API_BASE}{path}",
        headers={"Authorization": f"Bearer {access}", "Accept": "application/json"},
        params=p,
        timeout=30,
    )
    r.raise_for_status()
    return r.json()


def fetch_customer(access: str, company_id: str, cust_id: str) -> Optional[dict]:
    """Fetch a customer record (we need the current SyncToken for any update)."""
    try:
        data = _api_get(access, company_id, f"/v3/company/{company_id}/customer/{cust_id}")
        return data.get("Customer")
    except requests.HTTPError as e:
        return None


def mark_inactive(access: str, company_id: str, cust_id: str, sync_token: str, display_name: str) -> Tuple[str, str]:
    """
    Sparse-update a customer to Active=false.
    Returns (status, detail) where status in {ok, skipped, error}.
    """
    body = {
        "sparse": True,
        "Id": cust_id,
        "SyncToken": sync_token,
        "Active": False,
    }
    r = requests.post(
        f"{API_BASE}/v3/company/{company_id}/customer",
        headers={
            "Authorization": f"Bearer {access}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        params={"minorversion": "70"},
        json=body,
        timeout=30,
    )
    if r.status_code == 200:
        return ("ok", "marked inactive")

    # QBO refuses deactivation if there's an open balance, an active sub-customer, etc.
    # Capture the reason and keep going.
    try:
        err = r.json().get("Fault", {}).get("Error", [{}])[0]
        msg = err.get("Message", "") or ""
        detail = err.get("Detail", "") or ""
        reason = f"{msg}: {detail}".strip(": ").strip()
    except Exception:
        reason = f"HTTP {r.status_code}: {r.text[:200]}"
    return ("error", reason)


def main() -> int:
    ap = argparse.ArgumentParser(description="Bulk mark QBO customers inactive")
    ap.add_argument("--input", default="wip/qbo_close_list.json",
                    help="JSON output from qbo_close_list.py")
    ap.add_argument("--execute", action="store_true",
                    help="Actually mark customers inactive. Default: dry-run only.")
    ap.add_argument("--limit", type=int, default=None,
                    help="Process only the first N customers (for testing).")
    ap.add_argument("--output", default="wip/qbo_bulk_close_results.json",
                    help="Where to write the audit log.")
    args = ap.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"[ERR] Input not found: {in_path}")
        print("      Run qbo_close_list.py first to generate it.")
        return 1

    data = json.loads(in_path.read_text())
    close_list = data.get("close_these", [])
    if args.limit:
        close_list = close_list[: args.limit]

    n = len(close_list)
    print("=" * 70)
    print(f"  QBO Bulk Customer Closer")
    print("=" * 70)
    print(f"  Input file       : {in_path}")
    print(f"  Customers to close: {n}")
    print(f"  Mode             : {'EXECUTE (writes to QBO)' if args.execute else 'DRY-RUN (no writes)'}")
    if args.limit:
        print(f"  Limit            : first {args.limit} customers")

    if not args.execute:
        print()
        print("  Preview of first 20:")
        for entry in close_list[:20]:
            print(f"    {entry['project']:<16} {entry['id']:<10} {entry['name']}")
        if n > 20:
            print(f"    ... and {n - 20} more")
        print()
        print("  Re-run with --execute to actually close these in QBO.")
        return 0

    # ── Confirmation gate ────────────────────────────────────────────────────
    expected = f"CLOSE-{n}"
    print()
    print(f"  ⚠  This will mark {n} QBO customers as INACTIVE.")
    print(f"     Customers with open balance will be skipped, not failed.")
    print(f"     Action is reversible (Make Active again in QBO UI).")
    print()
    typed = input(f"  Type {expected} exactly to proceed (anything else aborts): ").strip()
    if typed != expected:
        print("  [aborted]")
        return 1

    # ── Auth ─────────────────────────────────────────────────────────────────
    print("\n  Authenticating to QBO (Touch ID)...")
    access, company_id = load_credentials()
    print(f"  ok — company {company_id}")

    # ── Process ──────────────────────────────────────────────────────────────
    results = []
    counts = {"ok": 0, "skipped_no_token": 0, "skipped_already_inactive": 0, "error": 0}

    print(f"\n  Processing {n} customers (~{n * RATE_LIMIT_DELAY_S / 60:.1f} min at {1/RATE_LIMIT_DELAY_S:.0f} req/sec)…")
    started = time.time()

    for i, entry in enumerate(close_list, 1):
        cust_id = entry["id"]
        proj = entry["project"]
        name = entry["name"]

        # Refetch to get current SyncToken — couple seconds upfront beats a 400 on stale token mid-run.
        cust = fetch_customer(access, company_id, cust_id)
        if not cust:
            counts["skipped_no_token"] += 1
            results.append({"id": cust_id, "project": proj, "name": name, "status": "skipped",
                            "reason": "could not fetch customer (deleted/permission?)"})
            print(f"  [{i:4d}/{n}] {proj:<16} {cust_id:<8} SKIP  (could not fetch)")
            time.sleep(RATE_LIMIT_DELAY_S)
            continue

        if cust.get("Active") is False:
            counts["skipped_already_inactive"] += 1
            results.append({"id": cust_id, "project": proj, "name": name, "status": "skipped",
                            "reason": "already inactive"})
            print(f"  [{i:4d}/{n}] {proj:<16} {cust_id:<8} SKIP  (already inactive)")
            continue

        sync_token = cust.get("SyncToken")
        status, detail = mark_inactive(access, company_id, cust_id, sync_token, name)
        counts[status] = counts.get(status, 0) + 1
        results.append({"id": cust_id, "project": proj, "name": name, "status": status, "reason": detail})

        marker = "OK" if status == "ok" else "ERR" if status == "error" else "SKIP"
        print(f"  [{i:4d}/{n}] {proj:<16} {cust_id:<8} {marker:<5} {detail}")

        time.sleep(RATE_LIMIT_DELAY_S)

    elapsed = time.time() - started

    # ── Save audit log ───────────────────────────────────────────────────────
    out = {
        "run_at": dt.datetime.utcnow().isoformat() + "Z",
        "input_file": str(in_path),
        "total_customers": n,
        "elapsed_seconds": round(elapsed, 1),
        "counts": counts,
        "results": results,
    }
    Path(args.output).write_text(json.dumps(out, indent=2))

    # ── Summary ──────────────────────────────────────────────────────────────
    print()
    print("=" * 70)
    print("  SUMMARY")
    print("=" * 70)
    print(f"  Marked inactive    : {counts.get('ok', 0)}")
    print(f"  Already inactive   : {counts.get('skipped_already_inactive', 0)}")
    print(f"  Could not fetch    : {counts.get('skipped_no_token', 0)}")
    print(f"  Errors             : {counts.get('error', 0)}")
    print(f"  Elapsed            : {elapsed:.1f}s")
    print(f"  Audit log          : {args.output}")

    if counts.get("error", 0):
        print()
        print("  Errors (first 10):")
        for r in results:
            if r["status"] == "error":
                print(f"    {r['project']:<16} {r['id']:<8} {r['reason']}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
