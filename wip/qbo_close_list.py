#!/usr/bin/env python3
"""
qbo_close_list.py — Diff Ted's "really open" project list vs QBO's active customers.

Outputs three sections:

  CLOSE THESE        — active in QBO, NOT in Ted's open list. Mark inactive in QBO.
                       MFD projects are EXCLUDED from this list — Ted handles MFD
                       closures manually. They appear in MFD ACTIVE instead.
  MFD ACTIVE         — every active MFD customer in QBO (manual review).
  KEEP ACTIVE        — exact matches between Ted's list and QBO.
  QUESTIONS          — in Ted's open list but NOT found as active in QBO.
                       Either inactive (needs reactivation) or never project-tagged,
                       OR base name like "CP861" when QBO only has "CP861-7E" / "CP861-BP".

Matching is **strict** — exact project # only. -FTW is a separate project,
not a variant of the base. So if you want both `RP7186` and `RP7186-FTW` open,
both must be in your list explicitly.

Run:
    python3 wip/qbo_close_list.py
    python3 wip/qbo_close_list.py --strict
    python3 wip/qbo_close_list.py --json out.json
"""
from __future__ import annotations

import argparse
import base64
import json
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_vault as kc

API_BASE = "https://quickbooks.api.intuit.com"

_PROJ_RE = re.compile(
    r"\b(RP\d{4}(?:-[A-Za-z]{2,6})?|CP\d{3,4}(?:-[A-Za-z0-9]{1,6})?|MFD\d{3,4})\b",
    re.IGNORECASE,
)


# ──────────────────────────────────────────────────────────────────────────────
# Ted's open list — the source of truth.
# ──────────────────────────────────────────────────────────────────────────────
TED_OPEN_PROJECTS = """
MFD133
MFD160
MFD166
MFD172
MFD177
MFD182
MFD183
MFD186
MFD192
MFD228
MFD231
MFD281
MFD295
MFD325
CP585
CP672
CP745
CP765
CP783
CP790
CP800
CP803
CP861
CP885
CP961
RP7315-FTW
RP7005-FTW
RP7278-FTW
RP7242-FTW
RP7207-FTW
RP7258-FTW
RP7333-FTW
RP7260-FTW
RP7261-FTW
RP7365-FTW
RP7366-FTW
RP7367-FTW
RP7370-FTW
RP7371-FTW
RP7374-FTW
RP7376-FTW
RP7256-FTW
RP7234-FTW
RP6586
RP6586-FTW
RP7060-FTW
RP7332-FTW
RP7272-FTW
RP7020-FTW
RP7425-FTW
RP7426-FTW
RP7427-FTW
RP7428-FTW
RP7431-FTW
RP7432-FTW
RP7433-FTW
RP7441-FTW
RP7434-FTW
RP7447-FTW
RP7450-FTW
RP7264-FTW
RP7444-FTW
RP7408-FTW
RP6764
RP7084-FTW
RP7327
RP7404
RP7471
RP7471-FTW
RP7484
RP7484-FTW
RP7485
RP7485-FTW
RP7363
RP7137
RP7379
RP7379-FTW
RP7228
RP7228-FTW
RP7417
RP7417-FTW
RP7267
RP7267-FTW
RP7455
RP7455-FTW
RP7407
RP7435
RP7387
RP7387-FTW
RP7388
RP7388-FTW
RP7385
RP7385-FTW
RP7420
RP6533
RP7461
RP7460
RP7456
RP7445
RP7466
RP6721
RP6721-FTW
RP7448
RP7359
RP7358
RP7437
RP7182
RP7186
"""


def parse_open_list() -> Set[str]:
    return {ln.strip().upper() for ln in TED_OPEN_PROJECTS.splitlines() if ln.strip()}


def base_num(p: str) -> str:
    return p.split("-")[0] if "-" in p else p


def extract_project(text: str) -> Optional[str]:
    if not text:
        return None
    m = _PROJ_RE.search(str(text))
    return m.group(1).upper() if m else None


# ──────────────────────────────────────────────────────────────────────────────
# QBO auth + customer fetch
# ──────────────────────────────────────────────────────────────────────────────
def load_credentials() -> Tuple[str, str]:
    if not kc.has_credentials():
        print("[ERR] No QBO credentials in Keychain")
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


def query_all(access: str, company_id: str, entity: str, where: str = "") -> List[dict]:
    out: List[dict] = []
    start = 1
    while True:
        q = f"SELECT * FROM {entity}"
        if where:
            q += f" WHERE {where}"
        q += f" STARTPOSITION {start} MAXRESULTS 500"
        r = requests.get(
            f"{API_BASE}/v3/company/{company_id}/query",
            headers={"Authorization": f"Bearer {access}", "Accept": "application/json"},
            params={"query": q, "minorversion": "70"},
            timeout=60,
        )
        r.raise_for_status()
        batch = r.json().get("QueryResponse", {}).get(entity, [])
        if not batch:
            break
        out.extend(batch)
        if len(batch) < 500:
            break
        start += 500
    return out


# ──────────────────────────────────────────────────────────────────────────────
# Diff
# ──────────────────────────────────────────────────────────────────────────────
def main() -> int:
    ap = argparse.ArgumentParser(description="Diff Ted's open list vs QBO active (strict matching)")
    ap.add_argument("--json", default=None, help="Optional JSON output path")
    args = ap.parse_args()

    open_list = parse_open_list()
    print(f"[info] Ted's open list: {len(open_list)} projects (strict matching)")

    print("[info] Authenticating to QBO (Touch ID)...")
    access, company_id = load_credentials()

    print("[info] Pulling ACTIVE customers...")
    active_customers = query_all(access, company_id, "Customer", where="Active = true")
    print(f"[info] {len(active_customers)} active customers in QBO")

    print("[info] Pulling INACTIVE customers...")
    inactive_customers = query_all(access, company_id, "Customer", where="Active = false")
    print(f"[info] {len(inactive_customers)} inactive customers in QBO")

    # Build maps: project# → list of (customer_id, display_name, active_flag)
    active_by_proj: Dict[str, List[Tuple[str, str]]] = defaultdict(list)
    for c in active_customers:
        name = c.get("DisplayName") or ""
        proj = extract_project(name)
        if proj:
            active_by_proj[proj].append((c["Id"], name))

    inactive_by_proj: Dict[str, List[Tuple[str, str]]] = defaultdict(list)
    for c in inactive_customers:
        name = c.get("DisplayName") or ""
        proj = extract_project(name)
        if proj:
            inactive_by_proj[proj].append((c["Id"], name))

    print(f"[info] active project-tagged: {len(active_by_proj)} unique project numbers")
    print(f"[info] inactive project-tagged: {len(inactive_by_proj)} unique project numbers")

    # Strict matching — exact project # only. -FTW is a separate project.
    def is_open(qbo_proj: str) -> bool:
        return qbo_proj in open_list

    close_these: List[Tuple[str, str, str]] = []   # (proj, cust_id, name) — to auto-close (NO MFDs)
    keep_active: List[Tuple[str, str, str]] = []
    mfd_active: List[Tuple[str, str, str]] = []    # ALL active MFDs — Ted handles manually

    for proj, customers in sorted(active_by_proj.items()):
        is_mfd = proj.upper().startswith("MFD")
        for cust_id, name in customers:
            if is_mfd:
                # MFDs always go to manual review, never auto-close.
                mfd_active.append((proj, cust_id, name))
                continue
            if is_open(proj):
                keep_active.append((proj, cust_id, name))
            else:
                close_these.append((proj, cust_id, name))

    # In Ted's list but not active in QBO (strict only — exact match)
    qbo_active_set = set(active_by_proj.keys())
    qbo_inactive_set = set(inactive_by_proj.keys())
    questions: List[Tuple[str, str]] = []
    for proj in sorted(open_list):
        in_active = proj in qbo_active_set
        in_inactive = proj in qbo_inactive_set
        if not in_active:
            status = "INACTIVE in QBO (reactivate?)" if in_inactive else "NOT FOUND in QBO"
            questions.append((proj, status))

    # ── Output ────────────────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    print(f"  CLOSE THESE — {len(close_these)} active QBO customers to auto-close (NO MFDs)")
    print("=" * 70)
    print(f"  {'Project #':<16} {'Customer ID':<10}  Display Name")
    print(f"  {'-'*16} {'-'*10}  {'-'*40}")
    for proj, cust_id, name in close_these:
        print(f"  {proj:<16} {cust_id:<10}  {name}")

    print("\n" + "=" * 70)
    print(f"  MFD ACTIVE — {len(mfd_active)} active MFD customers (your manual review)")
    print("=" * 70)
    print(f"  {'Project #':<16} {'Customer ID':<10}  Display Name")
    print(f"  {'-'*16} {'-'*10}  {'-'*40}")
    for proj, cust_id, name in mfd_active:
        print(f"  {proj:<16} {cust_id:<10}  {name}")

    print("\n" + "=" * 70)
    print(f"  KEEP ACTIVE — {len(keep_active)} customers matched (exact) to your open list")
    print("=" * 70)
    print(f"  {'Project #':<16} {'Customer ID':<10}  Display Name")
    print(f"  {'-'*16} {'-'*10}  {'-'*40}")
    for proj, cust_id, name in keep_active:
        print(f"  {proj:<16} {cust_id:<10}  {name}")

    print("\n" + "=" * 70)
    print(f"  QUESTIONS — {len(questions)} projects in your list NOT exact-matched in QBO")
    print("=" * 70)
    print(f"  (e.g. 'CP861' here means QBO has CP861-7E and CP861-BP — list those instead)")
    for proj, status in questions:
        print(f"  {proj:<16}  {status}")

    print("\n" + "=" * 70)
    print(f"  SUMMARY")
    print("=" * 70)
    print(f"  Ted's open list:     {len(open_list)}")
    print(f"  QBO active:          {len(active_customers)}")
    print(f"  QBO active w/ proj#: {sum(len(v) for v in active_by_proj.values())}")
    print(f"  → Auto-close (no MFD): {len(close_these)}")
    print(f"  → MFD active (manual): {len(mfd_active)}")
    print(f"  → Keep active:         {len(keep_active)}")
    print(f"  → Questions:           {len(questions)}")

    if args.json:
        Path(args.json).write_text(json.dumps({
            "ted_open_list": sorted(open_list),
            "matching_mode": "strict",
            "mfd_handling": "excluded from auto-close (manual review)",
            "close_these": [{"project": p, "id": cid, "name": n} for p, cid, n in close_these],
            "mfd_active": [{"project": p, "id": cid, "name": n} for p, cid, n in mfd_active],
            "keep_active": [{"project": p, "id": cid, "name": n} for p, cid, n in keep_active],
            "questions": [{"project": p, "status": s} for p, s in questions],
            "summary": {
                "ted_open": len(open_list),
                "qbo_active": len(active_customers),
                "close_count": len(close_these),
                "mfd_active_count": len(mfd_active),
                "keep_count": len(keep_active),
                "questions_count": len(questions),
            },
        }, indent=2))
        print(f"\n  JSON output: {args.json}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
