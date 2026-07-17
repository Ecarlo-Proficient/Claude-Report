#!/usr/bin/env python3
"""
jobtread_probe.py — can we read project budgets out of JobTread? (the user
2026-07-17: QBO's per-project budgets are locked away from the public API;
JobTread — where the takeoff's 'JobTread Cost Gral' numbers are entered —
has a real public API. This probe answers whether it can be our ETC source.)

READ-ONLY against JobTread. Auth is a grant key in its OWN Keychain blob
(service 'automation-jobtread' — same isolation pattern as QBO/Notion/Teams;
a bad key here can never touch the other vaults).

One-time setup (the key is typed into a hidden prompt, never chat/argv):
  In JobTread: Settings → Integrations → JobTread API → New Grant Key
  (read access is enough), then:
      python3 jobtread_probe.py --setup

Probe (discovery — prints what the API exposes, step by step):
  python3 jobtread_probe.py                 # who am I + org + a few jobs
  python3 jobtread_probe.py --job RP7538    # budget/cost data for one job
"""
from __future__ import annotations

import argparse
import base64
import getpass
import json
import os
import subprocess
import sys
import urllib.request

SERVICE = "automation-jobtread"
LABEL = "credentials"
ACCOUNT = os.environ.get("USER") or "user"
API_URL = "https://api.jobtread.com/pave"


def _sec(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(["/usr/bin/security", *args],
                          capture_output=True, text=True)


def store_key() -> None:
    key = getpass.getpass("JobTread grant key (hidden): ").strip()
    if not key:
        print("empty — nothing stored")
        sys.exit(1)
    blob = base64.b64encode(json.dumps({"JT_GRANT_KEY": key}).encode()).decode()
    _sec("delete-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL)
    r = _sec("add-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL,
             "-w", blob, "-T", "")
    if r.returncode != 0:
        print("keychain write failed:", r.stderr.strip())
        sys.exit(1)
    print("✓ stored in Keychain (service automation-jobtread)")


def load_key() -> str:
    r = _sec("find-generic-password", "-a", ACCOUNT, "-s", SERVICE,
             "-l", LABEL, "-w")
    if r.returncode != 0:
        print("No JobTread key stored — run:  python3 jobtread_probe.py --setup")
        sys.exit(1)
    return json.loads(base64.b64decode(r.stdout.strip()))["JT_GRANT_KEY"]


def pave(key: str, query: dict) -> dict:
    """POST one Pave query. The request shape mirrors the response shape."""
    body = json.dumps({"query": {"$": {"grantKey": key}, **query}}).encode()
    req = urllib.request.Request(
        API_URL, data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode())


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--setup", action="store_true", help="store the grant key")
    ap.add_argument("--job", help="search this job # and dump its cost data")
    args = ap.parse_args()
    if args.setup:
        store_key()
        return 0
    key = load_key()

    print("\n① who am I")
    r = pave(key, {"currentGrant": {"id": {}, "name": {}}})
    print(json.dumps(r.get("currentGrant"), indent=2)[:400])

    print("\n② organizations on this grant")
    r = pave(key, {"currentGrant": {"user": {"memberships": {"nodes": {
        "organization": {"id": {}, "name": {}}}}}}})
    orgs = [m["organization"] for m in
            (r.get("currentGrant", {}).get("user", {})
              .get("memberships", {}).get("nodes", []) or [])]
    print(json.dumps(orgs, indent=2)[:600])
    if not orgs:
        print("  (no orgs — inspect the raw reply above; schema may differ)")
        return 1
    org_id = orgs[0]["id"]

    if not args.job:
        print("\n③ first jobs in the org (name + number fields)")
        r = pave(key, {"organization": {"$": {"id": org_id}, "jobs": {
            "$": {"size": 10},
            "nodes": {"id": {}, "name": {}, "number": {}, "createdAt": {}}}}})
        print(json.dumps(r.get("organization", {}).get("jobs"), indent=2)[:1500])
        print("\nNext: python3 jobtread_probe.py --job RP####")
        return 0

    print(f"\n③ job search: {args.job}")
    r = pave(key, {"organization": {"$": {"id": org_id}, "jobs": {
        "$": {"size": 5, "where": {"and": [["name", "like", f"%{args.job}%"]]}},
        "nodes": {"id": {}, "name": {}, "number": {}}}}})
    jobs = r.get("organization", {}).get("jobs", {}).get("nodes", []) or []
    print(json.dumps(jobs, indent=2)[:800])
    if not jobs:
        print("  no match — try part of the address instead of the RP#")
        return 1
    jid = jobs[0]["id"]

    print("\n④ cost/budget data on that job — trying the documented shapes")
    for label, q in (
        ("costItems", {"job": {"$": {"id": jid}, "costItems": {
            "$": {"size": 50},
            "nodes": {"name": {}, "costCode": {"name": {}},
                      "unitCost": {}, "quantity": {}, "total": {}}}}}),
        ("estimates/documents", {"job": {"$": {"id": jid}, "documents": {
            "$": {"size": 10},
            "nodes": {"id": {}, "name": {}, "type": {}, "total": {}}}}}),
        ("budget rollup", {"job": {"$": {"id": jid},
                                   "estimatedCostTotal": {},
                                   "actualCostTotal": {},
                                   "estimatedPriceTotal": {}}}),
    ):
        try:
            r = pave(key, q)
            print(f"\n  ── {label}:")
            print(json.dumps(r.get("job"), indent=2)[:1200])
        except Exception as e:
            print(f"\n  ── {label}: ✗ {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
