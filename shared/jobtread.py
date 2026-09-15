"""
jobtread.py - the ONE JobTread (Pave API) client. READ-ONLY.

Moved out of one-offs/rp_jobtread_coverage.py (2026-09-15) the moment a second tool - the
ledger's RP review page - needed it (tools never import tools). The one-offs import these
names back, byte-compatible.

Auth: JT_GRANT_KEY lives in the shared Keychain blob (shared/qbo_vault) - one Touch ID per
run. Nothing here writes to JobTread.
"""
from __future__ import annotations

import json
import os
import urllib.request
from typing import Dict, Iterable, List, Optional, Tuple

ORG_ID = os.getenv("JT_ORG_ID", "22PFAfqHLF3a")
API_URL = "https://api.jobtread.com/pave"


def pave(key: str, query: dict) -> dict:
    """POST one Pave query with the grant key folded in; returns the decoded JSON."""
    body = json.dumps({"query": {"$": {"grantKey": key}, **query}}).encode()
    req = urllib.request.Request(API_URL, data=body,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as r:       # noqa: S310 - fixed https host
        return json.loads(r.read().decode())


def grant_key() -> Optional[str]:
    """JT_GRANT_KEY from the shared vault, or None when absent / locked."""
    try:
        from . import qbo_vault                                # sibling module
    except ImportError:                                        # pragma: no cover - script path
        import qbo_vault                                       # type: ignore
    try:
        return qbo_vault.get("JT_GRANT_KEY") or None
    except Exception:                                          # noqa: BLE001 - vault locked / absent
        return None


def approved_proposals(numbers: Iterable[str], key: Optional[str] = None,
                       log=None) -> Dict[str, List[Tuple[float, float, str]]]:
    """{job# -> [(price, cost, created yyyy-mm-dd)]} of APPROVED customerOrder documents,
    one query per job number. {} when there is no key. A failing number is logged and
    skipped, never raised - a report must not die on one bad job."""
    key = key or grant_key()
    out: Dict[str, List[Tuple[float, float, str]]] = {}
    if not key:
        return out
    for n in sorted(set(numbers)):
        try:
            r = pave(key, {"organization": {"$": {"id": ORG_ID}, "jobs": {
                "$": {"size": 3, "where": {"and": [["number", "=", n]]}},
                "nodes": {"documents": {"$": {"size": 50}, "nodes": {
                    "type": {}, "status": {}, "price": {}, "cost": {}, "createdAt": {}}}}}}})
        except Exception as e:                                 # noqa: BLE001
            if log:
                log(f"    JobTread {n}: {type(e).__name__}")
            continue
        docs = [d for j in r["organization"]["jobs"]["nodes"] for d in j["documents"]["nodes"]
                if d.get("type") == "customerOrder" and d.get("status") == "approved"]
        if docs:
            out[n] = [(float(d.get("price") or 0), float(d.get("cost") or 0),
                       str(d.get("createdAt") or "")[:10]) for d in docs]
    return out
