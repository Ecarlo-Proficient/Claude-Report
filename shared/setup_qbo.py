#!/usr/bin/env python3
"""
setup_qbo.py — Interactive QBO credential setup, then auth test.

WHY ONE TOUCH ID IS ENOUGH
  All five QBO keys live in ONE Keychain blob (see qbo_vault.py).
  One Touch ID prompt decrypts the whole blob for this process.
  Setup writes the blob in one operation; --test reads it in one
  operation; qbo_export.py reads it in one operation. No more
  per-key approval prompts.

FEATURES
  - Paste values are HIDDEN (getpass).
  - After each paste: masked preview + length + shape check,
    then Y/n confirm before the blob is updated.
  - PER-KEY ROTATION — bad paste in one key doesn't lose the others.
  - Auth test runs automatically after setup or rotate. On failure
    the test tells you EXACTLY which --rotate to run.

USAGE
  python3 setup_qbo.py              # prompt only missing keys, then --test
  python3 setup_qbo.py --rotate KEY # rotate one key (e.g. QBO_CLIENT_ID)
  python3 setup_qbo.py --status     # what's stored (one Touch ID)
  python3 setup_qbo.py --test       # auth test only (one Touch ID)
  python3 setup_qbo.py --all        # re-prompt every key
  python3 setup_qbo.py --purge      # wipe the blob
"""
from __future__ import annotations

import argparse
import base64
import getpass
import sys
from dataclasses import dataclass
from typing import Callable, Dict, Optional

import requests

import qbo_vault as kc


# ──────────────────────────  key specs  ──────────────────────────

@dataclass
class KeySpec:
    name: str
    help: str
    example: str
    min_len: int
    max_len: int
    shape: Optional[Callable[[str], bool]] = None


def _is_digits(v: str) -> bool:
    return v.isdigit()


# Production only. Sandbox support has been intentionally removed to eliminate
# the class of "wrong tab" mistakes that produce `invalid_client` errors.
SPECS = [
    KeySpec(
        "QBO_CLIENT_ID",
        "Intuit Developer > your app > Keys & OAuth > **Production Keys** tab > CLIENT ID",
        "(~40 chars, alphanumeric)",
        30, 120, None,
    ),
    KeySpec(
        "QBO_CLIENT_SECRET",
        "Intuit Developer > same page > **Production Keys** tab > CLIENT SECRET",
        "(~40 chars). Must come from the SAME 'Production Keys' tab as Client ID.",
        30, 120, None,
    ),
    KeySpec(
        "QBO_COMPANY_ID",
        "QBO > gear icon > Account and Settings > Billing & Subscription > Company ID",
        "1234567890123456789 (digits only, ~19 digits)",
        8, 25, _is_digits,
    ),
    KeySpec(
        "QBO_REFRESH_TOKEN",
        "The refresh token captured when you authorized your production app",
        "opaque string — format varies",
        20, 500, None,
    ),
    # Non-QBO keys in the same blob — the vault is the company key LIBRARY
    # (the user 2026-07-17): one place, one Touch ID. Blank = skip on setup.
    KeySpec(
        "JT_GRANT_KEY",
        "JobTread > Settings > Integrations > JobTread API > New Grant Key (read access)",
        "opaque alphanumeric string",
        10, 200, None,
    ),
]


# ──────────────────────────  prompting  ──────────────────────────

def _mask(v: str) -> str:
    return v if len(v) <= 12 else f"{v[:6]}…{v[-4:]}"


def prompt_one(spec: KeySpec, currently_stored: bool) -> Optional[str]:
    """Prompt for one value. Returns the string to store, or None to skip."""
    print()
    print(f"━━━  {spec.name}  ━━━")
    print(f"     Where: {spec.help}")
    print(f"     Shape: {spec.example}")
    if currently_stored:
        print(f"     (already stored — paste to overwrite, blank = keep existing)")

    raw = getpass.getpass("     Paste value (hidden): ").strip()
    if not raw:
        print("     ↷ kept existing")
        return None

    print(f"     Captured: {_mask(raw)}   (length {len(raw)})")

    problems = []
    if len(raw) < spec.min_len or len(raw) > spec.max_len:
        problems.append(f"length {len(raw)} outside expected {spec.min_len}-{spec.max_len}")
    if spec.shape and not spec.shape(raw):
        problems.append("does not match expected pattern")

    if problems:
        print(f"     ⚠ {'; '.join(problems)}")
        if input("     Save anyway? [y/N] ").strip().lower() != "y":
            print("     ↷ not saved")
            return None
    else:
        if input("     Save? [Y/n] ").strip().lower() == "n":
            print("     ↷ not saved")
            return None

    return raw


# ──────────────────────────  auth test  ──────────────────────────

# Production-only. Hardcoded — no env selector, no ambiguity.
API_BASE = "https://quickbooks.api.intuit.com"


def run_test() -> int:
    """Verify stored credentials. One Touch ID prompt for the whole check."""
    print()
    print("━" * 60)
    print("  QBO Auth Test")
    print("━" * 60)

    if not kc.has_credentials():
        print("  ✗ no credentials stored.")
        print("    fix:  python3 setup_qbo.py")
        return 1

    try:
        creds = kc.get_all()  # ← single Touch ID prompt
    except kc.SecretsError as e:
        print(f"  ✗ could not read blob: {e}")
        return 1

    missing = [s.name for s in SPECS if s.name not in creds or not creds[s.name]]
    if missing:
        print(f"  ✗ missing keys in blob: {', '.join(missing)}")
        print(f"    fix:  python3 setup_qbo.py")
        return 1
    print(f"  ✓ all {len(SPECS)} keys present in blob")

    cid = creds["QBO_CLIENT_ID"]
    sec = creds["QBO_CLIENT_SECRET"]
    rt = creds["QBO_REFRESH_TOKEN"]
    cmp_id = creds["QBO_COMPANY_ID"]

    print(f"  → api base:   {API_BASE}  (production)")
    print(f"  → company id: {cmp_id}")
    print(f"  → attempting token refresh...")

    basic = base64.b64encode(f"{cid}:{sec}".encode()).decode()
    try:
        r = requests.post(
            "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer",
            headers={
                "Authorization": f"Basic {basic}",
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
            },
            data={"grant_type": "refresh_token", "refresh_token": rt},
            timeout=30,
        )
    except requests.RequestException as e:
        print(f"  ✗ network error: {e}")
        return 1

    if r.status_code != 200:
        body = (r.text or "")[:300]
        print(f"  ✗ refresh failed  status={r.status_code}")
        print(f"    body: {body}")
        print()
        _diagnose_refresh_failure(r.status_code, body)
        return 1

    data = r.json()
    access = data["access_token"]
    new_rt = data.get("refresh_token")
    print(f"  ✓ token refresh ok  (access expires in {data.get('expires_in')}s)")

    if new_rt and new_rt != rt:
        try:
            kc.put("QBO_REFRESH_TOKEN", new_rt)
            print(f"  ✓ refresh token rotated — new one stored")
        except kc.SecretsError as e:
            print(f"  ⚠ new refresh token not stored: {e}")

    # Company probe confirms Company ID matches this app's authorized realm
    print(f"  → probing company info...")
    info = requests.get(
        f"{API_BASE}/v3/company/{cmp_id}/companyinfo/{cmp_id}",
        headers={"Authorization": f"Bearer {access}", "Accept": "application/json"},
        params={"minorversion": "70"},
        timeout=30,
    )
    if info.status_code != 200:
        body = (info.text or "")[:300]
        print(f"  ✗ company probe failed  status={info.status_code}")
        print(f"    body: {body}")
        print()
        _diagnose_company_failure(info.status_code, body)
        return 1

    company = info.json().get("CompanyInfo", {})
    name = company.get("CompanyName") or company.get("LegalName") or "(unknown)"
    print(f"  ✓ company: {name}")
    print()
    print("━" * 60)
    print("  ✓  AUTH OK — run:  python3 qbo_export.py")
    print("━" * 60)
    return 0


def _diagnose_refresh_failure(status: int, body: str) -> None:
    lo = body.lower()
    if "invalid_client" in lo:
        print("  ── invalid_client = Intuit rejected Client ID + Secret.")
        print("     Two possible causes (since we're production-only):")
        print("       (1) Client ID came from the Development Keys tab instead")
        print("           of Production Keys. They must BOTH come from the")
        print("           Production Keys tab, on the same page load.")
        print("       (2) The refresh token was issued under a DIFFERENT app's")
        print("           Client ID — mismatched app credentials.")
        print()
        print("     Fix: on intuit.com > your app > Keys & OAuth > PRODUCTION")
        print("     Keys tab, copy BOTH values from that single view, then:")
        print("       python3 setup_qbo.py --rotate QBO_CLIENT_ID")
        print("       python3 setup_qbo.py --rotate QBO_CLIENT_SECRET")
        return
    if "invalid_grant" in lo:
        print("  ── invalid_grant = refresh token expired or revoked.")
        print("     QBO refresh tokens last ~100 days from last use. Re-authorize")
        print("     your production app to get a fresh one, then:")
        print("       python3 setup_qbo.py --rotate QBO_REFRESH_TOKEN")
        return
    print(f"  ── unknown {status} from token endpoint. Full body above.")


def _diagnose_company_failure(status: int, body: str) -> None:
    if status in (401, 403):
        print("  ── Auth worked but this app doesn't have access to that company.")
        print("     Company ID doesn't match the realm that authorized your app.")
        print("     Fix: verify Company ID in QBO > gear > Account and Settings >")
        print("     Billing & Subscription, then:")
        print("       python3 setup_qbo.py --rotate QBO_COMPANY_ID")
        return
    if status == 404:
        print("  ── Company not found. Verify Company ID, then:")
        print("       python3 setup_qbo.py --rotate QBO_COMPANY_ID")
        return
    print(f"  ── unexpected {status}. Full body is above.")


# ──────────────────────────  commands  ──────────────────────────

def run_setup(all_keys: bool) -> int:
    print("━" * 60)
    print("  QBO Credential Setup")
    print("━" * 60)
    print("  Paste is hidden. After each paste you'll see a masked preview")
    print("  + length + shape check, then confirm before the blob is updated.")
    print("  ALL KEYS STORED IN ONE KEYCHAIN ENTRY — one Touch ID unlocks all.")
    print()

    existing: Dict[str, str] = {}
    try:
        if kc.has_credentials():
            existing = kc.get_all()
    except kc.SecretsError as e:
        print(f"  (existing blob unreadable — continuing fresh: {e})")

    updates: Dict[str, str] = {}
    kept = 0
    for spec in SPECS:
        has = spec.name in existing and existing[spec.name]
        if has and not all_keys:
            kept += 1
            continue
        raw = prompt_one(spec, currently_stored=bool(has))
        if raw is not None:
            updates[spec.name] = raw
        else:
            kept += 1

    if updates:
        try:
            kc.put_all(updates)
            print(f"\n  ✓ saved {len(updates)} keys to blob  (kept {kept})")
        except kc.SecretsError as e:
            print(f"\n  ✗ save failed: {e}")
            return 1
    else:
        print(f"\n  nothing to save  (kept {kept})")

    print()
    return run_test()


def run_rotate(key: str) -> int:
    spec = next((s for s in SPECS if s.name == key), None)
    if not spec:
        print(f"Unknown key: {key}")
        print(f"Valid: {', '.join(s.name for s in SPECS)}")
        return 2

    try:
        existing = kc.get_all() if kc.has_credentials() else {}
    except kc.SecretsError as e:
        print(f"✗ could not read blob: {e}")
        return 1

    has = spec.name in existing and existing[spec.name]
    raw = prompt_one(spec, currently_stored=bool(has))
    if raw is None:
        print("\n  (no change)")
        return 0

    try:
        kc.put(spec.name, raw)
        print(f"\n  ✓ {spec.name} updated")
    except kc.SecretsError as e:
        print(f"\n  ✗ save failed: {e}")
        return 1

    print()
    return run_test()


def run_status() -> int:
    print()
    if not kc.has_credentials():
        print("  blob: not stored")
        print("  run:  python3 setup_qbo.py")
        return 0

    print("  blob: present  (Touch ID required for details)")
    try:
        present = set(kc.list_stored())
    except kc.SecretsError as e:
        print(f"  ✗ could not read: {e}")
        return 1

    print()
    for spec in SPECS:
        mark = "✓" if spec.name in present else "·"
        print(f"    {mark}  {spec.name}")
    print()
    print(f"  {len(present & {s.name for s in SPECS})} / {len(SPECS)} stored")
    return 0


def run_purge() -> int:
    if input("Wipe the entire QBO Keychain blob? [y/N] ").strip().lower() != "y":
        print("canceled.")
        return 0
    n = kc.purge_all()
    print(f"deleted {n} entr{'y' if n == 1 else 'ies'}.")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true", help="Re-prompt every key")
    ap.add_argument("--rotate", metavar="KEY", help="Rotate one key")
    ap.add_argument("--status", action="store_true", help="Show stored keys")
    ap.add_argument("--test", action="store_true", help="Auth test only")
    ap.add_argument("--purge", action="store_true", help="Wipe the blob")
    args = ap.parse_args()

    try:
        if args.status: return run_status()
        if args.purge:  return run_purge()
        if args.test:   return run_test()
        if args.rotate: return run_rotate(args.rotate)
        return run_setup(all_keys=args.all)
    except KeyboardInterrupt:
        print("\ninterrupted.")
        return 130


if __name__ == "__main__":
    sys.exit(main())
