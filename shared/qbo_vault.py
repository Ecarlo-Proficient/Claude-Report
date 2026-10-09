#!/usr/bin/env python3
"""
qbo_vault.py — Cross-platform credential store for QBO.

PLATFORMS
  macOS  → single-blob login-keychain entry (service='automation-qbo').
           NOT biometric-gated: a standard login-keychain item, so the real
           gate is the login keychain being unlocked (see DESIGN below).
  Linux  → environment variables QBO_CLIENT_ID / QBO_CLIENT_SECRET /
           QBO_COMPANY_ID / QBO_REFRESH_TOKEN. Used by the Docker container.
           Token rotation writes to a JSON file at QBO_SECRETS_FILE (default
           /data/qbo_secrets.json) so rotated refresh tokens survive container
           restarts. Reads prefer the file when present, falling back to env vars.

DESIGN (macOS)
  All QBO keys live together in ONE login-keychain entry, stored as a
  base64-encoded JSON blob via `security add-generic-password ... -T ""`.

  IMPORTANT (corrected 2026-08-17): this is NOT a biometric / Touch ID ACL.
  The `security` CLI cannot create a Touch-ID-bound item (that needs a
  SecAccessControl through the Security framework). Reading the blob triggers
  macOS keychain access control: on an UNTRUSTED read it shows the "security
  wants to use your confidential information" confirmation, which on a Touch
  ID Mac you MAY approve with a fingerprint - that is the historical "one
  Touch ID per run". But the prompt is not guaranteed: once /usr/bin/security
  is trusted for the item (you clicked "Always Allow", and/or put() re-created
  the item on a refresh-token rotation), reads are SILENT. The real security
  boundary is the login keychain being unlocked (it unlocks at login), not a
  fingerprint.

  Why one blob: reading it = ONE keychain access instead of five, and every
  key becomes available for the rest of the process. Within a single Python
  process the blob is decrypted once and cached in memory; later get() calls
  hit the cache with zero extra Keychain interaction.

ISOLATION → LIBRARY (the user 2026-07-17)
  Original design: one blob per service. REVISED by the user: this blob
  (service 'automation-qbo') is now THE key library — every new
  integration's key lives here (JT_GRANT_KEY = JobTread joined
  2026-07-17), one place to track them all, one keychain read per run.
  The Notion/Teams/invoice-sync blobs predate the decision and stay
  where they are (historical exceptions, not the pattern).

Public API (identical across platforms):
  get_all()   -> dict[str, str]    # reads the blob on Mac (may prompt, then cached); env+file on Linux
  get(key)    -> str               # convenience on top of get_all
  put(key, value)                  # update one key (Keychain on Mac, file on Linux)
  put_all(values)                  # update multiple keys at once
  delete(key) -> bool              # remove one key
  has_credentials() -> bool        # existence check only on Mac (no blob read)
  list_stored() -> list[str]       # keys present
  purge_all() -> int               # wipe (Keychain on Mac, file on Linux)
  clear_cache()                    # force next get_all to re-read the blob
  KNOWN_KEYS, SecretsError
"""
from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Optional

try:
    from shared import key_broker
except ImportError:  # imported with shared/ itself on sys.path
    import key_broker  # type: ignore

SERVICE = "automation-qbo"
LABEL = "credentials"
ACCOUNT = os.environ.get("USER") or "user"

KNOWN_KEYS: List[str] = [
    "QBO_CLIENT_ID",
    "QBO_CLIENT_SECRET",
    "QBO_COMPANY_ID",
    "QBO_REFRESH_TOKEN",
    "JT_GRANT_KEY",       # JobTread Pave API grant key (read-only grant)
]

# Linux (Docker) only: other library keys the container takes from its environment (docker/secrets.env).
LINUX_EXTRA_KEYS: List[str] = ["MIRROR_KEY"]

# The four keys a QuickBooks login needs (JT_GRANT_KEY above is optional).
QBO_REQUIRED: List[str] = ["QBO_CLIENT_ID", "QBO_CLIENT_SECRET", "QBO_COMPANY_ID", "QBO_REFRESH_TOKEN"]

# Platform routing. Mac uses Keychain; Linux uses env vars + file persistence.
_IS_MAC = sys.platform == "darwin"

# Linux-only persistence file for rotated tokens. Defaults to /data which is a
# typical Docker volume mount point. Override with QBO_SECRETS_FILE env var.
_SECRETS_FILE = Path(os.getenv("QBO_SECRETS_FILE", "/data/qbo_secrets.json"))

# In-process cache of the decrypted blob.
_cache: Optional[Dict[str, str]] = None


class SecretsError(RuntimeError):
    pass


def _sec(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["/usr/bin/security", *args], capture_output=True, text=True
    )


# ────────── macOS Keychain backend ──────────

def _read_blob_mac() -> Dict[str, str]:
    """Fetch + decode the blob from Keychain. Empty dict if not yet created."""
    r = _sec("find-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL, "-w")
    if r.returncode == 44:  # errSecItemNotFound
        return {}
    if r.returncode != 0:
        raise SecretsError(f"Keychain read failed: {r.stderr.strip() or 'unknown error'}")
    raw = r.stdout.rstrip("\n")
    try:
        data = json.loads(base64.b64decode(raw).decode())
    except Exception as e:
        raise SecretsError(f"stored blob is corrupt ({e}). Run --purge then setup again.")
    if not isinstance(data, dict):
        raise SecretsError("stored blob is not a dict — run --purge then setup again.")
    return data


def _add_via_stdin(encoded: str) -> subprocess.CompletedProcess:
    """`security add-generic-password` with the blob on STDIN (`security -i`), never on the command line:
    an argv is readable by any process on the Mac (`ps`) while it runs, and this write happens on every
    refresh-token rotation (security review 09/29/2026). Same item, same ACL as the argv form - verified by
    comparing `dump-keychain -a` of both on throwaway items (no app trusted, `-T ""`)."""
    for v in (ACCOUNT, SERVICE, LABEL):
        if any(c in v for c in '"\\\n'):
            raise SecretsError("keychain account/service/label contains a quote, backslash or newline")
    # base64 is [A-Za-z0-9+/=] only, so it is safe inside the interactive parser's double quotes
    line = (f'add-generic-password -a "{ACCOUNT}" -s "{SERVICE}" -l "{LABEL}" '
            f'-w "{encoded}" -U -T ""\n')
    return subprocess.run(["/usr/bin/security", "-i"], input=line, capture_output=True, text=True)


def _write_blob_mac(data: Dict[str, str]) -> None:
    """Encode + store in the login keychain (security -T ""). Overwrites if
    exists. NOT a biometric ACL - see the module docstring. The blob travels on
    stdin, never argv (see _add_via_stdin)."""
    encoded = base64.b64encode(json.dumps(data).encode()).decode()
    _sec("delete-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL)
    r = _add_via_stdin(encoded)
    if r.returncode != 0:                   # the delete already ran - one retry before giving up
        r = _add_via_stdin(encoded)
    if r.returncode != 0:
        err = (r.stderr or "").replace(encoded, "<blob>").strip()
        raise SecretsError(f"Keychain write failed: {err or 'unknown error'} - "
                           "the stored keys were removed; restore with shared/setup_qbo.py")


# ────────── Linux (Docker) backend ──────────

def _read_blob_linux() -> Dict[str, str]:
    """
    Linux read order:
      1. Persisted file (QBO_SECRETS_FILE) — holds the most recently rotated
         refresh token, written by put() on prior runs.
      2. Environment variables — initial values from container env / .env.
    File takes precedence so a rotated token survives container restarts.
    """
    file_data: Dict[str, str] = {}
    if _SECRETS_FILE.exists():
        try:
            file_data = json.loads(_SECRETS_FILE.read_text())
            if not isinstance(file_data, dict):
                raise SecretsError(f"{_SECRETS_FILE} is not a JSON object.")
        except json.JSONDecodeError as e:
            raise SecretsError(f"{_SECRETS_FILE} is corrupt JSON: {e}")

    data: Dict[str, str] = {}
    # KNOWN_KEYS plus LINUX_EXTRA_KEYS: the office server (docker/) reads MIRROR_KEY from its env too - it
    # was missing here, so the mirror failed with "MIRROR_KEY is not in the Keychain library" (10/02/2026).
    for key in KNOWN_KEYS + LINUX_EXTRA_KEYS:
        # File first, env second — rotated tokens override initial bootstrap.
        val = file_data.get(key) or os.getenv(key) or ""
        if val:
            data[key] = val
    return data


def _write_blob_linux(data: Dict[str, str]) -> None:
    """Persist the blob to QBO_SECRETS_FILE (chmod 600). Creates parent dir."""
    _SECRETS_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp = _SECRETS_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, sort_keys=True))
    try:
        os.chmod(tmp, 0o600)
    except OSError:
        pass  # bind-mounted volumes may not allow chmod; not fatal
    os.replace(tmp, _SECRETS_FILE)


# ────────── Platform dispatch ──────────

def _read_blob() -> Dict[str, str]:
    return _read_blob_mac() if _IS_MAC else _read_blob_linux()


def _write_blob(data: Dict[str, str]) -> None:
    if _IS_MAC:
        _write_blob_mac(data)
    else:
        _write_blob_linux(data)


def _brokered() -> bool:
    """After the Key Helper adoption (keyhelper/, 09/29/2026) the Keychain item trusts only the helper:
    every call below asks it instead of reading the Keychain. It never hands out the QuickBooks refresh
    token or the app secret - QuickBooks passes come from shared/qbo_api.get_pass()."""
    return _IS_MAC and key_broker.active()


def _broker_call(fn, *a):
    try:
        return fn(*a)
    except key_broker.BrokerError as e:
        raise SecretsError(str(e))


def get_all() -> Dict[str, str]:
    """Return all stored keys. Reads the login-keychain blob once per process
    (a keychain access that MAY prompt; see module docstring), then caches it.
    Brokered: only the keys this workspace may read - never the QuickBooks master keys."""
    global _cache
    if _cache is None:
        _cache = _broker_call(key_broker.keys) if _brokered() else _read_blob()
    return dict(_cache)


def get(key: str) -> str:
    if _brokered():
        v = _broker_call(key_broker.key, key)
        if not v:
            raise SecretsError(f"{key} not stored in Keychain")
        return v
    data = get_all()
    if key not in data:
        raise SecretsError(f"{key} not stored in Keychain")
    return data[key]


def put(key: str, value: str) -> None:
    """Update one key. Reads current blob, mutates, writes back."""
    global _cache
    if _brokered():
        _broker_call(key_broker.put, {key: value})
        _cache = None
        return
    data = _read_blob()
    data[key] = value
    _write_blob(data)
    _cache = data


def put_all(values: Dict[str, str]) -> None:
    """Merge `values` into the blob and persist."""
    global _cache
    if _brokered():
        _broker_call(key_broker.put, dict(values))
        _cache = None
        return
    data = _read_blob()
    data.update(values)
    _write_blob(data)
    _cache = data


def delete(key: str) -> bool:
    global _cache
    if _brokered():
        if key not in _broker_call(key_broker.names):
            return False
        _broker_call(key_broker.delete, [key])
        _cache = None
        return True
    data = _read_blob()
    if key not in data:
        return False
    del data[key]
    if data:
        _write_blob(data)
    else:
        _sec("delete-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL)
    _cache = data
    return True


def has_credentials() -> bool:
    """True if creds are present. Metadata-only on Mac (no blob read, no prompt)."""
    if _IS_MAC:
        r = _sec("find-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL)   # attributes only
        return r.returncode == 0
    # Linux: have creds if env vars OR persisted file provide the four QuickBooks keys. Not every KNOWN_KEY -
    # JT_GRANT_KEY (JobTread) is optional, and the office server (docker/) has none, so requiring it failed every run.
    blob = _read_blob_linux()
    return all(blob.get(k) for k in QBO_REQUIRED)


def list_stored() -> List[str]:
    """Keys actually present in the blob. Reads the blob on Mac (may prompt)."""
    if _brokered():
        return _broker_call(key_broker.names)
    return list(get_all().keys())


def purge_all() -> int:
    """Delete the entire blob. Returns 1 if deleted, 0 if nothing there."""
    global _cache
    _cache = {}
    if _brokered():
        _broker_call(key_broker.purge)
        return 1
    if _IS_MAC:
        r = _sec("delete-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL)
        return 1 if r.returncode == 0 else 0
    if _SECRETS_FILE.exists():
        _SECRETS_FILE.unlink()
        return 1
    return 0


# ────────── Non-QuickBooks keys (Notion, Teams) ──────────
# Until 09/29/2026 the Notion token and the Teams paid-notice webhook sat in an older Keychain store
# (`keyring`, service `proficient-automation-worker`) that any program could read. They now live in the key
# library behind Key Helper under these names; keyhelper/migrate_keys.py moves them once.
LEGACY_SERVICE = "proficient-automation-worker"
LEGACY_ACCOUNTS = {                      # library name -> the old keyring account
    "NOTION_SECRET": "notion",
    "TEAMS_WEBHOOK_MFD_PAID": "teams_webhook_mfd_paid",
    "TEAMS_WEBHOOK_ALERTS": "teams_webhook_alerts",
}


def get_secret(name: str) -> str:
    """THE lookup for a non-QuickBooks key; "" when none is set (callers decide if that is fatal).
    Order: the environment variable of the same name (Linux / the office server / CI, or a deliberate
    override) -> Key Helper (Mac, once adopted) -> the old keyring item (only until migrate_keys.py runs)."""
    env = (os.getenv(name) or "").strip()
    if env:
        return env
    if _brokered():
        try:
            v = key_broker.key(name)
        except key_broker.BrokerError as e:
            raise SecretsError(str(e))
        if v:
            return v.strip()
    acct = LEGACY_ACCOUNTS.get(name)
    if acct and _IS_MAC:
        try:
            import keyring
            return (keyring.get_password(LEGACY_SERVICE, acct) or "").strip()
        except Exception:  # noqa: BLE001 - an absent / locked legacy item is simply "not set"
            return ""
    return ""


def clear_cache() -> None:
    """Forget the in-process cache. Next get_all() re-reads the blob (may prompt)."""
    global _cache
    _cache = None


if __name__ == "__main__":
    print(f"service={SERVICE} label={LABEL} account={ACCOUNT}")
    if has_credentials():
        print("blob: present (a keychain read is needed to enumerate keys)")
    else:
        print("blob: none")
