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


def _write_blob_mac(data: Dict[str, str]) -> None:
    """Encode + store in the login keychain (security -T ""). Overwrites if
    exists. NOT a biometric ACL - see the module docstring."""
    encoded = base64.b64encode(json.dumps(data).encode()).decode()
    _sec("delete-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL)
    r = _sec(
        "add-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL,
        "-w", encoded, "-U", "-T", "",
    )
    if r.returncode != 0:
        raise SecretsError(f"Keychain write failed: {r.stderr.strip() or 'unknown error'}")


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
    for key in KNOWN_KEYS:
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


def get_all() -> Dict[str, str]:
    """Return all stored keys. Reads the login-keychain blob once per process
    (a keychain access that MAY prompt; see module docstring), then caches it."""
    global _cache
    if _cache is None:
        _cache = _read_blob()
    return dict(_cache)


def get(key: str) -> str:
    data = get_all()
    if key not in data:
        raise SecretsError(f"{key} not stored in Keychain")
    return data[key]


def put(key: str, value: str) -> None:
    """Update one key. Reads current blob, mutates, writes back."""
    global _cache
    data = _read_blob()
    data[key] = value
    _write_blob(data)
    _cache = data


def put_all(values: Dict[str, str]) -> None:
    """Merge `values` into the blob and persist."""
    global _cache
    data = _read_blob()
    data.update(values)
    _write_blob(data)
    _cache = data


def delete(key: str) -> bool:
    global _cache
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
        r = _sec("find-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL)
        return r.returncode == 0
    # Linux: have creds if env vars OR persisted file provide all required keys.
    blob = _read_blob_linux()
    return all(blob.get(k) for k in KNOWN_KEYS)


def list_stored() -> List[str]:
    """Keys actually present in the blob. Reads the blob on Mac (may prompt)."""
    return list(get_all().keys())


def purge_all() -> int:
    """Delete the entire blob. Returns 1 if deleted, 0 if nothing there."""
    global _cache
    _cache = {}
    if _IS_MAC:
        r = _sec("delete-generic-password", "-a", ACCOUNT, "-s", SERVICE, "-l", LABEL)
        return 1 if r.returncode == 0 else 0
    if _SECRETS_FILE.exists():
        _SECRETS_FILE.unlink()
        return 1
    return 0


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
