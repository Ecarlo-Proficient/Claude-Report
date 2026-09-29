"""
shared/key_broker.py - the client for Key Helper (keyhelper/), the Mac's key broker.

After the one-time adoption (`bash keyhelper/install.sh --adopt`) the key library in the Keychain trusts
only Key Helper. Tools then never read it: they ask the helper, which unlocks once per work session
(Touch ID or password) and answers over a private socket:

  qbo_pass(profile)        -> (access_token, company_id)   a one-hour QuickBooks pass, never the refresh token
  qbo_get(profile, path)   -> (status, body)               a read the helper runs itself (read-only profiles)
  key(name) / keys()       -> non-QuickBooks keys this workspace may read (MIRROR_KEY, JT_GRANT_KEY, ...)
  put / delete / purge     -> key library changes (a fresh owner check every time)
  status() / lock()

active() is False until the adoption marker exists, so before the cutover every tool keeps the old path
(shared/qbo_vault.py). Linux (the future office server) never uses this module.
"""
from __future__ import annotations

import json
import os
import socket
import subprocess
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

KEYS_DIR = Path.home() / "Library" / "Application Support" / "Proficient" / "keys"
SOCK = KEYS_DIR / "k.sock"
MARKER = KEYS_DIR / "adopted"
APP = Path.home() / "Library" / "Application Support" / "Proficient" / "bin" / "Key Helper.app"
PROFILE = "proficient"          # this company's QuickBooks connection

# a request can wait on the owner typing a password, so the read timeout is generous
_CONNECT_WAIT = 15
_ANSWER_WAIT = 300


class BrokerError(RuntimeError):
    pass


def active() -> bool:
    """True once the key library has been handed to Key Helper on this Mac."""
    return sys.platform == "darwin" and MARKER.exists()


def _launch() -> None:
    if not APP.exists():
        raise BrokerError(f"Key Helper is not installed ({APP}) - run: bash keyhelper/install.sh")
    subprocess.run(["/usr/bin/open", "-g", str(APP)], check=False, capture_output=True)


def request(op: str, **fields) -> dict:
    """One request to the helper; starts it if it is not running. Raises BrokerError on a refusal."""
    payload = (json.dumps({"op": op, **fields}) + "\n").encode()
    deadline = time.time() + _CONNECT_WAIT
    launched = False
    while True:
        s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        s.settimeout(_ANSWER_WAIT)
        try:
            s.connect(str(SOCK))
            break
        except (FileNotFoundError, ConnectionRefusedError):
            s.close()
            if not launched:
                _launch()
                launched = True
            if time.time() > deadline:
                raise BrokerError("Key Helper did not start - open it from "
                                  "~/Library/Application Support/Proficient/bin/")
            time.sleep(0.3)
    try:
        s.sendall(payload)
        chunks = []
        while True:
            b = s.recv(1 << 20)
            if not b:
                break
            chunks.append(b)
    except socket.timeout:
        raise BrokerError("Key Helper did not answer in time (was the password prompt left open?)")
    finally:
        s.close()
    try:
        resp = json.loads(b"".join(chunks) or b"{}")
    except json.JSONDecodeError:
        raise BrokerError("Key Helper sent an unreadable answer")
    if resp.get("error"):
        raise BrokerError(resp["error"])
    return resp


def qbo_pass(profile: str = PROFILE, fresh: bool = False) -> Tuple[str, str]:
    r = request("pass", profile=profile, fresh=bool(fresh))
    return r["access_token"], r["company_id"]


def qbo_get(path: str, params: Optional[dict] = None, profile: str = PROFILE) -> Tuple[int, str]:
    """GET <company>/<path> run by the helper - the only way a read-only profile reaches QuickBooks."""
    r = request("get", profile=profile, path=path.lstrip("/"), params=params or {})
    return int(r["status"]), r["body"]


def key(name: str) -> str:
    return request("key", name=name)["values"].get(name, "")


def keys() -> Dict[str, str]:
    return dict(request("keys")["values"])


def names() -> List[str]:
    return list(request("names")["names"])


def put(values: Dict[str, str]) -> None:
    request("put", values=dict(values))


def delete(names_: List[str]) -> None:
    request("delete", names=list(names_))


def purge() -> None:
    request("purge")


def status() -> dict:
    return request("status")


def lock() -> None:
    request("lock")
