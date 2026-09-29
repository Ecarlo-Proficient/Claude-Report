"""Key Helper client + the routing around it (security review 09/29/2026). A fake helper on a real Unix
socket stands in for keyhelper/KeyHelper.swift, so this runs anywhere (CI is Linux): no Keychain, no
QuickBooks, no password prompt."""
import json
import os
import re
import shutil
import socket
import sys
import tempfile
import threading
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from shared import key_broker, qbo_api, qbo_vault   # noqa: E402


@pytest.fixture
def fake_helper(monkeypatch):
    """A socket server answering like Key Helper; records every request."""
    d = tempfile.mkdtemp(prefix="kb", dir="/tmp")          # short: AF_UNIX paths stop at 104 bytes on macOS
    sock = Path(d) / "k.sock"
    marker = Path(d) / "adopted"
    marker.write_text("x")
    seen = []
    answers = {
        "pass": {"ok": True, "access_token": "tok-1h", "company_id": "company-1", "expires_at": 0},
        "key": {"ok": True, "values": {"MIRROR_KEY": "m-key"}},
        "keys": {"ok": True, "values": {"MIRROR_KEY": "m-key", "JT_GRANT_KEY": "jt"}},
        "names": {"ok": True, "names": ["MIRROR_KEY", "QBO_REFRESH_TOKEN"]},
        "put": {"ok": True},
        "get": {"ok": True, "status": 200, "body": "{\"QueryResponse\": {}}"},
        "status": {"ok": True, "unlocked": False},
    }
    srv = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
    srv.bind(str(sock))
    srv.listen(8)
    stop = threading.Event()

    def loop():
        srv.settimeout(0.2)
        while not stop.is_set():
            try:
                c, _ = srv.accept()
            except (socket.timeout, OSError):
                continue
            buf = b""
            while not buf.endswith(b"\n"):
                b = c.recv(65536)
                if not b:
                    break
                buf += b
            req = json.loads(buf)
            seen.append(req)
            op = req["op"]
            if op == "pass" and req.get("profile") == "readonly":
                resp = {"error": "refused: 'readonly' is read-only here - use the read request, not a pass"}
            else:
                resp = answers.get(op, {"error": f"unknown request '{op}'"})
            c.sendall(json.dumps(resp).encode() + b"\n")
            c.close()

    t = threading.Thread(target=loop, daemon=True)
    t.start()
    monkeypatch.setattr(key_broker, "SOCK", sock)
    monkeypatch.setattr(key_broker, "MARKER", marker)
    monkeypatch.setattr(key_broker.sys, "platform", "darwin")
    monkeypatch.setattr(qbo_vault, "_IS_MAC", True)
    monkeypatch.setattr(qbo_vault, "_cache", None)
    yield seen
    stop.set()
    t.join(1)
    srv.close()
    shutil.rmtree(d, ignore_errors=True)


def test_inactive_until_adopted(monkeypatch, tmp_path):
    monkeypatch.setattr(key_broker, "MARKER", tmp_path / "adopted")
    monkeypatch.setattr(key_broker.sys, "platform", "darwin")
    assert not key_broker.active()
    (tmp_path / "adopted").write_text("x")
    assert key_broker.active()
    monkeypatch.setattr(key_broker.sys, "platform", "linux")
    assert not key_broker.active()                          # the office server never uses the Mac helper


def test_pass_and_keys(fake_helper):
    assert key_broker.qbo_pass() == ("tok-1h", "company-1")
    assert fake_helper[-1] == {"op": "pass", "profile": "proficient", "fresh": False}
    assert key_broker.key("MIRROR_KEY") == "m-key"
    assert key_broker.qbo_get("query", {"query": "select * from Bill"}) == (200, "{\"QueryResponse\": {}}")
    assert fake_helper[-1]["path"] == "query"


def test_refusal_is_an_error(fake_helper):
    with pytest.raises(key_broker.BrokerError, match="read-only"):
        key_broker.qbo_pass(profile="readonly")


def test_login_goes_through_the_helper(fake_helper):
    assert qbo_api.get_pass() == ("tok-1h", "company-1")
    qbo_api.get_pass(fresh=True)
    assert fake_helper[-1]["fresh"] is True


def test_vault_never_reads_the_keychain_once_adopted(fake_helper, monkeypatch):
    def boom(*a, **k):
        raise AssertionError("the Keychain was read directly")
    monkeypatch.setattr(qbo_vault, "_read_blob", boom)
    monkeypatch.setattr(qbo_vault, "_read_blob_mac", boom)
    assert qbo_vault.get("MIRROR_KEY") == "m-key"
    assert "QBO_REFRESH_TOKEN" not in qbo_vault.get_all()   # master keys are never handed out
    assert qbo_vault.list_stored() == ["MIRROR_KEY", "QBO_REFRESH_TOKEN"]
    qbo_vault.put("JT_GRANT_KEY", "new")
    assert fake_helper[-1] == {"op": "put", "values": {"JT_GRANT_KEY": "new"}}


def test_one_login_path():
    """Only shared/qbo_api.py (and the helper itself) may talk to Intuit's token endpoint - every tool
    gets its pass from the shared login, so no tool ever holds the refresh token."""
    allowed = {"shared/qbo_api.py", "keyhelper/KeyHelper.swift"}
    hits = set()
    for p in ROOT.rglob("*"):
        if p.suffix not in {".py", ".swift", ".sh"} or ".git" in p.parts or "tests" in p.parts:
            continue
        rel = p.relative_to(ROOT).as_posix()
        if re.search(r"oauth2/v1/tokens", p.read_text(errors="ignore")):
            hits.add(rel)
    assert hits <= allowed, f"a tool renews the QuickBooks login itself: {sorted(hits - allowed)}"
