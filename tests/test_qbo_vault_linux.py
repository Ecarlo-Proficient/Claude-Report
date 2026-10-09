"""Linux (Docker) key reading: the office server takes its keys from the environment (docker/secrets.env)."""
import base64
import importlib
import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


@pytest.fixture
def vault(monkeypatch, tmp_path):
    monkeypatch.setattr(sys, "platform", "linux")
    monkeypatch.setenv("QBO_SECRETS_FILE", str(tmp_path / "qbo_secrets.json"))
    for k in ("QBO_CLIENT_ID", "QBO_CLIENT_SECRET", "QBO_COMPANY_ID", "QBO_REFRESH_TOKEN", "MIRROR_KEY", "JT_GRANT_KEY"):
        monkeypatch.delenv(k, raising=False)
    from shared import qbo_vault
    mod = importlib.reload(qbo_vault)
    yield mod
    monkeypatch.undo()
    importlib.reload(qbo_vault)


def test_mirror_key_read_from_env(vault, monkeypatch):
    key = base64.b64encode(os.urandom(32)).decode()
    monkeypatch.setenv("MIRROR_KEY", key)
    vault._cache = None
    assert vault.get("MIRROR_KEY") == key


def test_qbo_login_needs_only_the_four_qbo_keys(vault, monkeypatch):
    for k in vault.QBO_REQUIRED:
        monkeypatch.setenv(k, "x")
    assert vault.has_credentials()          # no JT_GRANT_KEY on the server
