"""The QuickBooks company (realm) id never reaches stdout or logs (owner 2026-08-06): every error _api_get raises
masks it, while keeping the HTTP status and the QuickBooks message. No network, no Keychain."""
import sys
from pathlib import Path

import pytest
import requests

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from shared import qbo_api   # noqa: E402

REALM = "9130354123456789"
PATH = f"/v3/company/{REALM}/query"


class _Resp:
    def __init__(self, status, text):
        self.status_code = status
        self.text = text


def _raise_from(monkeypatch, fake_get) -> str:
    monkeypatch.setattr(qbo_api.requests, "get", fake_get)
    monkeypatch.setattr(qbo_api, "_CURRENT_ACCESS", None)
    monkeypatch.setattr("time.sleep", lambda *_: None)
    with pytest.raises(RuntimeError) as ei:
        qbo_api._api_get(PATH, "tok", {"query": "SELECT * FROM Bill"})
    return str(ei.value)


def test_4xx_message_masks_realm_keeps_status_and_fault(monkeypatch):
    body = f'{{"Fault":{{"Error":[{{"Message":"Invalid query","Detail":"realm {REALM} rejected"}}]}}}}'
    msg = _raise_from(monkeypatch, lambda *a, **k: _Resp(400, body))
    assert REALM not in msg
    assert "/v3/company/<company>/query" in msg
    assert "400" in msg and "Invalid query" in msg


def test_network_error_masks_realm_in_requests_url(monkeypatch):
    def boom(url, **k):
        raise requests.exceptions.ConnectionError(f"Max retries exceeded with url: {url}?query=x")
    msg = _raise_from(monkeypatch, boom)
    assert REALM not in msg
    assert "network error" in msg


def test_mask_realm_helper():
    deeplink = f"https://app.qbo.intuit.com/app/customerdetail?nameId=5&deeplinkcompanyid={REALM}"
    assert REALM not in qbo_api.mask_realm(deeplink)
    assert qbo_api.mask_realm(f"bare {REALM} id", REALM) == "bare <company> id"
    assert qbo_api.mask_realm("400: bad request") == "400: bad request"
