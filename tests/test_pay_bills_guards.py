"""The guardrails on the ledger's QBO money writes (ledger/pay_bills.py) - offline, no QBO, no Touch ID."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "ledger"))
sys.path.insert(0, str(ROOT))

import pay_bills as pb  # noqa: E402


def test_reference_cap_is_qbo_21():
    assert pb.clean_ref("ach", " 25783-123456831813153 ") == "25783-123456831813153"
    with pytest.raises(ValueError):
        pb.clean_ref("ach", "25783-1234568318131534")
    with pytest.raises(ValueError):
        pb.clean_ref("check", "")
    assert pb.clean_ref("print", "anything") == pb.TO_PRINT


def test_unauthorized_mac_is_refused(monkeypatch):
    monkeypatch.setattr(pb, "authorized", lambda: (False, "locked"))
    assert pb.commit({"run_token": "x" * 32})["error"] == "locked"
    assert pb.assign_number("123", "4567")["error"] == "locked"


def test_write_needs_a_live_review_token(monkeypatch):
    monkeypatch.setattr(pb, "authorized", lambda: (True, ""))
    monkeypatch.setattr(pb, "_audit", lambda *a, **k: None)
    assert "expired or was already used" in pb.commit({"run_token": "never-issued"})["error"]
    pb._TOKENS["t1"] = {"at": pb.time.time(), "vendors": {"9": {"100": 50.0}}, "credits": {"9": {}}, "names": {}}
    body = {"run_token": "t1", "vendors": [{"vendor_id": "9", "bills": [{"bill_id": "100", "amount": 51}]}]}
    assert "not what the review showed" in pb.commit(body)["error"]
    assert "expired or was already used" in pb.commit(body)["error"]      # spent even though it was refused


def test_approval_text_layout():
    t = pb.approval_text([("JMP", 10.5, "2026-09-28", "to print"), ("RCI", 2, "2026-09-28", "ACH 123")], "Frost ****8073", "to print")
    assert t.splitlines() == ["approve these vendor payments:", "JMP · $10.50 · 09/28/2026", "Frost ****8073 · to print",
                              "RCI · $2.00 · 09/28/2026", "Frost ****8073 · ACH 123", "Total $12.50"]
