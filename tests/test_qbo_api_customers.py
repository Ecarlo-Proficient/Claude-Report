"""shared/qbo_api.pick_customer - one customer per project # when QBO holds duplicates."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shared.qbo_api import pick_customer  # noqa: E402


def _c(i, name, created):
    return {"Id": i, "DisplayName": name, "MetaData": {"CreateTime": created}}


def test_exact_name_beats_a_typo_twin():
    a, b = _c("1", "RP7340", "2026-01-28T09:02:55"), _c("2", "RP7340 -FTW", "2026-01-28T09:03:20")
    assert pick_customer("RP7340", [b, a], {})["Id"] == "1"


def test_invoices_beat_an_empty_duplicate():
    a, b = _c("29851", "RP7074", "2025-10-14"), _c("29251", "RP7074", "2025-08-21")
    assert pick_customer("RP7074", [a, b], {"29251": 1})["Id"] == "29251"


def test_oldest_record_when_nothing_else_separates_them():
    a, b = _c("30312", "RP7340", "2026-01-28T09:02:55"), _c("30313", "RP7340", "2026-01-28T09:03:20")
    assert pick_customer("RP7340", [b, a], {})["Id"] == "30312"


# ── split billing (shared/job_rulings `customers: all`, RP7401-FTW) ──
def test_one_customer_reports_as_itself():
    from shared.qbo_api import customer_ids, report_customers
    assert customer_ids({"id": "30563"}) == ["30563"]
    assert report_customers({"id": "30563"}) == "30563"


def test_split_billing_reports_every_customer_main_first():
    from shared.qbo_api import customer_ids, report_customers
    cust = {"id": "30563", "ids": ["30563", "30565"]}
    assert customer_ids(cust) == ["30563", "30565"]
    assert report_customers(cust) == "30563,30565"


def test_invoice_pull_uses_IN_for_several_customers(monkeypatch):
    from shared import qbo_api
    seen = []
    monkeypatch.setattr(qbo_api, "query_all", lambda a, c, ent, where="": seen.append(where) or [])
    qbo_api.fetch_customer_invoices("t", "r", "30563,30565")
    qbo_api.fetch_customer_invoices("t", "r", "30563")
    assert seen == ["CustomerRef IN ('30563', '30565')", "CustomerRef = '30563'"]


def test_customers_ruling_is_exact_job_and_not_a_finding(tmp_path):
    import json
    from shared import job_rulings
    f = tmp_path / "r.json"
    f.write_text(json.dumps({"RP7401-FTW": [{"kind": "customers", "rule": "all", "note": "split"}]}))
    assert job_rulings.all_customers("RP7401-FTW", f)
    assert not job_rulings.all_customers("RP7401", f)
    assert job_rulings.findings("RP7401-FTW", f) == []
