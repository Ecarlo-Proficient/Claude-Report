"""shared/qbo_mirror - the raw QBO mirror, offline (in-memory SQLite, fake QBO)."""
import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shared import qbo_mirror as m  # noqa: E402


def _bill(i, total=10.0, upd="2026-09-16T10:00:00-05:00", tok=0):
    return {"Id": str(i), "SyncToken": str(tok), "TxnDate": "2026-09-01", "DocNumber": f"B{i}",
            "TotalAmt": total, "VendorRef": {"value": "7", "name": "RCI"},
            "MetaData": {"CreateTime": "2026-09-01T00:00:00-05:00", "LastUpdatedTime": upd},
            "Line": [{"Amount": total}]}


def test_upsert_replace_and_load_roundtrip():
    con = m.connect(":memory:")
    assert m.upsert_many(con, "Bill", [_bill(1), _bill(2)], "t1") == 2
    assert m.upsert_many(con, "Bill", [_bill(1, total=99.0, tok=1)], "t2") == 1
    rows = m.load("Bill", con=con)
    assert {r["Id"]: r["TotalAmt"] for r in rows} == {"1": 99.0, "2": 10.0}
    r = con.execute("SELECT sync_token, total, doc_number, seen_at FROM qbo_bill WHERE id='1'").fetchone()
    assert tuple(r) == (1, 99.0, "B1", "t2")
    assert m.get("Bill", "2", con=con)["Line"][0]["Amount"] == 10.0


def test_deleted_rows_are_kept_but_hidden_and_come_back_alive():
    con = m.connect(":memory:")
    m.upsert_many(con, "Bill", [_bill(1), _bill(2)], "t1")
    assert m.mark_deleted(con, "Bill", ["2", "nope"], "t2") == 1
    assert m.mark_deleted(con, "Bill", ["2"], "t3") == 0          # already flagged
    assert [r["Id"] for r in m.load("Bill", con=con)] == ["1"]
    assert len(m.load("Bill", con=con, include_deleted=True)) == 2
    m.upsert_many(con, "Bill", [_bill(2, tok=3)], "t4")               # QBO restored it
    assert [r["Id"] for r in m.load("Bill", con=con)] == ["1", "2"]


def test_load_where_over_summary_columns():
    con = m.connect(":memory:")
    m.upsert_many(con, "Bill", [_bill(1, total=5.0), _bill(2, total=500.0)], "t")
    assert [r["Id"] for r in m.load("Bill", "total >= ?", (100,), con=con)] == ["2"]


def test_parse_cdc_alive_deleted_capped():
    data = {"CDCResponse": [{"QueryResponse": [
        {"Bill": [_bill(1, upd="2026-09-17T01:00:00-05:00"),
                  {"Id": "9", "status": "Deleted",
                   "MetaData": {"LastUpdatedTime": "2026-09-17T02:00:00-05:00"}}],
         "startPosition": 1, "maxResults": 2},
        {"Invoice": [], "maxResults": 0},
    ]}]}
    p = m.parse_cdc(data)
    assert [r["Id"] for r in p["Bill"]["alive"]] == ["1"]
    assert p["Bill"]["deleted"] == [("9", "2026-09-17T02:00:00-05:00")]
    assert p["Bill"]["newest"] == "2026-09-17T02:00:00-05:00"
    assert p["Bill"]["capped"] is False
    assert p["Invoice"]["alive"] == [] and p["Invoice"]["capped"] is False
    capped = {"CDCResponse": [{"QueryResponse": [{"Bill": [_bill(i) for i in range(m.CDC_CAP)],
                                                   "maxResults": m.CDC_CAP}]}]}
    assert m.parse_cdc(capped)["Bill"]["capped"] is True


def test_refresh_uses_cdc_inside_reach_and_stamps_start(monkeypatch):
    con = m.connect(":memory:")
    m.upsert_many(con, "Bill", [_bill(1), _bill(2)], "t0")
    m._meta_set(con, "last_refresh", "2026-09-16T12:00:00Z")
    calls = []

    def fake_cdc(access, cid, entities, since):
        calls.append((tuple(entities), since))
        return {"CDCResponse": [{"QueryResponse": [
            {"Bill": [_bill(2, total=77.0, tok=5),
                      {"Id": "1", "status": "Deleted", "MetaData": {"LastUpdatedTime": "2026-09-16T13:00:00Z"}}],
             "maxResults": 2}]}]}

    monkeypatch.setattr(m, "_cdc_call", fake_cdc)
    fixed = dt.datetime(2026, 9, 17, 9, 0, tzinfo=dt.timezone.utc)
    monkeypatch.setattr(m, "now_utc", lambda: fixed)
    res = m.refresh(con, "a", "c", progress=None)
    assert res["mode"] == "cdc" and res["upserted"] == 1 and res["deleted"] == 1
    assert calls[0][1] == "2026-09-16T11:50:00Z"            # last stamp minus the overlap
    assert set(calls[0][0]) == set(m.ENTITIES)
    assert m._meta_get(con, "last_refresh") == "2026-09-17T09:00:00Z"
    assert [r["Id"] for r in m.load("Bill", con=con)] == ["2"]
    assert m.load("Bill", con=con)[0]["TotalAmt"] == 77.0
    run = con.execute("SELECT mode, upserted, deleted FROM mirror_run").fetchone()
    assert tuple(run) == ("cdc", 1, 1)


def test_refresh_past_feed_reach_falls_back_to_updated_plus_sweep(monkeypatch):
    con = m.connect(":memory:")
    m.upsert_many(con, "Bill", [_bill(1), _bill(2)], "t0")
    m._meta_set(con, "last_refresh", "2026-07-01T00:00:00Z")      # 78 days ago
    seen = []

    def fake_pages(access, cid, entity, where="", select="*"):
        seen.append((entity, where, select))
        if entity != "Bill":
            return iter([])
        if select == "Id":
            return iter([[{"Id": "2"}, {"Id": "3"}]])              # 1 is gone, 3 is new
        if where.startswith("MetaData.LastUpdatedTime"):
            return iter([[_bill(2, total=42.0, tok=2)]])
        if where.startswith("Id IN"):
            return iter([[_bill(3)]])
        return iter([])

    monkeypatch.setattr(m, "_query_pages", fake_pages)
    monkeypatch.setattr(m, "now_utc", lambda: dt.datetime(2026, 9, 17, 9, 0, tzinfo=dt.timezone.utc))
    res = m.refresh(con, "a", "c", progress=None)
    assert res["mode"] == "updated+sweep"
    ids = {r["Id"]: r["TotalAmt"] for r in m.load("Bill", con=con)}
    assert ids == {"2": 42.0, "3": 10.0}
    assert m.load("Bill", con=con, include_deleted=True)[0]["Id"] == "1"
    assert any(e == "Customer" and "Active IN (true, false)" in w for e, w, _ in seen)


def test_seed_flags_rows_the_full_pull_did_not_see(monkeypatch):
    con = m.connect(":memory:")
    m.upsert_many(con, "Term", [{"Id": "1", "Name": "Net 30", "SyncToken": "0", "MetaData": {}},
                                {"Id": "2", "Name": "Old", "SyncToken": "0", "MetaData": {}}], "2026-01-01T00:00:00Z")
    monkeypatch.setattr(m, "count", lambda a, c, e: 1)
    monkeypatch.setattr(m, "_query_pages",
                        lambda a, c, e, where="", select="*": iter([[{"Id": "1", "Name": "Net 30", "SyncToken": "1", "MetaData": {}}]]))
    res = m.seed(con, "a", "c", entities=["Term"], progress=None)
    assert res["upserted"] == 1 and res["deleted"] == 1
    assert [r["Name"] for r in m.load("Term", con=con)] == ["Net 30"]
    assert m._meta_get(con, "last_refresh") is None          # a partial seed never moves the stamp


def test_status_counts(monkeypatch):
    con = m.connect(":memory:")
    m.upsert_many(con, "Invoice", [_bill(1), _bill(2)], "t")
    m.mark_deleted(con, "Invoice", ["2"], "t")
    monkeypatch.setattr(m, "db_path", lambda: Path("/nonexistent/x.sqlite3"))
    s = m.status(con)
    assert s["entities"]["Invoice"] == {"alive": 1, "deleted": 1, "newest": "2026-09-16T10:00:00-05:00"}
    assert s["size_mb"] == 0.0


# ───────── serving query_all: the WHERE grammar the tools use ─────────

def _inv(i, date, bal, cust="12", doc=None, active=True):
    return {"Id": str(i), "SyncToken": "0", "TxnDate": date, "Balance": bal, "DocNumber": doc or f"I{i}",
            "CustomerRef": {"value": cust, "name": "GC:RP1"}, "Active": active,
            "MetaData": {"LastUpdatedTime": "2026-09-16T10:00:00-05:00"}}


def test_parse_where_grammar():
    t = m.parse_where("Balance = '0' AND TxnDate >= '2026-01-01'")
    assert t == [("Balance", "=", "0"), ("TxnDate", ">=", "2026-01-01")]
    assert m.parse_where("Active IN (true, false)") == [("Active", "IN", [True, False])]
    assert m.parse_where("AccountType IN ('Bank','Credit Card') AND Active=true") == [
        ("AccountType", "IN", ["Bank", "Credit Card"]), ("Active", "=", True)]
    assert m.parse_where("DisplayName LIKE '%Ready Mix AND Sons%'") == [("DisplayName", "LIKE", "%Ready Mix AND Sons%")]
    assert m.parse_where("") == []
    import pytest
    with pytest.raises(m.WhereError):
        m.parse_where("TxnDate BETWEEN 'a' AND 'b'")


def test_match_where_terms():
    r = _inv(1, "2026-06-01", 250.5)
    ok = lambda w: m.match_where(r, m.parse_where(w))   # noqa: E731
    assert ok("Balance > '0'") and not ok("Balance = '0'")
    assert ok("TxnDate >= '2026-06-01' AND TxnDate <= '2026-06-30'") and not ok("TxnDate < '2026-06-01'")
    assert ok("CustomerRef = '12'") and not ok("VendorRef = '12'")
    assert ok("Id IN ('1','7')") and ok("DocNumber IN ('I1')") and not ok("Id IN ('7')")
    assert ok("Active = true") and ok("Active IN (true, false)") and not ok("Active = false")
    assert m.match_where({"DisplayName": "Cowtown Ready Mix"}, m.parse_where("DisplayName LIKE '%Ready%'"))
    assert not m.match_where({"DisplayName": "Cowtown"}, m.parse_where("DisplayName LIKE '%Ready%'"))
    assert m.match_where({"DisplayName": "ACME concrete"}, m.parse_where("DisplayName LIKE '%Concrete%'"))  # QBO LIKE ignores case
    assert m.match_where({"AccountType": "Bank", "Active": True},
                         m.parse_where("AccountType IN ('Bank','Credit Card') AND Active=true"))
    assert m.match_where({"Classification": "Liability"}, m.parse_where("Classification = 'Liability'"))


def test_query_matches_qbo_semantics_and_pushdown(monkeypatch):
    con = m.connect(":memory:")
    m.upsert_many(con, "Invoice", [_inv(1, "2026-01-05", 0), _inv(2, "2026-06-01", 10.0),
                                   _inv(3, "2026-07-01", 5.0, cust="99")], "t")
    m.mark_deleted(con, "Invoice", ["3"], "t")
    m._meta_set(con, "last_refresh", "2026-09-17T12:00:00Z")
    monkeypatch.setattr(m, "_STAMP_SHOWN", [True])
    ids = lambda w: sorted(r["Id"] for r in m.query("Invoice", w, con=con))   # noqa: E731
    assert ids("") == ["1", "2"]                                   # deleted never served
    assert ids("Balance > '0'") == ["2"]
    assert ids("TxnDate >= '2026-01-01' AND TxnDate <= '2026-01-31'") == ["1"]
    assert ids("DocNumber IN ('I1','I2')") == ["1", "2"]
    # a name list hides inactive rows unless Active is mentioned - QBO's default
    m.upsert_many(con, "Customer", [{"Id": "5", "DisplayName": "A", "Active": True, "SyncToken": "0", "MetaData": {}},
                                    {"Id": "6", "DisplayName": "B", "Active": False, "SyncToken": "0", "MetaData": {}}], "t")
    assert [r["Id"] for r in m.query("Customer", "", con=con)] == ["5"]
    assert [r["Id"] for r in m.query("Customer", "Active = false", con=con)] == ["6"]
    assert sorted(r["Id"] for r in m.query("Customer", "Active IN (true, false)", con=con)) == ["5", "6"]


def test_serves_respects_live_switch_and_missing_mirror(monkeypatch, tmp_path):
    monkeypatch.setattr(m, "db_path", lambda: tmp_path / "none.sqlite3")
    assert m.serves("Bill") is False                               # never seeded here
    con = m.connect(tmp_path / "none.sqlite3")
    m._meta_set(con, "last_refresh", "2026-09-17T12:00:00Z"); con.commit(); con.close()
    assert m.serves("Bill") is True
    assert m.serves("Attachable") is False                         # outside the mirror
    monkeypatch.setenv("ACB_QBO_LIVE", "1")
    assert m.serves("Bill") is False
