"""invoice-sync: an invoice whose Notion page already matches QBO (and was synced today) is not rewritten.

The page is read back in Notion's shape (plain_text, select name, relation ids) and compared with the
properties the sync would write. Equal -> skip. Any difference, an unknown property shape, or a page not
stamped Last Synced today -> write, exactly as before.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "invoice-sync"))

import invoice_sync as s  # noqa: E402

TODAY = "2026-10-08"


def _write_props():
    return {
        "Invoice #": s._ts_prop_title("34547"),
        "Invoice ID": s._ts_prop_rich_text("98765"),
        "Memo": s._ts_prop_rich_text("MFD325 - August Draw\r\n(Period:07/02/2026 - 08/01/2026)"),
        "Date": {"date": {"start": "2026-08-01"}},
        "Due Date": s._ts_prop_date(None),
        "Total Amount": s._ts_prop_number(510453.1),
        "Open balance": s._ts_prop_number(510453.1),
        "Status": s._ts_prop_select("Unpaid"),
        "Aging Bucket": s._ts_prop_select(None),
        "QBO Link": {"url": "https://qbo.intuit.com/app/invoice?txnId=98765"},
        "Customer": s._ts_prop_relation("1a2b3c4d-0000-1111-2222-333344445555"),
        "Last Synced": {"date": {"start": TODAY}},
    }


def _read_back(last_synced=TODAY, **over):
    """The same page as Notion returns it from a query."""
    page = {
        "Invoice #": {"type": "title", "title": [{"plain_text": "34547"}]},
        "Invoice ID": {"type": "rich_text", "rich_text": [{"plain_text": "98765"}]},
        "Memo": {"type": "rich_text", "rich_text": [{"plain_text": "MFD325 - August Draw\n"},
                                                    {"plain_text": "(Period:07/02/2026 - 08/01/2026)"}]},
        "Date": {"type": "date", "date": {"start": "2026-08-01", "end": None}},
        "Due Date": {"type": "date", "date": None},
        "Total Amount": {"type": "number", "number": 510453.1},
        "Open balance": {"type": "number", "number": 510453.10000001},
        "Status": {"type": "select", "select": {"id": "x", "name": "Unpaid", "color": "red"}},
        "Aging Bucket": {"type": "select", "select": None},
        "QBO Link": {"type": "url", "url": "https://qbo.intuit.com/app/invoice?txnId=98765"},
        "Customer": {"type": "relation", "relation": [{"id": "1a2b3c4d000011112222333344445555"}],
                     "has_more": False},
        "Last Synced": {"type": "date", "date": {"start": last_synced, "end": None}},
        "Quick Status": {"type": "select", "select": {"name": "Called GC"}},
    }
    page.update(over)
    return page


def test_unchanged_page_synced_today_is_skipped():
    assert s._needs_write(_write_props(), _read_back(), TODAY) is False


def test_balance_change_is_written():
    page = _read_back(**{"Open balance": {"type": "number", "number": 250000.0}})
    assert s._needs_write(_write_props(), page, TODAY) is True


def test_status_change_is_written():
    page = _read_back(**{"Status": {"type": "select", "select": {"name": "Partially Paid"}}})
    assert s._needs_write(_write_props(), page, TODAY) is True


def test_not_synced_today_is_written_once():
    assert s._needs_write(_write_props(), _read_back(last_synced="2026-10-07"), TODAY) is True


def test_customer_relation_change_is_written():
    page = _read_back(**{"Customer": {"type": "relation", "relation": []}})
    assert s._needs_write(_write_props(), page, TODAY) is True


def test_missing_property_or_new_page_is_written():
    page = _read_back()
    del page["Aging Bucket"]
    assert s._needs_write({**_write_props(), "Aging Bucket": s._ts_prop_select("31-60")}, page, TODAY) is True
    assert s._needs_write(_write_props(), {}, TODAY) is True


def test_unknown_shape_is_never_skipped():
    props = {**_write_props(), "Odd": {"checkbox": True}}
    page = _read_back(Odd={"type": "checkbox", "checkbox": True})
    assert s._needs_write(props, page, TODAY) is True


def test_human_fields_are_ignored():
    page = _read_back(**{"Quick Status": {"type": "select", "select": {"name": "Lien filed"}}})
    assert s._needs_write(_write_props(), page, TODAY) is False
