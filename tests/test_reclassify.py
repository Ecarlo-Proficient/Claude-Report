"""ledger/reclassify.py - the line reader and the change rules (no QuickBooks, no mirror)."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import reclassify as R  # noqa: E402

L = {"account": {"10": {"Id": "10", "FullyQualifiedName": "Job Materials:Concrete", "Classification": "Expense"},
                 "11": {"Id": "11", "FullyQualifiedName": "Job Materials:Rebar", "Classification": "Expense"},
                 "20": {"Id": "20", "FullyQualifiedName": "Construction in Progress", "Classification": "Revenue"}},
     "item": {"1": {"Id": "1", "Name": "SL1", "ExpenseAccountRef": {"value": "10"}},
              "2": {"Id": "2", "Name": "SL2", "ExpenseAccountRef": {"value": "11"}, "IncomeAccountRef": {"value": "20"}}},
     "class": {"5": {"Id": "5", "FullyQualifiedName": "Residential"}, "6": {"Id": "6", "FullyQualifiedName": "Commercial"}},
     "project": {"100": {"Id": "100", "FullyQualifiedName": "GC:RP7242-FTW 1 Main"},
                 "101": {"Id": "101", "FullyQualifiedName": "GC:RP7242 1 Main"}}}

BILL = {"Id": "9", "DocNumber": "A1", "TxnDate": "2026-10-01", "VendorRef": {"name": "Supplier"}, "PrivateNote": "RP7242-FTW",
        "Line": [{"Id": "1", "Amount": 100.0, "Description": "pour FTW",
                  "ItemBasedExpenseLineDetail": {"ItemRef": {"value": "1"}, "ClassRef": {"value": "5"}, "CustomerRef": {"value": "100"}}},
                 {"Id": "2", "Amount": 50.0,
                  "AccountBasedExpenseLineDetail": {"AccountRef": {"value": "11"}, "CustomerRef": {"value": "100"}}}]}


def test_read_item_line_posts_to_items_account():
    r = R.read_line("Bill", BILL, BILL["Line"][0], L)
    assert (r["kind"], r["account_name"], r["item_name"], r["class_name"], r["proj"]) == \
        ("item", "Job Materials:Concrete", "SL1", "Residential", "RP7242-FTW")
    assert "account" not in r["edit"] and {"item", "class", "project", "memo", "docmemo"} <= set(r["edit"])


def test_category_line_takes_account_not_item():
    r = R.read_line("Bill", BILL, BILL["Line"][1], L)
    ch = R.clean_changes({"account": {"11": "10"}, "item": {"": "2"}})
    assert R.row_changes(r, ch) == {"account": ("11", "10")}


def test_project_class_memo_bulk():
    r = R.read_line("Bill", BILL, BILL["Line"][0], L)
    ch = R.clean_changes({"project": {"100": "101"}, "class": {"5": "6"},
                          "memo": {"mode": "replace", "find": "FTW", "text": "slab"},
                          "docmemo": {"mode": "set", "text": "RP7242"}})
    assert R.row_changes(r, ch) == {"project": ("100", "101"), "class": ("5", "6"),
                                    "memo": ("pour FTW", "pour slab"), "docmemo": ("RP7242-FTW", "RP7242")}


def test_unchanged_values_dropped():
    assert R.clean_changes({"class": {"5": "5", "6": ""}, "memo": {"mode": "replace", "find": ""}}) == {}


def test_signature_moves_with_any_field():
    r = R.read_line("Bill", BILL, BILL["Line"][0], L)
    t2 = {**BILL, "PrivateNote": "edited"}
    assert R.read_line("Bill", t2, BILL["Line"][0], L)["sig"] != r["sig"]


def test_journal_vendor_line_cannot_take_a_project():
    je = {"Id": "3", "TxnDate": "2026-10-01", "Line": [{"Id": "0", "Amount": 5.0, "JournalEntryLineDetail": {
        "PostingType": "Credit", "AccountRef": {"value": "10"}, "Entity": {"Type": "Vendor", "EntityRef": {"value": "7"}}}}]}
    r = R.read_line("JournalEntry", je, je["Line"][0], L)
    assert "project" not in r["edit"] and r["signed"] == -5.0


def test_apply_writes_customer_and_projects_tag():
    import copy
    t = copy.deepcopy(BILL)
    ln = t["Line"][0]
    ln["ProjectRef"] = {"value": "P100"}
    R._apply(t, ln, {"project": ("100", "101"), "item": ("1", "2")}, {"101": "P101"})
    d = ln["ItemBasedExpenseLineDetail"]
    assert d["CustomerRef"] == {"value": "101"} and ln["ProjectRef"] == {"value": "P101"} and d["ItemRef"] == {"value": "2"}
    R._apply(t, ln, {"project": ("101", "100")}, {})
    assert "ProjectRef" not in ln
