#!/usr/bin/env python3
"""
load_customers.py — land the Notion "Customer List" (CRM) into the ledger.

WHAT IT DOES
Reads the Notion Customer List data source (read-only) and fills two tables:
  * customer     — one row per client/lead: identity, current pipeline stage,
                   and Notion's own "Created by" / "Last edited by" system fields
                   (who sourced it / who worked it last — the honest per-rep
                   attribution, no manual Owner property to maintain).
  * sales_touch  — one row per "History of interactions" line in the page body
                   (the outreach touch log), with the date parsed when present.

This is the CRM half of "own the spine": the customer/pipeline lives in the ledger;
Notion is just the feed. It joins the job spine on nothing yet — leads become
projects downstream — but it puts sales activity in the same database as the WIP,
so "what has the outreach rep done" is a query, not a spreadsheet.

SAFETY
  * READ-ONLY on Notion (query + block reads only; never writes a page).
  * Idempotent: each run FULL-REPLACES source='notion_customer_list' rows
    (and their sales_touch), so the ledger mirrors the current list.
  * --dry-run pulls + reports; writes nothing to the ledger.
  * --selftest runs the whole parse+load OFFLINE on a throwaway DB (no Notion).

SETUP (one-time)
  * ACB_CUSTOMER_LIST_DS_ID = the Customer List data-source id (set in machine.env).
  * The Notion integration (Keychain proficient-automation-worker/notion, the same
    token sync_actions.py uses) must have the Customer List shared with it.

USAGE
  python3 ledger/load_customers.py --selftest
  python3 ledger/load_customers.py --dry-run --show 12
  python3 ledger/load_customers.py
  python3 ledger/load_customers.py --all-notes --show 12
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sqlite3
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths  # noqa: E402

HERE = Path(__file__).resolve().parent
SCHEMA_SQL = HERE / "schema.sql"
DEFAULT_DB = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3",
)
SOURCE = "notion_customer_list"
NOTES_HEADING = "history of interactions"   # the body section that holds the touch log


# ── property extraction (Notion API 2025-09-03 page object) ─────────────────
def _plain(rich) -> str | None:
    return ("".join(x.get("plain_text", "") for x in (rich or [])).strip()) or None


def prop(props: dict, name: str):
    """Value of a Customer List property, normalized to a scalar/str."""
    p = props.get(name) or {}
    t = p.get("type")
    if t in ("title", "rich_text"):
        return _plain(p.get(t))
    if t == "email":
        return p.get("email") or None
    if t == "phone_number":
        return p.get("phone_number") or None
    if t == "select":
        return ((p.get("select") or {}) or {}).get("name")
    if t == "status":
        return ((p.get("status") or {}) or {}).get("name")
    if t == "date":
        return (p.get("date") or {}).get("start")
    if t == "multi_select":
        names = [o.get("name", "") for o in (p.get("multi_select") or [])]
        return ", ".join(n for n in names if n) or None
    return None


def prop_user(props: dict, name: str, page_fallback=None):
    """Display name of a created_by / last_edited_by property (Notion embeds the
    person's name in the property value); fall back to the page-level user id."""
    p = props.get(name) or {}
    u = p.get(p.get("type"))
    if isinstance(u, dict):
        return u.get("name") or u.get("id")
    if isinstance(page_fallback, dict):
        return page_fallback.get("name") or page_fallback.get("id")
    return None


_DMY = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{2,4})")
_MDY = re.compile(r"([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})")


def parse_touch_date(note: str) -> str | None:
    """Best-effort ISO date from a touch line ('… 07/15/26' or 'Introduction July 10, 2026')."""
    m = _DMY.search(note)
    if m:
        mo, da, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        yr += 2000 if yr < 100 else 0
        try:
            return dt.date(yr, mo, da).isoformat()
        except ValueError:
            return None
    m = _MDY.search(note)
    if m:
        for fmt in ("%B %d %Y", "%b %d %Y"):
            try:
                return dt.datetime.strptime(f"{m.group(1)} {int(m.group(2))} {m.group(3)}", fmt).date().isoformat()
            except ValueError:
                continue
    return None


def touches_from_blocks(blocks: list) -> list[str]:
    """The bullet lines under the 'History of interactions' heading, in order."""
    section, out = "", []
    for b in blocks:
        t = b.get("type", "")
        if t.startswith("heading_"):
            section = (_plain(b.get(t, {}).get("rich_text")) or "").lower()
        elif t in ("bulleted_list_item", "numbered_list_item") and section.startswith(NOTES_HEADING):
            note = _plain(b.get(t, {}).get("rich_text"))
            if note:
                out.append(note)
    return out


def customer_from_page(page: dict) -> dict:
    props = page.get("properties") or {}
    return {
        "customer_key": (page.get("id") or "").replace("-", ""),
        "name": prop(props, "Client") or "(unnamed)",
        "division": prop(props, "Division"),
        "main_status": prop(props, "Main Status"),
        "sales_status": prop(props, "Sales Status"),
        "last_contacted": prop(props, "Last Contacted"),
        "follow_up_date": prop(props, "Follow Up Date"),
        "referral": prop(props, "Referral"),
        "primary_contact": prop(props, "Primary Contact"),
        "primary_email": prop(props, "Primary Email"),
        "primary_phone": prop(props, "Primary Phone"),
        "created_by": prop_user(props, "Created by", page.get("created_by")),
        "last_edited_by": prop_user(props, "Last edited by", page.get("last_edited_by")),
        "last_edited_time": page.get("last_edited_time"),
        "notion_url": page.get("url"),
    }


# ── pull ────────────────────────────────────────────────────────────────────
def pull(ds_id: str, all_notes: bool, limit: int | None):
    """Yield (customer_dict, [touch_note, ...]) for each Customer List page.
    Notes bodies are fetched for worked pages (status past 'Lead') by default —
    pure 'Lead' rows rarely carry a touch log; --all-notes fetches every body."""
    from shared.notion_client import NotionClient
    nc = NotionClient()
    n = 0
    for page in nc.query_data_source(ds_id):
        cust = customer_from_page(page)
        worked = (cust["sales_status"] or "Lead") not in ("Lead", "Follow up")
        notes: list[str] = []
        if all_notes or worked:
            notes = touches_from_blocks(list(nc.block_children(page["id"])))
        cust["n_touches"] = len(notes)
        yield cust, notes
        n += 1
        if limit and n >= limit:
            return


# ── write ─────────────────────────────────────────────────────────────────
CUST_COLS = ["customer_key", "name", "division", "main_status", "sales_status",
             "last_contacted", "follow_up_date", "referral", "primary_contact",
             "primary_email", "primary_phone", "created_by", "last_edited_by",
             "last_edited_time", "n_touches", "notion_url", "source", "loaded_at"]


def write(db_path: Path, records: list[tuple[dict, list[str]]]):
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(db_path)
    con.execute("PRAGMA foreign_keys = ON;")
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
    now = dt.datetime.now().isoformat(timespec="seconds")
    # full replace: this feed mirrors the current Customer List
    con.execute("DELETE FROM sales_touch WHERE customer_key IN "
                "(SELECT customer_key FROM customer WHERE source = ?)", (SOURCE,))
    con.execute("DELETE FROM customer WHERE source = ?", (SOURCE,))
    ph = ", ".join(f":{c}" for c in CUST_COLS)
    for cust, notes in records:
        cust = {**cust, "n_touches": len(notes), "source": SOURCE, "loaded_at": now}
        con.execute(f"INSERT INTO customer ({', '.join(CUST_COLS)}) VALUES ({ph})",
                    {c: cust.get(c) for c in CUST_COLS})
        for seq, note in enumerate(notes, 1):
            con.execute("INSERT INTO sales_touch (customer_key, seq, touch_date, note) VALUES (?,?,?,?)",
                        (cust["customer_key"], seq, parse_touch_date(note), note))
    con.commit()
    n_touch = con.execute("SELECT COUNT(*) FROM sales_touch").fetchone()[0]
    con.close()
    return len(records), n_touch


def summarize(records: list[tuple[dict, list[str]]], show: int):
    from collections import Counter
    pipe = Counter((c["sales_status"] or "(none)") for c, _ in records)
    reps = Counter((c["last_edited_by"] or "(unknown)") for c, _ in records)
    order = ["Lead", "Follow up", "Contacted", "Interested", "No response", "Closed - Won", "Closed - Lost", "(none)"]
    print("\nPipeline:")
    for st in sorted(pipe, key=lambda s: order.index(s) if s in order else 99):
        print(f"  {st:<16} {pipe[st]:>4}")
    print("\nWorked-by (last editor):")
    for rep, n in reps.most_common(8):
        print(f"  {rep[:28]:<28} {n:>4}")
    if show > 0:
        warm = [(c, t) for c, t in records if c["sales_status"] == "Interested"]
        print(f"\nInterested — warm accounts ({len(warm)}), top {show}:")
        for c, notes in warm[:show]:
            print(f"  {c['name'][:34]:<34} {(c['division'] or ''):<12} touches:{len(notes)} "
                  f"last:{c['last_contacted'] or '—'}")
            for note in notes[-3:]:
                print(f"      · {note[:76]}")


# ── offline selftest ────────────────────────────────────────────────────────
def selftest():
    import tempfile
    pages = [
        {"id": "aaaa-1", "url": "u1", "last_edited_time": "2026-08-07T20:00:00Z",
         "created_by": {"name": "Sourcer"}, "last_edited_by": {"name": "Rep One"},
         "properties": {
             "Client": {"type": "title", "title": [{"plain_text": "DOOLEY MACK"}]},
             "Division": {"type": "multi_select", "multi_select": [{"name": "Commercial"}]},
             "Sales Status": {"type": "status", "status": {"name": "Interested"}},
             "Main Status": {"type": "select", "select": {"name": "Qualified"}},
             "Last Contacted": {"type": "date", "date": {"start": "2026-07-10"}},
             "Primary Email": {"type": "email", "email": "p@dm.com"},
             "Created by": {"type": "created_by", "created_by": {"name": "Sourcer"}},
             "Last edited by": {"type": "last_edited_by", "last_edited_by": {"name": "Rep One"}},
         }},
        {"id": "bbbb-2", "url": "u2", "last_edited_time": "2026-08-01T00:00:00Z",
         "properties": {
             "Client": {"type": "title", "title": [{"plain_text": "COLD LEAD CO"}]},
             "Sales Status": {"type": "status", "status": {"name": "Lead"}},
         }},
    ]
    blocks = [
        {"type": "heading_2", "heading_2": {"rich_text": [{"plain_text": "Background info"}]}},
        {"type": "heading_2", "heading_2": {"rich_text": [{"plain_text": "History of interactions"}]}},
        {"type": "bulleted_list_item", "bulleted_list_item": {"rich_text": [{"plain_text": "Introduction July 10, 2026"}]}},
        {"type": "bulleted_list_item", "bulleted_list_item": {"rich_text": [{"plain_text": "Quote sent 07/15/26"}]}},
        {"type": "heading_2", "heading_2": {"rich_text": [{"plain_text": "Action items"}]}},
        {"type": "to_do", "to_do": {"rich_text": [{"plain_text": "ignore me"}]}},
    ]
    # extraction
    c0 = customer_from_page(pages[0])
    assert c0["name"] == "DOOLEY MACK" and c0["sales_status"] == "Interested", c0
    assert c0["last_edited_by"] == "Rep One" and c0["created_by"] == "Sourcer", c0
    assert c0["customer_key"] == "aaaa1", c0
    touches = touches_from_blocks(blocks)
    assert touches == ["Introduction July 10, 2026", "Quote sent 07/15/26"], touches
    assert parse_touch_date("Quote sent 07/15/26") == "2026-07-15", parse_touch_date("Quote sent 07/15/26")
    assert parse_touch_date("Introduction July 10, 2026") == "2026-07-10"
    assert parse_touch_date("no date here") is None
    # load into a throwaway DB
    recs = [(customer_from_page(pages[0]), touches), (customer_from_page(pages[1]), [])]
    recs[0][0]["n_touches"] = 2
    with tempfile.TemporaryDirectory() as d:
        db = Path(d) / "t.sqlite3"
        n_cust, n_touch = write(db, recs)
        con = sqlite3.connect(db)
        assert n_cust == 2 and n_touch == 2, (n_cust, n_touch)
        # idempotent: a second identical load does not duplicate
        write(db, recs)
        again = con.execute("SELECT COUNT(*) FROM customer").fetchone()[0]
        assert again == 2, again
        pipe = dict(con.execute("SELECT sales_status, customers FROM v_sales_pipeline"))
        assert pipe.get("Interested") == 1 and pipe.get("Lead") == 1, pipe
        rep = con.execute("SELECT interested FROM v_sales_by_rep WHERE rep = 'Rep One'").fetchone()[0]
        assert rep == 1, rep
        d0 = con.execute("SELECT touch_date FROM sales_touch WHERE note LIKE 'Quote%'").fetchone()[0]
        assert d0 == "2026-07-15", d0
        con.close()
    print("selftest OK — extraction, date parse, idempotent load, and views all pass.")


def main():
    ap = argparse.ArgumentParser(description="Load the Notion Customer List (CRM) into the ledger.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB, help="SQLite ledger to write.")
    ap.add_argument("--ds", default=None, help="Customer List data-source id (default: ACB_CUSTOMER_LIST_DS_ID).")
    ap.add_argument("--dry-run", action="store_true", help="Pull + report; write nothing.")
    ap.add_argument("--all-notes", action="store_true", help="Fetch the body of EVERY page (default: only worked ones).")
    ap.add_argument("--limit", type=int, default=None, help="Stop after N pages (testing).")
    ap.add_argument("--show", type=int, default=0, help="Print N warm (Interested) accounts with their last touches.")
    ap.add_argument("--selftest", action="store_true", help="Run the offline parse+load test (no Notion) and exit.")
    args = ap.parse_args()

    if args.selftest:
        selftest()
        return

    ds_id = args.ds or paths.get("ACB_CUSTOMER_LIST_DS_ID")
    if not ds_id:
        sys.exit("Set ACB_CUSTOMER_LIST_DS_ID (the Customer List data-source id) in machine.env, or pass --ds.")

    print(f"Reading Notion Customer List (read-only): {ds_id}")
    records = list(pull(ds_id, args.all_notes, args.limit))
    n_touch = sum(len(t) for _, t in records)
    print(f"Pulled {len(records)} customers, {n_touch} interaction-log lines.")
    summarize(records, args.show)

    if args.dry_run:
        print("\n--dry-run: nothing written.")
        return
    n_cust, n_touch = write(args.db, records)
    print(f"\nWrote {n_cust} customers + {n_touch} touches -> {args.db}")


if __name__ == "__main__":
    main()
