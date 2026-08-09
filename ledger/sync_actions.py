#!/usr/bin/env python3
"""
sync_actions.py — mirror ledger action items to Notion (the folder-memory link).

MVP scope: **draws ready to turn in** (funded + every vendor paid + every
unconditional waiver in hand). For each, upsert a page in the Notion "Ledger
Actions" database keyed by a stable Action Key, read its Status back, and record
the page URL + status in the local `action` table. The dashboard shows the link +
status; the work / thread / done lives in the Notion page. The ledger stays the
RADAR.

SAFETY
  * Read-only on the ledger except the `action` table it writes.
  * Notion writes are scoped to the Actions DB (create/update pages only).
  * --dry-run computes + prints the ready draws and writes NOTHING (no Notion).

SETUP (one-time)
  * ACB_ACTIONS_DS_ID = the "Ledger Actions" data-source id (set in machine.env).
  * Notion secret in Keychain (proficient-automation-worker/notion) or NOTION_SECRET.

USAGE
  python3 ledger/sync_actions.py --dry-run
  python3 ledger/sync_actions.py
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
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


def _wk(mi, vendor, bill):
    return hashlib.sha1(f"{mi or ''}|{vendor or ''}|{bill or ''}".encode()).hexdigest()[:16]


def ready_draws(con) -> list:
    """Draws that are funded + all vendors paid + all unconditional waivers in hand.
    RP excluded (residential bills at completion/milestones, not formal draws)."""
    rows = con.execute(
        "SELECT matched_invoice mi, project_no, vendor, bill_ref, MAX(bill_total) amt, "
        "MAX(gc_paid_date) gc, MAX(pay_date) pd FROM ap_bill_line "
        "WHERE matched_invoice IS NOT NULL AND matched_invoice <> '' "
        "AND COALESCE(project_no,'') NOT LIKE 'RP%' AND matched_invoice NOT LIKE '%— RP%' "
        "GROUP BY matched_invoice, vendor, bill_ref").fetchall()
    wmap = {w[0]: w[1] for w in con.execute("SELECT waiver_key, received FROM waiver")}
    draws: dict = {}
    for r in rows:
        d = draws.setdefault(r["mi"], {"mi": r["mi"], "project_no": r["project_no"], "bills": []})
        d["bills"].append({"gc": r["gc"], "pd": r["pd"], "amt": r["amt"] or 0,
                           "waiver": bool(wmap.get(_wk(r["mi"], r["vendor"], r["bill_ref"]), 0))})
    out = []
    for mi, d in draws.items():
        b = d["bills"]
        if b and any(x["gc"] for x in b) and all(x["pd"] for x in b) and all(x["waiver"] for x in b):
            inv = (mi or "").split("—")[0].strip()
            out.append({"action_key": f"draw:{inv}", "mi": mi, "project_no": d["project_no"],
                        "label": (mi or "").split("\n")[0].strip(),
                        "amount": round(sum(x["amt"] for x in b)), "n": len(b)})
    return out


def _notion_props(a: dict) -> dict:
    return {
        "Name": {"title": [{"text": {"content": ("Turn in draw: " + a["label"])[:200]}}]},
        "Key": {"rich_text": [{"text": {"content": a["action_key"]}}]},
        "Type": {"select": {"name": "draw"}},
        "Project": {"rich_text": [{"text": {"content": a["project_no"] or ""}}]},
        "Amount": {"number": a["amount"]},
    }


def _notion_body(a: dict) -> list:
    return [
        {"object": "block", "type": "callout", "callout": {
            "rich_text": [{"text": {"content": f"{a['label']} — {a['n']} bills, "
                                    "all paid, all unconditional waivers in hand. Ready to turn in."}}],
            "icon": {"emoji": "✅"}}},
        {"object": "block", "type": "to_do", "to_do": {
            "rich_text": [{"text": {"content": "Turn the waiver packet in to the GC to unlock the next draw"}}]}},
    ]


def sync(db_path: Path, dry_run: bool):
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    con.executescript(SCHEMA_SQL.read_text(encoding="utf-8"))
    ready = ready_draws(con)
    print(f"Draws ready to turn in: {len(ready)}")
    for a in ready:
        print(f"  {a['action_key']:<14} {a['label'][:52]:<52} ${a['amount']:>10,} ({a['n']} bills)")
    if dry_run:
        print("\n--dry-run: nothing written (no Notion).")
        con.close()
        return
    if not ready:
        con.close()
        return

    ds = paths.get("ACB_ACTIONS_DS_ID")
    if not ds:
        con.close()
        sys.exit("Set ACB_ACTIONS_DS_ID (the 'Ledger Actions' data-source id) in machine.env first.")
    from shared.notion_client import NotionClient
    nc = NotionClient()
    now = dt.datetime.now().isoformat(timespec="seconds")
    n_new = n_upd = 0
    for a in ready:
        page = nc.query_by_property(ds, "Key", "rich_text", a["action_key"])
        if page:
            nc.update_page(page["id"], _notion_props(a))
            status = (((page.get("properties") or {}).get("Status") or {}).get("select") or {}).get("name") or "Open"
            page_id, url = page["id"], page.get("url")
            n_upd += 1
        else:
            props = _notion_props(a)
            props["Status"] = {"select": {"name": "Open"}}
            created = nc.create_page(ds, props, children=_notion_body(a))
            status, page_id, url = "Open", created["id"], created.get("url")
            n_new += 1
        con.execute(
            "INSERT INTO action (action_key, type, project_no, title, amount, status, "
            "notion_page_id, notion_url, synced_at) VALUES (?,?,?,?,?,?,?,?,?) "
            "ON CONFLICT(action_key) DO UPDATE SET status=excluded.status, "
            "notion_page_id=excluded.notion_page_id, notion_url=excluded.notion_url, "
            "title=excluded.title, amount=excluded.amount, synced_at=excluded.synced_at",
            (a["action_key"], "draw", a["project_no"], a["label"], a["amount"], status,
             page_id, url, now))
    con.commit()
    con.close()
    print(f"\nNotion synced: {n_new} created, {n_upd} updated. Recorded in the ledger `action` table.")


def main():
    ap = argparse.ArgumentParser(description="Mirror ledger action items (draws ready) to Notion.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--dry-run", action="store_true", help="Compute + print; write nothing (no Notion).")
    args = ap.parse_args()
    sync(args.db, args.dry_run)


if __name__ == "__main__":
    main()
