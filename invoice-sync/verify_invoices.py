#!/usr/bin/env python3
"""
verify_invoices.py — compare QBO open invoices against Notion Invoice Trackers.

Reads QBO open invoices live, reads both Notion DBs live, compares them
invoice-by-invoice, and emits a markdown audit report. The report is what
you show stakeholders to prove the sync is keeping the trackers complete.

Usage:
    python3 verify_invoices.py                    # run + print to stdout
    python3 verify_invoices.py --out report.md    # write to file

Exit codes:
    0  perfect match (every routable QBO invoice is in Notion)
    1  drift detected (some QBO invoices missing from Notion)
    2  fatal error
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from pathlib import Path

from config import load_config
from notion_client import NotionClient
import qbo_client


PROJECT_NUM_RE = re.compile(r"\b((?:MFD|CP|RP)\d+(?:-FTW)?)\b", re.IGNORECASE)


def _load_invoice_cache(notion: NotionClient, ds_id: str) -> dict:
    """{Invoice ID → {'invoice_num', 'status', 'page_id', 'open_balance'}}"""
    cache = {}
    for page in notion.query_data_source(ds_id, page_size=100):
        page_id = page.get("id")
        props = page.get("properties") or {}

        inv_id_prop = props.get("Invoice ID") or {}
        inv_id_rich = inv_id_prop.get("rich_text") or []
        inv_id = "".join(t.get("plain_text", "") for t in inv_id_rich).strip()
        if not inv_id:
            continue

        title_prop = props.get("Invoice #") or {}
        title_arr = title_prop.get("title") or []
        invoice_num = "".join(t.get("plain_text", "") for t in title_arr).strip()

        status = (props.get("Status") or {}).get("select", {})
        status_name = status.get("name") if status else None

        open_bal = (props.get("Open balance") or {}).get("number")

        cache[inv_id] = {
            "page_id": page_id,
            "invoice_num": invoice_num,
            "status": status_name,
            "open_balance": open_bal,
        }
    return cache


def _extract_project(customer_name: str, memo: str) -> str:
    text = (customer_name or "") + " " + (memo or "")
    m = PROJECT_NUM_RE.search(text)
    return m.group(1).upper() if m else ""


def _division_for(project: str) -> str:
    if project.startswith("MFD"):
        return "MFD"
    if project.startswith("RP"):
        return "RP"
    if project.startswith("CP"):
        return "CP"
    return ""


def main() -> int:
    parser = argparse.ArgumentParser(description="QBO ↔ Notion invoice audit")
    parser.add_argument("--out", help="Write markdown report to this path")
    args = parser.parse_args()

    config = load_config()
    notion = NotionClient(
        secret=config.notion_secret,
        api_base=config.notion_api_base,
        version=config.notion_version,
    )

    print("[1/3] Authenticating to QBO and pulling open invoices…", file=sys.stderr)
    qbo_creds = qbo_client.load_qbo_credentials()
    qbo_invoices = qbo_client.query_all(qbo_creds, "Invoice", where="Balance > '0'")

    print("[2/3] Loading Notion invoice caches…", file=sys.stderr)
    res_com = _load_invoice_cache(notion, config.invoice_res_com_ds_id)
    mfd = _load_invoice_cache(notion, config.invoice_mfd_ds_id)
    notion_all = {**res_com, **mfd}  # all invoices regardless of DB

    print("[3/3] Comparing…", file=sys.stderr)
    qbo_routable_in_notion = []
    qbo_routable_missing = []
    qbo_unroutable = []
    total_open_balance_routable = 0.0
    total_open_balance_unroutable = 0.0

    for inv in qbo_invoices:
        qbo_id = str(inv.get("Id") or "")
        doc_num = str(inv.get("DocNumber") or "")
        cust = (inv.get("CustomerRef") or {}).get("name") or ""
        memo = inv.get("PrivateNote") or ""
        balance = float(inv.get("Balance") or 0)
        project = _extract_project(cust, memo)

        if not project:
            qbo_unroutable.append({"qbo_id": qbo_id, "doc_num": doc_num,
                                    "customer": cust, "memo": memo, "balance": balance})
            total_open_balance_unroutable += balance
            continue

        total_open_balance_routable += balance
        if qbo_id in notion_all:
            qbo_routable_in_notion.append({"qbo_id": qbo_id, "doc_num": doc_num,
                                            "project": project, "balance": balance})
        else:
            qbo_routable_missing.append({"qbo_id": qbo_id, "doc_num": doc_num,
                                          "project": project, "customer": cust,
                                          "memo": memo, "balance": balance})

    qbo_open_ids = {str(inv.get("Id")) for inv in qbo_invoices if inv.get("Id")}
    notion_paid_or_stale = []
    for inv_id, page in notion_all.items():
        if inv_id not in qbo_open_ids and page["status"] != "Paid":
            notion_paid_or_stale.append({"inv_id": inv_id, **page})

    # Build report
    today = dt.date.today().isoformat()
    rate = (len(qbo_routable_in_notion) / len(qbo_invoices) * 100) if qbo_invoices else 0
    routable_match_rate = (
        len(qbo_routable_in_notion) / (len(qbo_routable_in_notion) + len(qbo_routable_missing)) * 100
        if qbo_routable_in_notion or qbo_routable_missing else 0
    )

    md = []
    md.append(f"# Invoice Tracker Audit — {today}\n")
    md.append("Compares live QBO open invoices against the Notion Invoice Tracker (Res/Com) "
              "and Invoice Tracker (MFD) databases. Run anytime to re-verify completeness.\n")
    md.append("## Top-line numbers\n")
    md.append(f"| Metric | Value |")
    md.append(f"|---|---:|")
    md.append(f"| QBO open invoices (total) | {len(qbo_invoices)} |")
    md.append(f"| ↳ with project # (routable) | {len(qbo_routable_in_notion) + len(qbo_routable_missing)} |")
    md.append(f"| ↳ no project # (equipment leases / non-project AR) | {len(qbo_unroutable)} |")
    md.append(f"| Notion invoices in trackers (Res/Com + MFD) | {len(notion_all)} |")
    md.append(f"| Open balance (routable) | ${total_open_balance_routable:,.2f} |")
    md.append(f"| Open balance (unroutable) | ${total_open_balance_unroutable:,.2f} |")
    md.append(f"| **Routable match rate** | **{routable_match_rate:.1f}%** |\n")

    md.append("## Verdict\n")
    if not qbo_routable_missing:
        md.append("**PASS** — every routable QBO open invoice is present in Notion. The Notion "
                  "trackers are a complete mirror of QBO open AR (excluding equipment-lease invoices "
                  "without a project number, which are intentionally outside the project-based "
                  "collection workflow).\n")
    else:
        md.append(f"**DRIFT** — {len(qbo_routable_missing)} routable QBO invoice(s) are not in Notion. "
                  "See section below.\n")

    if qbo_routable_missing:
        md.append("## Routable invoices missing from Notion (action required)\n")
        md.append("| QBO Inv # | Project # | Customer | Open Bal |")
        md.append("|---|---|---|---:|")
        for r in qbo_routable_missing:
            md.append(f"| {r['doc_num']} | {r['project']} | {r['customer']} | ${r['balance']:,.2f} |")
        md.append("")

    if qbo_unroutable:
        md.append(f"## Unroutable QBO invoices ({len(qbo_unroutable)} — by design, NOT in Notion)\n")
        md.append("These QBO invoices have no project # in either CustomerRef.name or PrivateNote. "
                  "They're typically equipment-lease income, sub-rental fees, or other non-project AR. "
                  "They're intentionally excluded from the project-based Invoice Trackers.\n")
        md.append("| QBO Inv # | Customer | Memo | Open Bal |")
        md.append("|---|---|---|---:|")
        for r in qbo_unroutable[:20]:
            md.append(f"| {r['doc_num']} | {r['customer']} | {(r['memo'] or '')[:60]} | ${r['balance']:,.2f} |")
        if len(qbo_unroutable) > 20:
            md.append(f"| _…and {len(qbo_unroutable) - 20} more_ |  |  |  |")
        md.append("")

    if notion_paid_or_stale:
        md.append(f"## Notion-marked-open invoices not in QBO open list ({len(notion_paid_or_stale)})\n")
        md.append("These are invoices Notion still shows as open, but QBO no longer reports as open. "
                  "The next sync's flip-to-paid sweep will mark them Paid automatically.\n")

    md.append("---\n")
    md.append("_Generated by `verify_invoices.py`. Re-run anytime: "
              "`cd automation-worker && python3 verify_invoices.py --out \"$HOME/Library/Logs/Proficient/automation-worker/audit-$(date +%Y-%m-%d).md\"`_")

    report = "\n".join(md)
    if args.out:
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.out).write_text(report)
        print(f"\nReport written to {args.out}", file=sys.stderr)
    else:
        print()
        print(report)

    print(f"\nSummary: {len(qbo_routable_in_notion)}/{len(qbo_routable_in_notion) + len(qbo_routable_missing)} "
          f"routable invoices matched ({routable_match_rate:.1f}%)", file=sys.stderr)
    return 0 if not qbo_routable_missing else 1


if __name__ == "__main__":
    sys.exit(main())
