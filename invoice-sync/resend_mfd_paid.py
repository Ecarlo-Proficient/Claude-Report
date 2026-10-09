#!/usr/bin/env python3
"""
resend_mfd_paid.py — re-fire MFD "paid" Teams cards for specific invoices.

Recovery tool for when the Teams flow connection was down (e.g. an org password
change de-authenticated it) and paid notifications were lost. The sync fires the
paid card only once, at the open→Paid transition, and never retries — so a lost
card has to be re-sent by hand. This pulls each invoice live from QBO and posts
the same card the sync would have.

Usage (on the Mac, where the webhook lives in Keychain):
    python3 resend_mfd_paid.py --dry-run 34202 34234    # preview, post nothing
    python3 resend_mfd_paid.py 34202 34234              # post to the MFD channel

Exit codes: 0 ok · 1 an invoice wasn't found · 2 webhook not configured
"""
import argparse
import json
import sys

from config import _get_teams_webhook
import qbo_client
from invoice_sync import (
    _extract_project_num, _positive_line_items, _resolve_parent_customer, DIVISION_MFD,
)
from teams_notify import notify_invoice_event, _payload


def _fetch(creds, doc_number: str):
    safe = str(doc_number).replace("'", "")
    data = qbo_client.query(creds, f"SELECT * FROM Invoice WHERE DocNumber = '{safe}'")
    rows = data.get("QueryResponse", {}).get("Invoice", [])
    return rows[0] if rows else None


def _event_for(inv: dict, doc_number: str, hierarchy: dict, company_id: str) -> dict:
    raw = (inv.get("CustomerRef") or {}).get("name", "").strip()
    # PARENT customer (GC / developer), not the project-# sub-customer.
    customer = _resolve_parent_customer(inv, hierarchy) or raw
    project = (_extract_project_num(raw)
               or _extract_project_num(inv.get("PrivateNote") or ""))
    return dict(
        event_type="paid",
        division=DIVISION_MFD,
        invoice_num=doc_number,
        customer=customer or "(unknown customer)",
        amount=float(inv.get("TotalAmt") or 0.0),
        project=project or "",
        qbo_link=qbo_client.invoice_deep_link(company_id, str(inv.get("Id") or "")),
        line_items=_positive_line_items(inv),
    )


def main() -> int:
    ap = argparse.ArgumentParser(description="Resend MFD paid Teams cards by DocNumber")
    ap.add_argument("doc_numbers", nargs="+", help="Invoice DocNumbers, e.g. 34202 34234")
    ap.add_argument("--dry-run", action="store_true", help="Preview only — post nothing")
    args = ap.parse_args()

    webhook = _get_teams_webhook()
    if not webhook and not args.dry_run:
        print("FATAL: Teams webhook not configured (setup_keychain.py --teams).",
              file=sys.stderr)
        return 2
    print(f"Webhook resolved: {'YES' if webhook else 'NO'}"
          + (f" ({len(webhook)} chars)" if webhook else ""))

    creds = qbo_client.load_qbo_credentials()
    print("Loading QBO customer hierarchy (to resolve parent customers)…")
    hierarchy = qbo_client.fetch_customer_hierarchy(creds)
    rc = 0
    for dn in args.doc_numbers:
        inv = _fetch(creds, dn)
        if not inv:
            print(f"  {dn}: NOT found in QBO — skipped")
            rc = 1
            continue
        event = _event_for(inv, dn, hierarchy, creds.company_id)
        billed = ", ".join(f"{li['description']} +${li['amount']:,.2f}"
                           for li in event["line_items"]) or "(no positive lines)"
        if args.dry_run:
            print(f"  {dn}: would post → customer='{event['customer']}' "
                  f"project={event['project']} amount=${event['amount']:,.2f}")
            print(f"       billed: {billed}")
            print(json.dumps(_payload(**event), indent=2))
        else:
            notify_invoice_event(webhook, **event)
            print(f"  {dn}: posted → {event['customer']} ${event['amount']:,.2f} | {billed}")
    return rc


if __name__ == "__main__":
    sys.exit(main())
