#!/usr/bin/env python3
"""
test_teams_webhook.py — verify the Teams webhook wiring.

Two modes:
    --dry-run   Resolve the webhook (report configured / not, redacted) and print
                the exact Adaptive Card that WOULD be sent. Posts NOTHING to the
                channel. Use this to confirm wiring without spamming Teams.
    (default)   Actually POST a fake event to the channel.

Usage:
    python3 test_teams_webhook.py --dry-run          # check + preview, no post
    python3 test_teams_webhook.py --dry-run --short  # preview a short-pay card
    python3 test_teams_webhook.py --paid             # POST a fake Paid event
    python3 test_teams_webhook.py --short            # POST a fake Short-Pay event

Exit codes:
    0 — dry-run completed, or message accepted by webhook (HTTP 2xx)
    1 — webhook returned 4xx/5xx or network error
    2 — webhook not configured (live mode only)
"""
import argparse
import json
import sys

from dotenv import load_dotenv
load_dotenv()  # pick up .env in CWD (Linux/Docker fallback)

from config import _get_teams_webhook       # noqa: E402
from teams_notify import _payload, notify_invoice_event  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Test the Teams webhook wiring")
    parser.add_argument("--paid", action="store_true", help="Paid event (default)")
    parser.add_argument("--short", action="store_true", help="Short-Pay event")
    parser.add_argument("--dry-run", action="store_true",
                        help="Resolve + preview the card but DO NOT post to the channel")
    args = parser.parse_args()

    # Same resolution as the live sync (Keychain-first on Mac, env on Linux/Docker).
    url = _get_teams_webhook()
    configured = bool(url)
    print(f"Webhook resolved: {'YES' if configured else 'NO'}"
          + (f"  ({len(url)} chars, redacted)" if configured else ""))

    event_type = "short_pay" if args.short else "paid"

    # Sample event — includes positive AND negative line items so the dry-run
    # shows that only positive lines render in the 'Billed' section.
    event = dict(
        event_type=event_type,
        division="MFD",
        invoice_num="TEST-99999",
        customer="TEST CUSTOMER (webhook wiring test)",
        amount=8995.67,
        project="MFD999",
        qbo_link="https://qbo.intuit.com/app/login?pagereq=invoice%3FtxnId%3D99999&deeplinkcompanyid=TESTREALM",
        line_items=[
            {"description": "Draw 5 - Slab & Foundation", "amount": 12345.67},
            {"description": "Retainage Billed", "amount": 650.00},
            {"description": "Less prior billing", "amount": -4000.00},
        ],
    )
    if event_type == "short_pay":
        event["short_pay_amount"] = 5000.00

    if args.dry_run:
        payload = _payload(**event)
        print(f"\nDRY RUN — nothing posted. The {event_type.upper()} card that "
              f"WOULD be sent:\n")
        print(json.dumps(payload, indent=2))
        if not configured:
            print("\nNOTE: webhook is NOT configured — a real run would silently "
                  "skip posting (this is why nothing reached the channel). Store it "
                  "with `python3 setup_keychain.py --teams`.")
        else:
            print("\nWebhook is configured and the card builds. Run without "
                  "--dry-run to post it for real.")
        return 0

    # ---- live post ----
    if not configured:
        print("FATAL: Teams webhook not configured. Store it with "
              "`python3 setup_keychain.py --teams` (Mac) or set "
              "TEAMS_WEBHOOK_MFD_PAID (Linux/Docker).", file=sys.stderr)
        return 2

    print(f"Posting fake {event_type.upper()} event to the Teams channel…")
    notify_invoice_event(url, **event)
    print("Done. Check the Teams channel. If nothing arrived, see the WARNING "
          "line above and confirm the Workflow is enabled in Teams.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
