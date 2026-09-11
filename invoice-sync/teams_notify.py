"""
teams_notify.py — POST invoice events to a Microsoft Teams channel.

Uses the modern Power Automate "Workflows" webhook (not the deprecated
Office 365 Incoming Webhook connector). The Workflow template lives in
Teams (channel → ⋯ → Workflows → "Post to a channel when a webhook
request is received") and gives back a logic.azure.com URL.

Phase 1 scope: MFD invoices only. Two event types:
  - "paid"      → Status flipped Unpaid/Partially Paid → Paid
  - "short_pay" → Status flipped Unpaid → Partially Paid

Best-effort delivery: any failure is logged at WARNING and swallowed so
the sync run still succeeds. Notion is the source of truth; if Teams
misses a message, the user still sees it in Notion.
"""
from __future__ import annotations

import datetime as dt
import logging
from typing import Any, Dict, Optional

import requests


log = logging.getLogger("automation_worker.teams_notify")


def _to_float(v: Any):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _payload(
    event_type: str,
    division: str,
    invoice_num: str,
    customer: str,
    amount: float,
    project: str,
    qbo_link: str,
    short_pay_amount: Optional[float] = None,
    line_items: Optional[list] = None,
) -> Dict[str, Any]:
    """
    Build the Adaptive Card payload Teams Workflows expects.

    The "Send webhook alerts to a channel" Workflow template wraps the request
    body in a "Post card in a chat or channel" action — that action reads the
    standard {type: message, attachments: [{contentType, content}]} envelope
    and renders the embedded Adaptive Card into the channel.

    Reference: https://adaptivecards.io/schemas/adaptive-card.json
    """
    if event_type == "paid":
        title = f"{division} Invoice Paid in Full"
        accent = "Good"   # Adaptive Cards color — renders green
        facts = [
            {"title": "Customer", "value": customer},
            {"title": "Invoice #", "value": invoice_num},
            {"title": "Project #", "value": project},
            {"title": "Amount", "value": f"${amount:,.2f}"},
        ]
    elif event_type == "short_pay":
        title = f"{division} Invoice Short-Paid"
        accent = "Warning"   # renders orange/yellow
        paid_str = f"${short_pay_amount:,.2f}" if short_pay_amount else "partial"
        facts = [
            {"title": "Customer", "value": customer},
            {"title": "Invoice #", "value": invoice_num},
            {"title": "Project #", "value": project},
            {"title": "Paid this time", "value": paid_str},
            {"title": "Remaining", "value": f"${amount:,.2f}"},
        ]
    else:
        raise ValueError(f"Unknown event_type: {event_type!r}")

    # Memo — the POSITIVE line item description(s) (the draw name, "City
    # Retainage", etc.), shown as a plain field with NO dollar. The amount is
    # already the "Amount" fact above (the net invoice total); we deliberately
    # don't repeat a figure here, and never show the gross line amount (it would
    # mislead whenever a retainage line lowers the total). Negatives/subtotals
    # are excluded by the empty/positive-amount filter.
    memo_descs = [
        str(li.get("description") or "").strip()
        for li in (line_items or [])
        if (_to_float(li.get("amount")) or 0) > 0 and str(li.get("description") or "").strip()
    ]
    if memo_descs:
        facts.append({"title": "Memo", "value": "; ".join(memo_descs)[:100]})

    body = [
        {
            "type": "TextBlock",
            "size": "Medium",
            "weight": "Bolder",
            "color": accent,
            "text": title,
            "wrap": True,
        },
        {
            "type": "FactSet",
            "facts": facts,
        },
    ]

    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": body,
    }

    return {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "contentUrl": None,
                "content": card,
            }
        ],
    }


def notify_sync_alert(
    webhook_url: str,
    *,
    runtime: str,
    title: str,
    detail: str,
    severity: str = "warning",
) -> None:
    """
    Fire-and-forget OPERATIONS alert to Teams — used when a sync run fails or
    finishes with errors. Especially important for the unattended Docker
    container, where nobody is watching the terminal.

    `runtime` identifies WHICH instance fired (e.g. 'v1.0.0 (docker)' or
    'mv1 (mac)') so alerts are unambiguous while both run during testing.
    `severity` 'error' renders red (Attention); anything else renders orange
    (Warning). Empty webhook_url → no-op. Never raises.
    """
    if not webhook_url:
        return
    accent = "Attention" if severity == "error" else "Warning"
    icon = "🛑" if severity == "error" else "⚠️"
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": [
            {
                "type": "TextBlock", "size": "Medium", "weight": "Bolder",
                "color": accent, "wrap": True,
                "text": f"{icon} Invoice Sync Alert",
            },
            {
                "type": "FactSet",
                "facts": [
                    {"title": "Runtime", "value": runtime},
                    {"title": "Status", "value": title},
                    {"title": "Detail", "value": (detail or "")[:400]},
                    {"title": "When (UTC)",
                     "value": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M:%S")},
                ],
            },
        ],
    }
    payload = {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "contentUrl": None,
            "content": card,
        }],
    }
    try:
        r = requests.post(webhook_url, json=payload, timeout=10)
        if r.status_code >= 400:
            log.warning("Teams alert rejected status=%d body=%s",
                        r.status_code, (r.text or "")[:200])
        else:
            log.info("Teams alert sent: %s — %s", runtime, title)
    except requests.RequestException as e:
        log.warning("Teams alert failed (network): %s", e)


def notify_invoice_event(
    webhook_url: str,
    *,
    event_type: str,
    division: str,
    invoice_num: str,
    customer: str,
    amount: float,
    project: str,
    qbo_link: str,
    short_pay_amount: Optional[float] = None,
    line_items: Optional[list] = None,
) -> None:
    """
    Fire-and-forget Teams notification. Never raises — failures are warned.

    `webhook_url` empty/None → no-op (lets you disable notifications by
    leaving the env var unset).
    """
    if not webhook_url:
        return
    try:
        payload = _payload(
            event_type=event_type,
            division=division,
            invoice_num=invoice_num,
            customer=customer,
            amount=amount,
            project=project,
            qbo_link=qbo_link,
            short_pay_amount=short_pay_amount,
            line_items=line_items,
        )
    except ValueError as e:
        log.warning("teams_notify: %s — skipping", e)
        return

    try:
        r = requests.post(webhook_url, json=payload, timeout=10)
        if r.status_code >= 400:
            log.warning(
                "Teams webhook %s rejected status=%d body=%s",
                event_type, r.status_code, (r.text or "")[:200],
            )
        else:
            log.info(
                "Teams notify %s sent: %s inv=%s",
                event_type, division, invoice_num,
            )
    except requests.RequestException as e:
        log.warning(
            "Teams webhook %s failed (network): %s", event_type, e,
        )
