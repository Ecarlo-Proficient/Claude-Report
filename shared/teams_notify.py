"""shared/teams_notify.py — post Adaptive Cards to a Microsoft Teams channel.

The ONE shared Teams poster (tools never import tools; invoice-sync keeps its own
tool-local teams_notify by design). Uses the modern Power Automate "Workflows"
webhook (channel -> ⋯ -> Workflows -> "Post to a channel when a webhook request
is received"), same envelope the invoice-sync poster uses.

Best-effort: an empty webhook or any network error is a no-op that returns False
and never raises, so a failed post can never break the run that called it.

Statement-reconciler use: ONE card per open vendor-month, so the bill clerk can
react ✅ to each vendor independently (a single combined message would make "done"
ambiguous). Webhook lives in the automation-qbo keychain blob (key library rule)
as TEAMS_STMT_WEBHOOK.
"""
from __future__ import annotations

from typing import Dict, List, Optional

import requests


def _envelope(body: List[dict]) -> dict:
    return {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "contentUrl": None,
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard", "version": "1.4", "body": body,
            },
        }],
    }


def post(webhook_url: str, payload: dict) -> bool:
    """POST one Adaptive Card envelope. Empty webhook or any error -> False, never
    raises."""
    if not webhook_url:
        return False
    try:
        r = requests.post(webhook_url, json=payload, timeout=10)
        return r.status_code < 400
    except requests.RequestException:
        return False


def post_statement_task(webhook_url: str, vendor: str, month: str,
                        items: Dict[str, int]) -> bool:
    """One task card for a vendor-month that still needs work. `items` is a
    label->count map (0s dropped). The clerk reacts ✅ on THIS message when this
    vendor is fully entered/printed - one message per vendor keeps that clear."""
    facts = [{"title": k, "value": str(v)} for k, v in items.items() if v]
    if not facts:
        return False
    body = [
        {"type": "TextBlock", "size": "Medium", "weight": "Bolder",
         "color": "Warning", "wrap": True, "text": f"🔧 {vendor} — {month}"},
        {"type": "FactSet", "facts": facts},
        {"type": "TextBlock", "size": "Small", "isSubtle": True, "wrap": True,
         "text": "React ✅ on this message once this vendor is fully entered / printed."},
    ]
    return post(webhook_url, _envelope(body))


def post_month_clean(webhook_url: str, vendor: str, month: str) -> bool:
    """A green 'this vendor-month is clean' card after a refresh finds nothing left
    - the owner's cue to rename the folder '<MM-YYYY> DONE'."""
    body = [
        {"type": "TextBlock", "size": "Medium", "weight": "Bolder",
         "color": "Good", "wrap": True, "text": f"✅ {vendor} — {month} is clean"},
        {"type": "TextBlock", "wrap": True,
         "text": "Nothing left to fix or print. Rename the month folder "
                 f"'{month} DONE' to finalize it."},
    ]
    return post(webhook_url, _envelope(body))
