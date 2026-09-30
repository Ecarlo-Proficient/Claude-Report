"""shared/teams_notify.py — post Adaptive Cards to a Microsoft Teams channel.

The ONE shared Teams poster (tools never import tools; invoice-sync keeps its own
tool-local teams_notify by design). Uses the modern Power Automate "Workflows"
webhook (channel -> ⋯ -> Workflows -> "Post to a channel when a webhook request
is received"), same envelope the invoice-sync poster uses.

Best-effort: an empty webhook or any network error is a no-op that returns False
and never raises, so a failed post can never break the run that called it.

Statement-reconciler use: ONE digest card per run that points at the Notion
"Vendor Statements" board, where each vendor-month is a live checklist page
(2026-09-29 - the old card-per-vendor + ✅ reaction could not be read back or
rewritten by a webhook). Webhook lives in the automation-qbo keychain blob (key library rule)
as TEAMS_STMT_WEBHOOK.
"""
from __future__ import annotations

from typing import List

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


def post_statement_digest(webhook_url: str, entries: List[dict], board_url: str = "") -> bool:
    """ONE compact card per run: the vendors NOT ready for a pay run (a line each,
    linked to its Notion page) and a count of the ones whose statement bills are
    all entered, plus a button to the board. The work happens on the Notion page;
    Teams is only the nudge. entries = [{name, status, standing, url}]."""
    if not entries:
        return False
    blocked = [e for e in entries if e.get("status") != "All entered"]
    ready = len(entries) - len(blocked)

    def _name(e: dict) -> str:
        return f"[{e['name']}]({e['url']})" if e.get("url") else e["name"]

    lines = [f"{_name(e)} · {e.get('standing', e.get('status', ''))}" for e in blocked]
    if ready:
        lines.append(f"✅ {ready} vendor(s): every statement bill entered")
    body = [{"type": "TextBlock", "weight": "Bolder", "wrap": True,
             "text": (f"Vendor statements · {len(blocked)} not ready to pay"
                      if blocked else "Vendor statements · every bill entered")},
            {"type": "TextBlock", "wrap": True, "spacing": "Small",
             "text": "\n\n".join(lines)}]
    payload = _envelope(body)
    if board_url:
        payload["attachments"][0]["content"]["actions"] = [
            {"type": "Action.OpenUrl", "title": "Open the board", "url": board_url}]
    return post(webhook_url, payload)
