"""shared/lien_status.py - resolve an invoice's Notion Lien Tracker status.

Two tools need the SAME answer and must never drift: the ledger (`load_invoices.py`,
the dashboard's Open Invoices Lien column) and invoice-sync (the AR Aging Excel's
"Lien status" column). Given an invoice page's `Lien` relation and a one-shot index
of the Lien Tracker, this returns the lien Status to show. Kept in shared/ because a
file a second tool needs moves here (repo rule) - so the workbook and the site agree.

The Lien Tracker (Notion) is manually maintained; its `Status` is a Notion *status*
property (values below). This is a DIFFERENT thing from the computed Texas notice
CLOCK (`shared/lien_clock.py`): the clock is a deadline, this is what the owner has
actually done about it (Not started -> Ready to Mail -> Mailed -> Lien filed).

Client-agnostic on purpose: it operates on the Notion page-property dicts both tools
already hold, so each keeps its own Notion client (ledger: shared/notion_client;
invoice-sync: its tool-local one).
"""
from __future__ import annotations

# Lien Tracker Status values, MOST-ESCALATED FIRST. When an invoice relates to more
# than one lien row, the most-escalated status is the one worth showing.
PRIORITY = ["Lien", "Mailed", "Ready to Mail", "In progress", "Ready to Review",
            "Not started", "Did Not Send", "Paid", "Closed"]

# Short labels for a narrow column (Excel / the dashboard grid).
SHORT = {
    "Lien": "Lien filed", "Mailed": "Mailed", "Ready to Mail": "Ready to mail",
    "In progress": "In progress", "Ready to Review": "Review",
    "Not started": "Not started", "Did Not Send": "Skipped",
    "Paid": "Paid", "Closed": "Closed",
}


def _status_name(prop) -> str | None:
    """Name of a Notion 'status' (or 'select') property dict, or None."""
    if not prop:
        return None
    t = prop.get("type")
    if t == "status":
        return (prop.get("status") or {}).get("name")
    if t == "select":
        return (prop.get("select") or {}).get("name")
    return None


def _multi_first(prop) -> str | None:
    """First value of a Notion 'multi_select' property (e.g. Notice Type), or None."""
    if not prop or prop.get("type") != "multi_select":
        return None
    names = [o.get("name") for o in (prop.get("multi_select") or [])]
    return names[0] if names else None


def relation_ids(props: dict, name: str) -> list:
    """Page ids a relation property points at (e.g. an invoice's `Lien` links)."""
    p = props.get(name) or {}
    if p.get("type") != "relation":
        return []
    return [r.get("id") for r in (p.get("relation") or []) if r.get("id")]


def index_from_pages(pages, status_prop: str = "Status",
                     notice_prop: str = "Notice Type") -> dict:
    """One pass over Lien Tracker pages -> {page_id: {status, notice}}."""
    idx: dict = {}
    for page in pages:
        pid = page.get("id")
        if not pid:
            continue
        props = page.get("properties", {})
        idx[pid] = {"status": _status_name(props.get(status_prop)),
                    "notice": _multi_first(props.get(notice_prop))}
    return idx


def pick(relation_id_list, index: dict) -> tuple:
    """Most-escalated (status, notice) among an invoice's related liens; (None, None)
    if it relates to none we know."""
    hits = [index[i] for i in relation_id_list if i in index]
    if not hits:
        return (None, None)

    def rank(h):
        s = h.get("status")
        return PRIORITY.index(s) if s in PRIORITY else len(PRIORITY)

    best = min(hits, key=rank)
    return (best.get("status"), best.get("notice"))


def for_invoice(invoice_props: dict, index: dict, relation_name: str = "Lien") -> tuple:
    """(status, notice) for one invoice page given the Lien Tracker index."""
    return pick(relation_ids(invoice_props, relation_name), index)


def short(status) -> str | None:
    """Compact label for a status (falls back to the raw value)."""
    return None if not status else SHORT.get(status, status)


def _selftest() -> None:
    liens = [
        {"id": "L1", "properties": {"Status": {"type": "status", "status": {"name": "Mailed"}},
                                    "Notice Type": {"type": "multi_select", "multi_select": [{"name": "RP Notice"}]}}},
        {"id": "L2", "properties": {"Status": {"type": "status", "status": {"name": "Lien"}},
                                    "Notice Type": {"type": "multi_select", "multi_select": [{"name": "Affidavit of Lien Claimed"}]}}},
        {"id": "L3", "properties": {"Status": {"type": "status", "status": {"name": "Not started"}},
                                    "Notice Type": {"type": "multi_select", "multi_select": []}}},
    ]
    idx = index_from_pages(liens)
    assert idx["L1"] == {"status": "Mailed", "notice": "RP Notice"}, idx["L1"]

    inv = {"Lien": {"type": "relation", "relation": [{"id": "L1"}, {"id": "L2"}]}}
    assert for_invoice(inv, idx) == ("Lien", "Affidavit of Lien Claimed")   # most-escalated wins
    one = {"Lien": {"type": "relation", "relation": [{"id": "L3"}]}}
    assert for_invoice(one, idx) == ("Not started", None)
    none = {"Lien": {"type": "relation", "relation": []}}
    assert for_invoice(none, idx) == (None, None)
    assert relation_ids({"Lien": {"type": "relation", "relation": [{"id": "X"}]}}, "Lien") == ["X"]
    assert short("Lien") == "Lien filed" and short(None) is None
    print("shared/lien_status selftest OK: index, most-escalated pick, empty, short labels.")


if __name__ == "__main__":
    _selftest()
