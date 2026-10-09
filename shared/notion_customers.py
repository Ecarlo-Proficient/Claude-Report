"""shared/notion_customers.py - resolve an invoice's PARENT customer (the GC / client)
from its Notion `Customer` relation, via a one-shot {page_id -> title} cache of a customer DB.

An Invoice Tracker page carries two customer fields: the `Customer` RELATION (points at the
parent customer page - the GC) and `Customer (raw)` TEXT (the project-level child, e.g.
"MFD177 - MERRITT PARK"). The client column everywhere should show the PARENT, falling back
to the raw name only when the relation doesn't resolve. Both invoice-sync (AR Aging Excel) and
the ledger (Open Invoices tab) need the same answer, so this lives in shared/ (repo rule: a file
a second tool needs moves here) - the workbook and the site then name the client identically.

Client-agnostic: it operates on the Notion page-property dicts each tool already holds, so each
keeps its own Notion client.
"""
from __future__ import annotations


def _title_text(props: dict, title_prop: str) -> str:
    """Plain text of a page's title property. Tries `title_prop`, then any title-typed
    property (customer DBs name it "Client", but stay robust if that ever changes)."""
    p = props.get(title_prop) or {}
    if p.get("type") not in ("title", None) or not p.get("title"):
        p = next((v for v in props.values() if isinstance(v, dict) and v.get("type") == "title"), p)
    return "".join(t.get("plain_text", "") for t in (p.get("title") or [])).strip()


def build_title_cache(pages, title_prop: str = "Client") -> dict:
    """{page_id -> display name} for every customer page in a customer DB."""
    cache: dict = {}
    for page in pages:
        pid = page.get("id")
        if not pid:
            continue
        name = _title_text(page.get("properties") or {}, title_prop)
        if name:
            cache[pid] = name
    return cache


def relation_title(relation_prop, cache: dict) -> str:
    """Title of the FIRST page a relation points at (e.g. an invoice's parent `Customer`),
    looked up in `cache`; "" when the relation is empty or unresolved."""
    rel = (relation_prop or {}).get("relation") or []
    if not rel:
        return ""
    return cache.get((rel[0] or {}).get("id"), "")


def _selftest() -> None:
    pages = [
        {"id": "C1", "properties": {"Client": {"type": "title", "title": [{"plain_text": "Firestone Building Co"}]}}},
        {"id": "C2", "properties": {"Name": {"type": "title", "title": [{"plain_text": "Mesquite ISD"}]}}},  # non-"Client" title
        {"id": "C3", "properties": {"Client": {"type": "title", "title": []}}},                               # blank -> skipped
    ]
    cache = build_title_cache(pages)
    assert cache == {"C1": "Firestone Building Co", "C2": "Mesquite ISD"}, cache
    inv = {"Customer": {"type": "relation", "relation": [{"id": "C1"}]}}
    assert relation_title(inv["Customer"], cache) == "Firestone Building Co"
    assert relation_title({"relation": [{"id": "NOPE"}]}, cache) == ""     # unresolved
    assert relation_title({"relation": []}, cache) == "" and relation_title(None, cache) == ""
    print("shared/notion_customers selftest OK: title cache (Client + fallback title), relation resolve, empties.")


if __name__ == "__main__":
    _selftest()
