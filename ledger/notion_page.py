"""notion_page.py - one Notion page, whole, for the ledger's side panel.

`GET /api/invoice/notion?url=<page url>` renders an Invoice Tracker page the way the owner
would see it in Notion - EVERY property, the page body (paragraphs, headings, bullets, to-dos,
callouts, quotes, toggles), and the comment thread - so collections can be worked without
opening Notion (owner 2026-09-02: "i need all the Notion page contents, all of it").

Read-only (GET only), on demand (never in the bulk load), cached 60 s per page. Auth is the
shared Notion integration secret (`shared/notion_client.load_secret`). Relation properties
are shown as a link count (resolving their titles would be one request each); people show
their Notion display name. `--selftest` proves the renderer offline on a fixture.
"""
from __future__ import annotations

import datetime as dt
import re
import sys
import time
from pathlib import Path
from typing import Iterator, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

_ID_RE = re.compile(r"([0-9a-f]{32})(?:[?#].*)?$", re.IGNORECASE)
_CACHE: dict = {}          # page_id -> (expires, payload)
_TTL_S = 60


def page_id_from_url(url: str) -> Optional[str]:
    """'https://app.notion.com/p/33947-351b24f7...441a3a' → '351b24f7-5585-813a-9994-d8e4cc441a3a'."""
    m = _ID_RE.search((url or "").strip())
    if not m:
        return None
    h = m.group(1).lower()
    return f"{h[:8]}-{h[8:12]}-{h[12:16]}-{h[16:20]}-{h[20:]}"


def _rt(rich: list) -> str:
    return "".join((r.get("plain_text") or "") for r in (rich or []))


def _date(v: dict) -> str:
    if not v:
        return ""
    s, e = v.get("start") or "", v.get("end") or ""
    return f"{s} → {e}" if e else s


def prop_value(p: dict) -> str:
    """One property → the text a person reads in Notion. Every type the tracker uses."""
    t = p.get("type")
    v = p.get(t)
    if t in ("title", "rich_text"):
        return _rt(v)
    if t == "number":
        return "" if v is None else (f"{v:,.2f}" if isinstance(v, float) and v != int(v) else f"{v:,.0f}" if isinstance(v, (int, float)) else str(v))
    if t in ("select", "status"):
        return (v or {}).get("name") or ""
    if t == "multi_select":
        return ", ".join(o.get("name") or "" for o in (v or []))
    if t == "date":
        return _date(v or {})
    if t == "checkbox":
        return "✓" if v else "–"
    if t in ("url", "email", "phone_number"):
        return v or ""
    if t == "people":
        return ", ".join((u.get("name") or "someone") for u in (v or []))
    if t == "relation":
        n = len(v or [])
        return "" if not n else f"{n} linked"
    if t == "rollup":
        r = v or {}
        rt = r.get("type")
        if rt == "number":
            return "" if r.get("number") is None else f"{r['number']:,.2f}"
        if rt == "date":
            return _date(r.get("date") or {})
        if rt == "array":
            return ", ".join(prop_value(x) for x in (r.get("array") or []) if isinstance(x, dict))
        return ""
    if t == "formula":
        f = v or {}
        ft = f.get("type")
        if ft == "string":
            return f.get("string") or ""
        if ft == "number":
            return "" if f.get("number") is None else f"{f['number']:,.2f}"
        if ft == "boolean":
            return "✓" if f.get("boolean") else "–"
        if ft == "date":
            return _date(f.get("date") or {})
        return ""
    if t in ("created_time", "last_edited_time"):
        return (v or "")[:16].replace("T", " ")
    if t in ("created_by", "last_edited_by"):
        return (v or {}).get("name") or ""
    if t == "files":
        return ", ".join((f.get("name") or "file") for f in (v or []))
    if t == "unique_id":
        u = v or {}
        return f"{u.get('prefix') or ''}{'-' if u.get('prefix') else ''}{u.get('number') or ''}"
    return ""


_TEXT_BLOCKS = {"paragraph", "heading_1", "heading_2", "heading_3", "bulleted_list_item",
                "numbered_list_item", "to_do", "toggle", "quote", "callout", "code"}


def flatten_blocks(client, block_id: str, depth: int = 0, budget: list = None) -> Iterator[dict]:
    """Depth-first walk of a page's body → flat text lines with depth (max 3 levels, 400 blocks)."""
    budget = budget if budget is not None else [400]
    for b in client.block_children(block_id):
        if budget[0] <= 0:
            return
        budget[0] -= 1
        t = b.get("type")
        body = b.get(t) or {}
        if t in _TEXT_BLOCKS:
            yield {"type": t, "depth": depth, "text": _rt(body.get("rich_text")),
                   "checked": body.get("checked") if t == "to_do" else None,
                   "at": (b.get("created_time") or "")[:16].replace("T", " ")}
        elif t == "divider":
            yield {"type": "divider", "depth": depth, "text": "", "checked": None, "at": ""}
        elif t in ("child_page", "child_database"):
            yield {"type": t, "depth": depth, "text": body.get("title") or "", "checked": None, "at": ""}
        elif t in ("image", "file", "pdf", "bookmark", "embed"):
            src = (body.get("external") or {}).get("url") or (body.get("file") or {}).get("url") or body.get("url") or ""
            yield {"type": t, "depth": depth, "text": _rt(body.get("caption")) or src, "checked": None, "at": "", "url": src}
        elif t == "table":
            for row in client.block_children(b["id"]):
                cells = (row.get("table_row") or {}).get("cells") or []
                yield {"type": "table_row", "depth": depth, "text": " | ".join(_rt(c) for c in cells), "checked": None, "at": ""}
            continue
        if b.get("has_children") and depth < 3 and t != "table":
            yield from flatten_blocks(client, b["id"], depth + 1, budget)


def comments(client, page_id: str) -> list:
    """The page's comment thread (newest last). Empty when the integration lacks comment access."""
    out = []
    try:
        start = None
        while True:
            q = f"?block_id={page_id}&page_size=100" + (f"&start_cursor={start}" if start else "")
            data = client._request("GET", f"/comments{q}")
            for c in data.get("results", []):
                out.append({"text": _rt(c.get("rich_text")),
                            "by": ((c.get("created_by") or {}).get("name")) or "",
                            "at": (c.get("created_time") or "")[:16].replace("T", " ")})
            if not data.get("has_more"):
                break
            start = data.get("next_cursor")
    except Exception:  # noqa: BLE001 - comments are a bonus; never fail the page for them
        return out
    return out


def render(page: dict, blocks: list, thread: list) -> dict:
    props = page.get("properties") or {}
    title = ""
    plist = []
    for name, p in props.items():
        val = prop_value(p)
        if p.get("type") == "title":
            title = val
        plist.append({"name": name, "type": p.get("type"), "value": val})
    # title first, then the properties people read first (status / notes / dates), then the rest A-Z
    lead = ("Quick Status", "Status", "Notes", "Collections", "Aging Bucket", "Due Date", "Paid Date")
    plist.sort(key=lambda d: (0 if d["type"] == "title" else 1 if d["name"] in lead else 2,
                              lead.index(d["name"]) if d["name"] in lead else 0, d["name"].lower()))
    return {"ok": True, "title": title, "url": page.get("url"), "properties": plist, "blocks": blocks,
            "comments": thread, "last_edited": (page.get("last_edited_time") or "")[:16].replace("T", " "),
            "fetched_at": dt.datetime.now().isoformat(timespec="seconds")}


def fetch(url: str) -> dict:
    pid = page_id_from_url(url)
    if not pid:
        return {"ok": False, "error": "not a Notion page url"}
    hit = _CACHE.get(pid)
    if hit and hit[0] > time.time():
        return hit[1]
    from shared.notion_client import NotionClient, NotionError
    try:
        nc = NotionClient()
        page = nc.retrieve_page(pid)
        blocks = list(flatten_blocks(nc, pid))
        thread = comments(nc, pid)
    except NotionError as e:
        return {"ok": False, "error": str(e)}
    except Exception as e:  # noqa: BLE001 - surface the reason in the panel, never a 500
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}
    out = render(page, blocks, thread)
    _CACHE[pid] = (time.time() + _TTL_S, out)
    return out


def _selftest() -> None:
    assert page_id_from_url("https://app.notion.com/p/33947-351b24f75585813a9994d8e4cc441a3a") == "351b24f7-5585-813a-9994-d8e4cc441a3a"
    assert page_id_from_url("https://www.notion.so/Invoice-34415-351b24f75585813a9994d8e4cc441a3a?pvs=4") == "351b24f7-5585-813a-9994-d8e4cc441a3a"
    assert page_id_from_url("nope") is None
    page = {"url": "u", "last_edited_time": "2026-09-02T14:00:00.000Z", "properties": {
        "Invoice #": {"type": "title", "title": [{"plain_text": "34415"}]},
        "Quick Status": {"type": "rich_text", "rich_text": [{"plain_text": "Unconditionals needed"}]},
        "Total Amount": {"type": "number", "number": 103312.0},
        "Status": {"type": "select", "select": {"name": "Unpaid"}},
        "Due Date": {"type": "date", "date": {"start": "2026-07-01", "end": None}},
        "Litigation": {"type": "checkbox", "checkbox": False},
        "Lien": {"type": "relation", "relation": [{"id": "a"}, {"id": "b"}]},
        "Owner": {"type": "people", "people": [{"name": "Rep A"}]},
        "Days": {"type": "formula", "formula": {"type": "number", "number": 63}},
        "Paid": {"type": "rollup", "rollup": {"type": "number", "number": 0}},
    }}
    blocks = [{"type": "paragraph", "depth": 0, "text": "Called the PM.", "checked": None, "at": "2026-08-12 10:00"},
              {"type": "to_do", "depth": 1, "text": "send waiver", "checked": True, "at": ""}]
    out = render(page, blocks, [{"text": "paying Friday", "by": "Rep A", "at": "2026-08-20 09:00"}])
    assert out["title"] == "34415" and out["properties"][0]["name"] == "Invoice #"
    got = {p["name"]: p["value"] for p in out["properties"]}
    assert got["Quick Status"] == "Unconditionals needed" and got["Total Amount"] == "103,312"
    assert got["Status"] == "Unpaid" and got["Due Date"] == "2026-07-01" and got["Litigation"] == "–"
    assert got["Lien"] == "2 linked" and got["Owner"] == "Rep A" and got["Days"] == "63.00" and got["Paid"] == "0.00"
    assert out["properties"][1]["name"] == "Quick Status"            # the lead group comes right after the title
    assert out["comments"][0]["by"] == "Rep A" and len(out["blocks"]) == 2
    print("notion_page selftest OK - id from url, every property type, lead ordering, body + comments")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        _selftest()
    elif len(sys.argv) > 1:
        import json
        print(json.dumps(fetch(sys.argv[1]), indent=1)[:3000])
    else:
        print("usage: notion_page.py --selftest | <notion page url>")
