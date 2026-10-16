#!/usr/bin/env python3
"""
reclassify.py - Reclassify transactions: find QuickBooks LINES and bulk-change their account, item, class,
project, line memo and document memo.

QuickBooks' own reclassify, at line level (the owner 10/06 + 10/09/2026: "you look at it from a P&L, you click
the account and select the transactions to move ... i need that but for line items so that we can move that
safely ... reclassify all, item or category, memo, class, project ... bulk edit all of these things").
A sub's bill that spans several jobs keeps its other lines exactly as they are - only the ticked lines change,
and only the fields asked for. Amounts, dates, vendors and payments are never touched.

The ledger's Company -> Reclassify transactions page runs it: options() (the pickers), plan(filters) (the lines,
or a P&L by account to click into; mirror read, no write), commit(keys, sigs, changes, n, amount) (the write).

Lines it reads and what each can change:
    Bill / Expense / Vendor credit, category line   account · class · project · memo
    Bill / Expense / Vendor credit, item line       item (cost code) · class · project · memo
    Journal entry line                               account · class · project (customer lines) · memo
    Invoice / Credit memo / Sales receipt line       item · class · memo (the customer is the whole document's)
    any of them                                      the document memo
A category line stays a category line and an item line stays an item line (as in QuickBooks): an account goes on
category lines, an item on item lines.

changes = {"account": {old id: new id}, "item": {...}, "class": {...}, "project": {old customer id: new id},
           "memo": {"mode": "set" | "replace", "find": "...", "text": "..."}, "docmemo": {...same}}
An old id of "" is a line with nothing in that field (no class, no project).

Guards (live from QuickBooks at write time, one document at a time):
  - the server refuses unless every ticked line is still exactly what the preview showed (each line's signature,
    the same count, the same amount)
  - Touch ID on this Mac (ledger/presence.py), naming the line count and amount
  - each document is re-read right before the write (fresh SyncToken); a line that changed in QuickBooks since
    the preview stops the run before that document is written
  - nothing dated on or before the books-closed date is touched (closing password - by hand)
  - the live document is saved to ~/Library/Logs/Proficient/move-project-lines/ before its write
  - after the write the total, open balance and every line amount must match what was there before, and every
    changed field must read back as asked; if not, the run stops and names the document
QuickBooks Projects keeps a second project tag on each line (`Line.ProjectRef`, a Projects id, not the customer
id) and resets the line's customer to match it - a CustomerRef-only edit is silently undone (found 10/06/2026).
A project move writes both: the tag read off the lines already on the new project; a project with no lines yet
has no known tag, so the tag is dropped and the read-back proves QuickBooks took it.

Command line (the old project-to-project move):
    ledger/reclassify.py RP7242-FTW RP7242                          preview
    ledger/reclassify.py RP7242-FTW RP7242 --recode FW2=SL2,FW3=SL3  also swap these items on the moved lines
    ledger/reclassify.py RP7242-FTW RP7242 --commit                 write it
"""

import argparse
import datetime as dt
import hashlib
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import requests  # noqa: E402

from shared import qbo_api  # noqa: E402
from shared import qbo_mirror  # noqa: E402

EXPENSE = ("Bill", "Purchase", "VendorCredit")
SALES = ("Invoice", "CreditMemo", "SalesReceipt")
ENTITIES = EXPENSE + ("JournalEntry",) + SALES
KINDS = {"AccountBasedExpenseLineDetail": "category", "ItemBasedExpenseLineDetail": "item",
         "JournalEntryLineDetail": "journal", "SalesItemLineDetail": "sales"}
EDIT = {"category": {"account", "class", "project", "memo", "docmemo"},
        "item": {"item", "class", "project", "memo", "docmemo"},
        "journal": {"account", "class", "project", "memo", "docmemo"},
        "sales": {"item", "class", "memo", "docmemo"}}
REF_FIELDS = ("account", "item", "class", "project")
TEXT_FIELDS = ("memo", "docmemo")
TEXT_MAX = 4000
MAX_ROWS = 5000
BACKUP_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "move-project-lines"
EPS = 0.005


class PlanError(RuntimeError):
    """A reason the plan cannot be made - safe to show on the page."""


def _proj(name: str) -> str:
    """The project number a customer name carries (`Parent:RP7242-FTW 123 Main` -> RP7242-FTW)."""
    leaf = (name or "").split(":")[-1].strip().upper()
    return leaf.split()[0] if leaf else ""


def _fqn(c: dict) -> str:
    return c.get("FullyQualifiedName") or c.get("DisplayName") or c.get("Name") or ""


def _ref(r) -> tuple:
    r = r or {}
    return str(r.get("value") or ""), r.get("name") or ""


def _auth():
    return qbo_api.load_credentials()


# ── the lists ────────────────────────────────────────────────────────────────────────────────────

def _lists() -> dict:
    """Every pick list, from the mirror, keyed by id."""
    acc = {a["Id"]: a for a in qbo_mirror.load("Account")}
    items = {i["Id"]: i for i in qbo_mirror.load("Item")}
    classes = {c["Id"]: c for c in qbo_mirror.load("Class")}
    cust = {c["Id"]: c for c in qbo_mirror.load("Customer")}
    return {"account": acc, "item": items, "class": classes, "project": cust}


def options() -> dict:
    """The page's pickers: active accounts, items, classes and customers (projects carry their project #)."""
    L = _lists()
    act = lambda d: [x for x in d.values() if x.get("Active", True)]   # noqa: E731
    accounts = sorted(({"id": a["Id"], "name": _fqn(a), "type": a.get("AccountType") or "",
                        "pl": a.get("Classification") in ("Revenue", "Expense")} for a in act(L["account"])),
                      key=lambda x: x["name"].lower())
    items = sorted(({"id": i["Id"], "name": i.get("Name") or "", "type": i.get("Type") or ""}
                    for i in act(L["item"]) if i.get("Type") in ("Service", "NonInventory", "Inventory")),
                   key=lambda x: x["name"].lower())
    classes = sorted(({"id": c["Id"], "name": _fqn(c)} for c in act(L["class"])), key=lambda x: x["name"].lower())
    by_proj = {}
    for c in act(L["project"]):
        p = _proj(_fqn(c))
        if qbo_api.PROJ_RE.match(p or ""):
            by_proj.setdefault(p, []).append(c["Id"])
    projects = sorted(({"id": c["Id"], "name": _fqn(c),
                        "proj": p if len(by_proj.get(p := _proj(_fqn(c)), [])) == 1 else ""}
                       for c in act(L["project"])), key=lambda x: x["name"].lower())
    return {"ok": True, "accounts": accounts, "items": items, "classes": classes, "projects": projects}


# ── reading a line ───────────────────────────────────────────────────────────────────────────────

def _kind(ln: dict):
    for k, kind in KINDS.items():
        if k in ln:
            return kind, ln[k]
    return None, None


def _sign(ent: str, t: dict, ln: dict, kind: str) -> int:
    if kind == "journal":
        return -1 if (ln.get("JournalEntryLineDetail") or {}).get("PostingType") == "Credit" else 1
    if ent in ("VendorCredit", "CreditMemo") or (ent == "Purchase" and t.get("Credit")):
        return -1
    return 1


def read_line(ent: str, t: dict, ln: dict, L: dict):
    """One line as the page sees it, or None for a subtotal / description / group line."""
    kind, det = _kind(ln)
    if not kind:
        return None
    acct = _ref(det.get("AccountRef"))
    item = _ref(det.get("ItemRef"))
    if kind == "item":
        i = L["item"].get(item[0]) or {}
        acct = _ref(i.get("ExpenseAccountRef"))
    elif kind == "sales":
        i = L["item"].get(item[0]) or {}
        acct = _ref(det.get("ItemAccountRef") or i.get("IncomeAccountRef"))
    if acct[0] and not acct[1]:
        acct = (acct[0], _fqn(L["account"].get(acct[0]) or {}))
    elif acct[0]:
        acct = (acct[0], _fqn(L["account"].get(acct[0]) or {}) or acct[1])
    cls = _ref(det.get("ClassRef"))
    proj_ok = kind in ("category", "item")
    if kind == "journal":
        ent_ref = det.get("Entity") or {}
        proj_ok = not ent_ref or ent_ref.get("Type") == "Customer"
        proj = _ref(ent_ref.get("EntityRef")) if ent_ref.get("Type") == "Customer" else ("", "")
    elif kind == "sales":
        proj = _ref(t.get("CustomerRef"))
    else:
        proj = _ref(det.get("CustomerRef"))
    if proj[0]:
        proj = (proj[0], _fqn(L["project"].get(proj[0]) or {}) or proj[1])
    if cls[0]:
        cls = (cls[0], _fqn(L["class"].get(cls[0]) or {}) or cls[1])
    if item[0]:
        item = (item[0], (L["item"].get(item[0]) or {}).get("Name") or item[1])
    party = t.get("VendorRef") or t.get("EntityRef") or t.get("CustomerRef") or {}
    edit = set(EDIT[kind]) - (set() if proj_ok else {"project"})
    amt = round(float(ln.get("Amount") or 0), 2)
    row = {"key": f"{ent}:{t['Id']}:{ln.get('Id') or ''}", "entity": ent, "id": t["Id"], "line": str(ln.get("Id") or ""),
           "doc": t.get("DocNumber") or "", "date": t.get("TxnDate") or "", "name": party.get("name") or "",
           "kind": kind, "edit": sorted(edit),
           "account": acct[0], "account_name": acct[1], "item": item[0], "item_name": item[1],
           "class": cls[0], "class_name": cls[1], "project": proj[0], "project_name": proj[1],
           "proj": _proj(proj[1]) if proj[1] else "",
           "memo": (ln.get("Description") or ""), "docmemo": (t.get("PrivateNote") or ""),
           "amount": amt, "signed": round(_sign(ent, t, ln, kind) * amt, 2),
           "doc_lines": sum(1 for x in t.get("Line") or [] if _kind(x)[0])}
    row["sig"] = _sig(row)
    return row


def _sig(r: dict) -> str:
    s = json.dumps([r["account"], r["item"], r["class"], r["project"], r["memo"], r["docmemo"], r["amount"]])
    return hashlib.sha1(s.encode()).hexdigest()[:16]


def _descendants(L: dict, acct_id: str) -> set:
    """An account and every sub-account under it (a P&L parent row holds its children)."""
    root = _fqn(L["account"].get(acct_id) or {})
    if not root:
        return {acct_id}
    return {i for i, a in L["account"].items() if i == acct_id or _fqn(a).startswith(root + ":")}


# ── finding lines ────────────────────────────────────────────────────────────────────────────────

def _closed_date(access: str, cid: str) -> str:
    prefs = qbo_api._api_get(f"/v3/company/{cid}/preferences", access).get("Preferences", {})
    return (prefs.get("AccountingInfoPrefs") or {}).get("BookCloseDate") or ""


def _load_docs(types, d1: str, d2: str) -> dict:
    where, params = [], []
    if d1:
        where.append("txn_date >= ?")
        params.append(d1)
    if d2:
        where.append("txn_date <= ?")
        params.append(d2)
    return {e: qbo_mirror.load(e, " AND ".join(where), tuple(params)) for e in types}


def plan(f: dict) -> dict:
    """The lines matching the filters (mode=lines), or a P&L by account for the date range (mode=pnl).
    Writes nothing."""
    f = {k: str(v or "").strip() for k, v in (f or {}).items()}
    types = [t for t in (f.get("types") or "").split(",") if t in ENTITIES] or list(ENTITIES)
    d1, d2 = f.get("d1", ""), f.get("d2", "")
    narrowing = [k for k in ("account", "item", "class", "project", "name", "memo", "doc") if f.get(k)]
    if f.get("mode") != "pnl" and not narrowing:
        return {"ok": False, "error": "Pick an account, item, class, project, name, memo or doc # - or open the P&L."}
    L = _lists()
    accts = _descendants(L, f["account"]) if f.get("account") else None
    name, memo, doc = f.get("name", "").lower(), f.get("memo", "").lower(), f.get("doc", "").lower()
    docs = _load_docs(types, d1, d2)
    access, cid = _auth()
    closed = _closed_date(access, cid)
    rows = []
    for ent, recs in docs.items():
        for t in recs:
            if doc and doc not in (t.get("DocNumber") or "").lower():
                continue
            for ln in t.get("Line") or []:
                r = read_line(ent, t, ln, L)
                if not r or (accts is not None and r["account"] not in accts):
                    continue
                if f.get("item") and r["item"] != f["item"]:
                    continue
                if f.get("class") and r["class"] != ("" if f["class"] == "none" else f["class"]):
                    continue
                if f.get("project") and r["project"] != ("" if f["project"] == "none" else f["project"]):
                    continue
                if name and name not in r["name"].lower():
                    continue
                if memo and memo not in r["memo"].lower() and memo not in r["docmemo"].lower():
                    continue
                r["closed"] = bool(closed and r["date"] and r["date"] <= closed)
                rows.append(r)
    if f.get("mode") == "pnl":
        return _pnl(rows, L, d1, d2)
    rows.sort(key=lambda r: (r["date"], r["entity"], r["id"], r["line"]))
    total = len(rows)
    return {"ok": True, "closed_date": closed, "rows": rows[:MAX_ROWS], "total": total,
            "truncated": total > MAX_ROWS, "amount": round(sum(r["signed"] for r in rows), 2)}


def _pnl(rows: list, L: dict, d1: str, d2: str) -> dict:
    """Account totals the way the P&L shows them (income credit-positive, costs debit-positive)."""
    out = {}
    for r in rows:
        a = L["account"].get(r["account"]) or {}
        cls = a.get("Classification")
        if cls not in ("Revenue", "Expense"):
            continue
        v = r["signed"]
        if cls == "Revenue" and r["kind"] == "journal":
            v = -v
        x = out.setdefault(r["account"], {"id": r["account"], "name": r["account_name"], "cls": cls,
                                          "type": a.get("AccountType") or "", "lines": 0, "amount": 0.0})
        x["lines"] += 1
        x["amount"] = round(x["amount"] + v, 2)
    accts = sorted(out.values(), key=lambda x: (x["cls"] != "Revenue", x["name"].lower()))
    return {"ok": True, "mode": "pnl", "d1": d1, "d2": d2, "accounts": accts}


# ── the change ───────────────────────────────────────────────────────────────────────────────────

def _text_new(old: str, spec) -> str:
    if not isinstance(spec, dict) or spec.get("mode") not in ("set", "replace"):
        return old
    text = str(spec.get("text") or "")
    if spec["mode"] == "set":
        return text
    find = str(spec.get("find") or "")
    return old.replace(find, text) if find else old


def clean_changes(changes) -> dict:
    """Only the shapes the writer understands; ids as strings."""
    out = {}
    if not isinstance(changes, dict):
        return out
    for f in REF_FIELDS:
        m = changes.get(f)
        if isinstance(m, dict):
            m = {str(o): str(n) for o, n in m.items() if str(n or "").strip() and str(n) != str(o)}
            if m:
                out[f] = m
    for f in TEXT_FIELDS:
        s = changes.get(f)
        if isinstance(s, dict) and s.get("mode") in ("set", "replace"):
            if s["mode"] == "replace" and not str(s.get("find") or ""):
                continue
            out[f] = {"mode": s["mode"], "find": str(s.get("find") or ""), "text": str(s.get("text") or "")}
    return out


def row_changes(r: dict, changes: dict) -> dict:
    """{field: (old, new)} this line would get - only fields it can carry, only real changes."""
    out = {}
    for f in REF_FIELDS:
        if f in r["edit"] and f in changes and r[f] in changes[f]:
            out[f] = (r[f], changes[f][r[f]])
    for f in TEXT_FIELDS:
        if f in r["edit"] and f in changes:
            new = _text_new(r[f], changes[f])
            if new != r[f]:
                out[f] = (r[f], new)
    return out


def validate(changes: dict, L: dict) -> None:
    """Every new value must be one active QuickBooks record of the right kind."""
    for f in REF_FIELDS:
        for new in (changes.get(f) or {}).values():
            rec = L[f].get(new)
            if not rec or not rec.get("Active", True):
                raise PlanError(f"{f} id {new}: not an active QuickBooks {f}")
    for f in TEXT_FIELDS:
        if f in changes and len(changes[f]["text"]) > TEXT_MAX:
            raise PlanError(f"{f}: longer than QuickBooks takes ({TEXT_MAX} characters)")


def project_tags(L_docs, ids) -> dict:
    """{customer id -> the Line.ProjectRef its lines carry}; one tag per project or the move stops."""
    seen = {i: set() for i in ids}
    for recs in L_docs.values():
        for t in recs:
            for ln in t.get("Line") or []:
                kind, d = _kind(ln)
                if not kind:
                    continue
                v = ((d.get("Entity") or {}).get("EntityRef") if kind == "journal" else d.get("CustomerRef")) or {}
                v = v.get("value")
                if v in seen and (ln.get("ProjectRef") or {}).get("value"):
                    seen[v].add(ln["ProjectRef"]["value"])
    out = {}
    for i, tags in seen.items():
        if len(tags) > 1:
            raise PlanError(f"customer {i}: its lines carry {len(tags)} different project tags - fix by hand")
        out[i] = next(iter(tags), None)
    return out


def _apply(t: dict, ln: dict, ch: dict, tags: dict) -> None:
    kind, det = _kind(ln)
    if "account" in ch:
        det["AccountRef"] = {"value": ch["account"][1]}
    if "item" in ch:
        det["ItemRef"] = {"value": ch["item"][1]}
        det.pop("ItemAccountRef", None)
    if "class" in ch:
        det["ClassRef"] = {"value": ch["class"][1]}
    if "project" in ch:
        new = ch["project"][1]
        if kind == "journal":
            det["Entity"] = {"Type": "Customer", "EntityRef": {"value": new}}
        else:
            det["CustomerRef"] = {"value": new}
        if tags.get(new):
            ln["ProjectRef"] = {"value": tags[new]}
        else:
            ln.pop("ProjectRef", None)
    if "memo" in ch:
        ln["Description"] = ch["memo"][1]
    if "docmemo" in ch:
        t["PrivateNote"] = ch["docmemo"][1]


def _get(access: str, cid: str, ent: str, tid: str) -> dict:
    return qbo_api._api_get(f"/v3/company/{cid}/{ent.lower()}/{tid}", access).get(ent, {})


def _post(access: str, cid: str, ent: str, body: dict) -> dict:
    r = requests.post(f"{qbo_api.API_BASE}/v3/company/{cid}/{ent.lower()}",
                      headers={"Authorization": f"Bearer {qbo_api._CURRENT_ACCESS or access}",
                               "Content-Type": "application/json", "Accept": "application/json"},
                      params={"minorversion": qbo_api.MINOR_VERSION}, json=body, timeout=120)
    if r.status_code != 200:
        try:
            e = r.json()["Fault"]["Error"][0]
            msg = f"{e.get('Message')}: {e.get('Detail')}"
        except Exception:                                   # noqa: BLE001
            msg = r.text[:400]
        raise RuntimeError(f"QuickBooks refused the update ({r.status_code}): {msg} - nothing was changed on it")
    return r.json()[ent]


def _fingerprint(t: dict) -> tuple:
    return (round(float(t.get("TotalAmt") or 0), 2), round(float(t.get("Balance") or 0), 2),
            tuple(sorted((str(ln.get("Id")), round(float(ln.get("Amount") or 0), 2)) for ln in t.get("Line") or [])))


def _describe(changes: dict, L: dict) -> str:
    bits = []
    for f in REF_FIELDS:
        news = {n for n in (changes.get(f) or {}).values()}
        if news:
            names = sorted(_fqn(L[f].get(n) or {}) or n for n in news)
            bits.append(f"{f} -> {', '.join(names[:3])}{' …' if len(names) > 3 else ''}")
    bits += [f"{'line memo' if f == 'memo' else 'document memo'} edited" for f in TEXT_FIELDS if f in changes]
    return "; ".join(bits)


def commit(keys: list, sigs: dict, changes, n: int, amount: float) -> dict:
    """Write the lines the owner ticked - only if each is still exactly what the preview showed (same signature),
    the change touches every one of them, and the count and amount match."""
    changes = clean_changes(changes)
    if not changes:
        return {"ok": False, "error": "Nothing to change - pick a new value first."}
    L = _lists()
    try:
        validate(changes, L)
    except PlanError as e:
        return {"ok": False, "error": str(e)}
    keys = list(dict.fromkeys(str(k) for k in keys or []))
    want = {}
    for k in keys:
        ent, tid, lid = (k.split(":") + ["", "", ""])[:3]
        if ent not in ENTITIES or not tid:
            return {"ok": False, "error": f"bad line {k}"}
        want.setdefault((ent, tid), set()).add(lid)
    rows, mirror_docs = {}, {}
    for (ent, tid), lids in want.items():
        t = qbo_mirror.get(ent, tid)
        if not t:
            return {"ok": False, "error": f"{ent} {tid} is no longer in QuickBooks - find the lines again."}
        mirror_docs[(ent, tid)] = t
        for ln in t.get("Line") or []:
            if str(ln.get("Id") or "") in lids:
                r = read_line(ent, t, ln, L)
                if r:
                    rows[r["key"]] = r
    access, cid = _auth()
    closed = _closed_date(access, cid)
    plan_rows = []
    for k in keys:
        r = rows.get(k)
        if not r or r["sig"] != str((sigs or {}).get(k) or "") or (closed and r["date"] <= closed):
            return {"ok": False, "error": "The lines changed in QuickBooks since the preview - find them again."}
        ch = row_changes(r, changes)
        if not ch:
            return {"ok": False, "error": f"{r['entity']} {r['doc'] or r['id']}: the change does not apply to that line - untick it."}
        plan_rows.append((r, ch))
    amt = round(sum(r["signed"] for r, _ in plan_rows), 2)
    if len(plan_rows) != int(n) or abs(amt - float(amount)) > EPS:
        return {"ok": False, "error": "The lines changed in QuickBooks since the preview - find them again."}
    tags = {}
    if "project" in changes:
        try:
            tags = project_tags(_load_docs(ENTITIES, "", ""), set(changes["project"].values()))
        except PlanError as e:
            return {"ok": False, "error": str(e)}
    what = _describe(changes, L)
    import presence                                       # noqa: PLC0415  (ledger-local: macOS' own Touch ID dialog)
    ok, why = presence.confirm(f"reclassify {len(plan_rows)} lines ({amt:,.2f}) in QuickBooks: {what}")
    if not ok:
        return {"ok": False, "error": f"not confirmed on this Mac ({why}) - nothing was written"}
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    by_doc = {}
    for r, ch in plan_rows:
        by_doc.setdefault((r["entity"], r["id"]), {})[r["line"]] = (r, ch)
    out, stopped = [], ""
    for (ent, tid), lines in by_doc.items():
        live = _get(access, cid, ent, tid)
        doc = live.get("DocNumber") or tid
        if closed and live.get("TxnDate", "") <= closed:
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "skipped", "detail": "closed period"})
            continue
        changed_since = False
        for ln in live.get("Line") or []:
            lid = str(ln.get("Id") or "")
            if lid in lines:
                now = read_line(ent, live, ln, L)
                if not now or now["sig"] != lines[lid][0]["sig"]:
                    changed_since = True
        if changed_since or not set(lines) <= {str(ln.get("Id") or "") for ln in live.get("Line") or []}:
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "changed",
                        "detail": "edited in QuickBooks since the preview - not written"})
            stopped = f"{ent} {doc}: edited in QuickBooks since the preview - find the lines again"
            break
        before = _fingerprint(live)
        (BACKUP_DIR / f"{stamp}_{ent}_{tid}.json").write_text(json.dumps(live, indent=1))
        for ln in live.get("Line") or []:
            lid = str(ln.get("Id") or "")
            if lid in lines:
                _apply(live, ln, lines[lid][1], tags)
        body = {k: v for k, v in live.items() if k not in ("MetaData", "domain", "sparse")}
        try:
            after = _post(access, cid, ent, body)
        except RuntimeError as e:
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "refused", "detail": str(e)})
            stopped = f"{ent} {doc}: {e}"
            break
        missed = []
        for ln in after.get("Line") or []:
            lid = str(ln.get("Id") or "")
            if lid in lines:
                got = read_line(ent, after, ln, L) or {}
                missed += [f for f, (_, new) in lines[lid][1].items() if got.get(f) != new]
        if missed:
            fields = ", ".join(sorted(set(missed)))
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "not changed",
                        "detail": f"QuickBooks did not keep: {fields}"})
            stopped = f"{ent} {doc}: QuickBooks did not keep the new {fields} - check it"
            break
        if _fingerprint(after) != before:
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "money changed",
                        "detail": "total / balance / a line amount moved after the write - check it"})
            stopped = f"{ent} {doc}: an amount moved after the write - check it in QuickBooks"
            break
        out.append({"entity": ent, "id": tid, "doc": doc, "result": "done",
                    "detail": f"{len(lines)} line(s)"})
    with open(BACKUP_DIR / "runs.jsonl", "a") as f:
        f.write(json.dumps({"at": stamp, "tool": "reclassify", "changes": changes, "what": what,
                            "lines": len(plan_rows), "amount": amt, "results": out, "stopped": stopped}) + "\n")
    done = [r for r in out if r["result"] == "done"]
    return {"ok": not stopped, "error": stopped or None, "results": out, "moved_docs": len(done),
            "moved_lines": sum(int(r["detail"].split()[0]) for r in done), "backup_dir": str(BACKUP_DIR)}


# ── command line: the old project -> project move ───────────────────────────────────────────────

def _customer_id(proj: str) -> str:
    hits = [c for c in qbo_mirror.load("Customer") if c.get("Active", True) and _proj(_fqn(c)) == proj.upper()]
    if len(hits) != 1:
        raise PlanError(f"{proj}: expected exactly one active QuickBooks customer, found {len(hits)}")
    return hits[0]["Id"]


def main() -> None:
    ap = argparse.ArgumentParser(description="Reclassify: move lines from one project to another (line level).")
    ap.add_argument("src", help="project number to move OFF, e.g. RP7242-FTW")
    ap.add_argument("dst", help="project number to move ONTO, e.g. RP7242")
    ap.add_argument("--commit", action="store_true", help="write to QuickBooks (default: preview)")
    ap.add_argument("--recode", default="", help="also swap items on the moved lines, e.g. FW2=SL2,FW3=SL3")
    a = ap.parse_args()
    try:
        src, dst = _customer_id(a.src), _customer_id(a.dst)
        items = {(i.get("Name") or "").upper(): i["Id"] for i in qbo_mirror.load("Item") if i.get("Active", True)}
        recode = {}
        for pair in filter(None, a.recode.split(",")):
            o, nw = (x.strip().upper() for x in pair.split("=", 1))
            if o not in items or nw not in items:
                raise PlanError(f"cost code {o if o not in items else nw}: no active QuickBooks item")
            recode[items[o]] = items[nw]
        p = plan({"project": src, "types": ",".join(EXPENSE + ("JournalEntry",))})
    except (PlanError, qbo_api.AuthError) as e:
        sys.exit(f"✗  {e}")
    if not p["ok"]:
        sys.exit(f"✗  {p['error']}")
    changes = clean_changes({"project": {src: dst}, "item": recode})
    pick = [r for r in p["rows"] if not r["closed"] and row_changes(r, changes)]
    for r in p["rows"]:
        d = dt.date.fromisoformat(r["date"]).strftime("%m/%d/%Y") if r["date"] else ""
        mark = "  CLOSED" if r["closed"] else ("" if r in pick else "  not movable here")
        print(f"{r['entity']:<12} {r['doc'][:14]:<14} {d:<10} {r['name'][:28]:<28} "
              f"{(r['item_name'] or r['account_name'])[:28]:<28} {r['signed']:>12,.2f}{mark}")
    amt = round(sum(r["signed"] for r in pick), 2)
    print(f"\n{len(pick)} line(s), net {amt:,.2f}: {a.src.upper()} -> {a.dst.upper()}")
    if not a.commit:
        print("\nPreview - nothing written. Add --commit to move these lines.")
        return
    res = commit([r["key"] for r in pick], {r["key"]: r["sig"] for r in pick}, changes, len(pick), amt)
    for r in res.get("results") or []:
        print(f"   {r['result']:<14} {r['entity']} {r['doc']}  {r['detail']}")
    if not res["ok"]:
        sys.exit(f"✗  {res['error']}")
    print(f"\nDone: {res['moved_lines']} line(s) on {res['moved_docs']} document(s). Before copies: {res['backup_dir']}")


if __name__ == "__main__":
    main()
