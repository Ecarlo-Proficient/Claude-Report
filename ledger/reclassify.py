#!/usr/bin/env python3
"""
reclassify.py - Reclassify transactions: move cost LINES from one QuickBooks project to another.

Line level only (the owner 10/06/2026): a sub's bill that spans several jobs keeps its other lines
exactly as they are - only the chosen lines on FROM change, and only their project (plus, when asked, their
cost code). Amount, account, class, description and the bill's payments are left alone.

The ledger's Company -> Reclassify transactions page runs it (preview -> tick lines -> "Are you sure?" ->
write): options() / plan() / commit(). The command line does the same:

    ledger/reclassify.py RP7242-FTW RP7242                          preview: every line that would move
    ledger/reclassify.py RP7242-FTW RP7242 --recode FW2=SL2,FW3=SL3  also swap these cost codes on the moved lines
    ledger/reclassify.py RP7242-FTW RP7242 --commit                 write it to QuickBooks

Covers bills, expenses / checks (Purchase) and vendor credits. Journal entry lines are listed, never
written - those are fixed by hand.

Guards (live from QuickBooks at write time, one document at a time):
  - the server refuses unless the plan is still the one the owner confirmed (same lines, same amount)
  - Touch ID on this Mac (ledger/presence.py), naming the line count, amount and both projects
  - the document is re-read right before the write (fresh SyncToken); a line no longer on FROM is left
  - nothing dated on or before the books-closed date is touched (listed instead - that needs the
    closing password, so the owner moves those by hand)
  - the live document is saved to ~/Library/Logs/Proficient/move-project-lines/ before its write
  - after the write the total, open balance and every line amount must match what was there before,
    and every moved line must come back on TO; if not, the run stops and names the document
QuickBooks Projects keeps a second project tag on each line (`Line.ProjectRef`, a Projects id, not the
customer id) and resets the line's customer to match it - a CustomerRef-only edit is silently undone
(found 10/06/2026). The tool reads each project's tag off the lines already on it and moves both; a TO with
no lines yet has no known tag, so the tag is dropped and the landing check proves QuickBooks took it.
Projects are matched by the EXACT project number (RP7242 and RP7242-FTW are different projects).
"""

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import requests  # noqa: E402

from shared import qbo_api  # noqa: E402
from shared.qbo_costs import is_cost_code  # noqa: E402

ENTITIES = ("Bill", "Purchase", "VendorCredit")
DETAILS = ("AccountBasedExpenseLineDetail", "ItemBasedExpenseLineDetail")
BACKUP_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "move-project-lines"
EPS = 0.005


class PlanError(RuntimeError):
    """A reason the move cannot be planned - safe to show on the page."""


def _proj(name: str) -> str:
    """The project number a customer name carries (`Parent:RP7242-FTW 123 Main` -> RP7242-FTW)."""
    leaf = (name or "").split(":")[-1].strip().upper()
    return leaf.split()[0] if leaf else ""


def _fqn(c: dict) -> str:
    return c.get("FullyQualifiedName") or c.get("DisplayName") or ""


def find_customer(customers: list, proj: str) -> dict:
    hits = [c for c in customers if c.get("Active", True) and _proj(_fqn(c)) == proj]
    if len(hits) != 1:
        names = ", ".join(_fqn(c) for c in hits) or "none"
        raise PlanError(f"{proj}: expected exactly one active QuickBooks customer, found {len(hits)} ({names})")
    return hits[0]


def _code(det: dict) -> str:
    return (det.get("ItemRef") or {}).get("name") or (det.get("AccountRef") or {}).get("name") or ""


def _det(ln: dict):
    dk = next((k for k in DETAILS if k in ln), None)
    return ln[dk] if dk else None


def _key(ent: str, tid: str, lid: str) -> str:
    return f"{ent}:{tid}:{lid}"


def _auth():
    access, cid = qbo_api.load_credentials()
    return access, cid


def options() -> dict:
    """The page's pickers: every active project # (one customer each) and every cost-code item."""
    access, cid = _auth()
    seen = {}
    for c in qbo_api.query_all(access, cid, "Customer"):
        p = _proj(_fqn(c))
        if c.get("Active", True) and qbo_api.PROJ_RE.match(p or ""):
            seen.setdefault(p, []).append(_fqn(c))
    projects = [{"proj": p, "name": n[0]} for p, n in sorted(seen.items()) if len(n) == 1]
    codes = sorted({(i.get("Name") or "").upper() for i in qbo_api.query_all(access, cid, "Item")
                    if i.get("Active", True) and is_cost_code(i.get("Name") or "")})
    return {"ok": True, "projects": projects, "codes": codes}


def project_tags(docs: dict, ids: tuple) -> dict:
    """{customer id -> the Line.ProjectRef its lines carry}; one tag per project or the move stops."""
    seen = {i: set() for i in ids}
    for recs in docs.values():
        for t in recs:
            for ln in t.get("Line") or []:
                d = _det(ln)
                v = d and (d.get("CustomerRef") or {}).get("value")
                if v in seen and (ln.get("ProjectRef") or {}).get("value"):
                    seen[v].add(ln["ProjectRef"]["value"])
    out = {}
    for i, tags in seen.items():
        if len(tags) > 1:
            raise PlanError(f"customer {i}: its lines carry {len(tags)} different project tags - fix by hand")
        out[i] = next(iter(tags), None)
    return out


def item_map(items: list, recode: dict) -> dict:
    """{old code -> new Item ref}; each new code must be exactly one QuickBooks item."""
    out = {}
    for old, new in recode.items():
        hits = [i for i in items if i.get("Active", True) and (i.get("Name") or "").upper() == new]
        if len(hits) != 1:
            raise PlanError(f"cost code {new}: expected exactly one QuickBooks item, found {len(hits)}")
        out[old] = {"value": hits[0]["Id"], "name": hits[0].get("Name")}
    return out


def _parse_recode(recode) -> dict:
    if isinstance(recode, dict):
        pairs = recode.items()
    else:
        pairs = (x.split("=", 1) for x in str(recode or "").split(",") if "=" in x)
    return {str(o).strip().upper(): str(n).strip().upper() for o, n in pairs if str(o).strip() and str(n).strip()}


def _build(src_p: str, dst_p: str, recode) -> dict:
    """Everything a preview or a write needs, read from the mirror."""
    src_p, dst_p = (src_p or "").upper().strip(), (dst_p or "").upper().strip()
    if not src_p or not dst_p:
        raise PlanError("pick both projects")
    access, cid = _auth()
    customers = qbo_api.query_all(access, cid, "Customer")
    src, dst = find_customer(customers, src_p), find_customer(customers, dst_p)
    if src["Id"] == dst["Id"]:
        raise PlanError("From and To are the same project")
    rc = _parse_recode(recode)
    items = item_map(qbo_api.query_all(access, cid, "Item"), rc) if rc else {}
    docs = {ent: qbo_api.query_all(access, cid, ent) for ent in ENTITIES}
    tags = project_tags(docs, (src["Id"], dst["Id"]))
    prefs = qbo_api._api_get(f"/v3/company/{cid}/preferences", access).get("Preferences", {})
    closed = (prefs.get("AccountingInfoPrefs") or {}).get("BookCloseDate") or ""

    rows = []
    for ent, recs in docs.items():
        for t in recs:
            for ln in t.get("Line") or []:
                d = _det(ln)
                if d and (d.get("CustomerRef") or {}).get("value") == src["Id"]:
                    code = _code(d)
                    date = t.get("TxnDate") or ""
                    rows.append({"key": _key(ent, t["Id"], str(ln.get("Id") or "")), "entity": ent, "id": t["Id"],
                                 "line": str(ln.get("Id") or ""), "doc": t.get("DocNumber") or "", "date": date,
                                 "vendor": (t.get("VendorRef") or t.get("EntityRef") or {}).get("name") or "",
                                 "code": code, "new_code": (items.get(code.upper()) or {}).get("name") or "",
                                 "amount": round(float(ln.get("Amount") or 0), 2),
                                 "signed": round((-1 if ent == "VendorCredit" else 1) * float(ln.get("Amount") or 0), 2),
                                 "doc_lines": len([x for x in t.get("Line") or [] if _det(x)]),
                                 "desc": (ln.get("Description") or "")[:120],
                                 "closed": bool(closed and date and date <= closed)})
    je = []
    for t in qbo_api.query_all(access, cid, "JournalEntry"):
        for ln in t.get("Line") or []:
            ref = ((ln.get("JournalEntryLineDetail") or {}).get("Entity") or {}).get("EntityRef") or {}
            if ref.get("value") == src["Id"]:
                je.append({"id": t["Id"], "doc": t.get("DocNumber") or "", "date": t.get("TxnDate") or "",
                           "amount": round(float(ln.get("Amount") or 0), 2), "desc": (ln.get("Description") or "")[:120]})
    rows.sort(key=lambda r: (r["date"], r["entity"], r["id"], r["line"]))
    return {"access": access, "cid": cid, "src": src, "dst": dst, "src_p": src_p, "dst_p": dst_p,
            "recode": items, "dst_tag": tags[dst["Id"]], "closed": closed, "rows": rows, "je": je}


def plan(src_p: str, dst_p: str, recode="") -> dict:
    """The preview: every line on FROM. Writes nothing."""
    try:
        b = _build(src_p, dst_p, recode)
    except PlanError as e:
        return {"ok": False, "error": str(e)}
    open_rows = [r for r in b["rows"] if not r["closed"]]
    return {"ok": True, "from": b["src_p"], "to": b["dst_p"], "from_name": _fqn(b["src"]), "to_name": _fqn(b["dst"]),
            "closed_date": b["closed"], "rows": b["rows"], "je": b["je"],
            "recode": {k: v["name"] for k, v in b["recode"].items()},
            "codes_on_from": sorted({r["code"] for r in b["rows"] if r["code"]}),
            "n": len(open_rows), "amount": round(sum(r["signed"] for r in open_rows), 2)}


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


def commit(src_p: str, dst_p: str, recode, keys: list, n: int, amount: float) -> dict:
    """Write the lines the owner ticked - only if they are still exactly what the preview showed
    (every ticked line still on FROM and open, same count, same amount)."""
    try:
        b = _build(src_p, dst_p, recode)
    except PlanError as e:
        return {"ok": False, "error": str(e)}
    by_key = {r["key"]: r for r in b["rows"]}
    keys = [str(k) for k in keys or []]
    pick = [by_key[k] for k in keys if k in by_key and not by_key[k]["closed"]]
    if not keys or len(pick) != len(set(keys)) or len(pick) != int(n) or \
            abs(round(sum(r["signed"] for r in pick), 2) - float(amount)) > EPS:
        return {"ok": False, "error": "The lines changed in QuickBooks since the preview - preview again."}
    import presence                                       # noqa: PLC0415  (ledger-local: macOS' own Touch ID dialog)
    amt = round(sum(r["signed"] for r in pick), 2)
    ok, why = presence.confirm(f"move {len(pick)} cost lines ({amt:,.2f}) from {b['src_p']} to {b['dst_p']} in QuickBooks")
    if not ok:
        return {"ok": False, "error": f"not confirmed on this Mac ({why}) - nothing was written"}
    access, cid, src, dst = b["access"], b["cid"], b["src"], b["dst"]
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    docs = {}
    for r in pick:
        docs.setdefault((r["entity"], r["id"]), set()).add(r["line"])
    out, stopped = [], ""
    for (ent, tid), want in docs.items():
        live = _get(access, cid, ent, tid)
        doc = live.get("DocNumber") or tid
        if b["closed"] and live.get("TxnDate", "") <= b["closed"]:
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "skipped", "detail": "closed period"})
            continue
        moved = 0
        for ln in live.get("Line") or []:
            d = _det(ln)
            if d and str(ln.get("Id")) in want and (d.get("CustomerRef") or {}).get("value") == src["Id"]:
                d["CustomerRef"] = {"value": dst["Id"]}
                if b["dst_tag"]:
                    ln["ProjectRef"] = {"value": b["dst_tag"]}
                else:
                    ln.pop("ProjectRef", None)
                old = (d.get("ItemRef") or {}).get("name", "").upper()
                if old in b["recode"]:
                    d["ItemRef"] = dict(b["recode"][old])
                moved += 1
        if not moved:
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "already moved", "detail": ""})
            continue
        before = _fingerprint(live)
        (BACKUP_DIR / f"{stamp}_{ent}_{tid}.json").write_text(json.dumps(_get(access, cid, ent, tid), indent=1))
        body = {k: v for k, v in live.items() if k not in ("MetaData", "domain", "sparse")}
        try:
            after = _post(access, cid, ent, body)
        except RuntimeError as e:
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "refused", "detail": str(e)})
            stopped = f"{ent} {doc}: {e}"
            break
        landed = sum(1 for ln in after.get("Line") or [] if str(ln.get("Id")) in want
                     and ((_det(ln) or {}).get("CustomerRef") or {}).get("value") == dst["Id"])
        if landed != moved:
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "not moved",
                        "detail": f"QuickBooks kept {moved - landed} line(s) on {b['src_p']}"})
            stopped = f"{ent} {doc}: QuickBooks kept {moved - landed} line(s) on {b['src_p']}"
            break
        if _fingerprint(after) != before:
            out.append({"entity": ent, "id": tid, "doc": doc, "result": "money changed",
                        "detail": "total / balance / a line amount moved after the write - check it"})
            stopped = f"{ent} {doc}: an amount moved after the write - check it in QuickBooks"
            break
        out.append({"entity": ent, "id": tid, "doc": doc, "result": "moved", "detail": f"{moved} line(s)"})
    with open(BACKUP_DIR / "runs.jsonl", "a") as f:
        f.write(json.dumps({"at": stamp, "from": b["src_p"], "to": b["dst_p"],
                            "recode": {k: v["name"] for k, v in b["recode"].items()},
                            "lines": len(pick), "amount": round(sum(r["signed"] for r in pick), 2),
                            "results": out, "stopped": stopped}) + "\n")
    return {"ok": not stopped, "error": stopped or None, "results": out,
            "moved_docs": sum(1 for r in out if r["result"] == "moved"),
            "moved_lines": sum(int(r["detail"].split()[0]) for r in out if r["result"] == "moved"),
            "backup_dir": str(BACKUP_DIR)}


def main() -> None:
    ap = argparse.ArgumentParser(description="Reclassify: move cost lines from one project to another (line level).")
    ap.add_argument("src", help="project number to move OFF, e.g. RP7242-FTW")
    ap.add_argument("dst", help="project number to move ONTO, e.g. RP7242")
    ap.add_argument("--commit", action="store_true", help="write to QuickBooks (default: preview)")
    ap.add_argument("--recode", default="", help="also swap cost codes on the moved lines, e.g. FW2=SL2,FW3=SL3")
    a = ap.parse_args()
    try:
        p = plan(a.src, a.dst, a.recode)
    except qbo_api.AuthError as e:
        sys.exit(f"✗  {e}")
    if not p["ok"]:
        sys.exit(f"✗  {p['error']}")
    print(f"FROM {p['from_name']}\nTO   {p['to_name']}\n")
    for o, nw in p["recode"].items():
        print(f"cost code {o} -> {nw} on the moved lines")
    print(f"{'type':<12} {'doc #':<14} {'date':<10} {'vendor':<28} {'code':<22} {'amount':>12}  lines on doc")
    for r in p["rows"]:
        d = dt.date.fromisoformat(r["date"]).strftime("%m/%d/%Y") if r["date"] else ""
        code = r["code"] + (f" -> {r['new_code']}" if r["new_code"] else "")
        print(f"{r['entity']:<12} {r['doc'][:14]:<14} {d:<10} {r['vendor'][:28]:<28} {code[:22]:<22} "
              f"{r['signed']:>12,.2f}  {r['doc_lines']}{'  CLOSED' if r['closed'] else ''}")
    print(f"\n{p['n']} open line(s), net {p['amount']:,.2f}")
    for j in p["je"]:
        print(f"   journal entry {j['doc'] or j['id']}  {j['date']}  {j['amount']:,.2f} - not moved, fix by hand")
    if not a.commit:
        print("\nPreview - nothing written. Add --commit to move these lines.")
        return
    res = commit(a.src, a.dst, a.recode, [r["key"] for r in p["rows"] if not r["closed"]], p["n"], p["amount"])
    for r in res.get("results") or []:
        print(f"   {r['result']:<14} {r['entity']} {r['doc']}  {r['detail']}")
    if not res["ok"]:
        sys.exit(f"✗  {res['error']}")
    print(f"\nDone: {res['moved_lines']} line(s) on {res['moved_docs']} document(s) moved. Before copies: {res['backup_dir']}")


if __name__ == "__main__":
    main()
