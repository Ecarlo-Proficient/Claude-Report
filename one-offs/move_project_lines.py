#!/usr/bin/env python3
"""
move_project_lines.py - move every cost LINE coded to one QuickBooks project onto another.

Line level only (the owner 10/06/2026): a sub's bill that spans several jobs keeps its other lines
exactly as they are - only the lines whose project is FROM change, and only their project. Amount,
account / cost code, class, description and the bill's payments are left alone.

    one-offs/move_project_lines.py RP7242-FTW RP7242            dry run: every line that would move
    one-offs/move_project_lines.py RP7242-FTW RP7242 --commit   write it to QuickBooks
    ... --recode FW2=SL2,FW3=SL3                                also swap these cost codes on the moved lines

Covers bills, expenses / checks (Purchase) and vendor credits. Journal entry lines are listed, never
written - those are fixed by hand.

Guards (live from QuickBooks at write time, one document at a time):
  - the document is re-read right before the write (fresh SyncToken); a line no longer on FROM is left
  - nothing dated on or before the books-closed date is touched (listed instead)
  - the live document is saved to ~/Library/Logs/Proficient/move-project-lines/ before its write
  - after the write the total, open balance and every line amount must match what was there before,
    and every moved line must come back on TO; if not, the run stops and names the document
QuickBooks Projects keeps a second project tag on each line (`Line.ProjectRef`, a Projects id, not the
customer id) and resets the line's customer to match it - a CustomerRef-only edit is silently undone
(found 10/06/2026). The tool reads each project's tag off the lines already on it and moves both.
Projects are matched by the EXACT project number (RP7242 and RP7242-FTW are different projects).
"""

import argparse
import csv
import datetime as dt
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import requests  # noqa: E402

from shared import paths, qbo_api  # noqa: E402
from shared.xlsx_guard import csv_cell  # noqa: E402

ENTITIES = ("Bill", "Purchase", "VendorCredit")
DETAILS = ("AccountBasedExpenseLineDetail", "ItemBasedExpenseLineDetail")
BACKUP_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "move-project-lines"


def _proj(name: str) -> str:
    """The project number a customer name carries (`Parent:RP7242-FTW 123 Main` -> RP7242-FTW)."""
    leaf = (name or "").split(":")[-1].strip().upper()
    return leaf.split()[0] if leaf else ""


def find_customer(customers: list, proj: str) -> dict:
    hits = [c for c in customers if _proj(c.get("FullyQualifiedName") or c.get("DisplayName") or "") == proj]
    if len(hits) != 1:
        names = ", ".join(c.get("FullyQualifiedName") or "" for c in hits) or "none"
        sys.exit(f"✗  {proj}: expected exactly one QuickBooks customer, found {len(hits)} ({names})")
    return hits[0]


def _code(det: dict) -> str:
    return (det.get("ItemRef") or {}).get("name") or (det.get("AccountRef") or {}).get("name") or ""


def project_tags(access: str, cid: str, ids: tuple) -> dict:
    """{customer id -> the Line.ProjectRef its lines carry}; one tag per project or the run stops."""
    seen = {i: set() for i in ids}
    for ent in ENTITIES:
        for t in qbo_api.query_all(access, cid, ent):
            for ln in t.get("Line") or []:
                dk = next((k for k in DETAILS if k in ln), None)
                v = dk and (ln[dk].get("CustomerRef") or {}).get("value")
                if v in seen and (ln.get("ProjectRef") or {}).get("value"):
                    seen[v].add(ln["ProjectRef"]["value"])
    out = {}
    for i, tags in seen.items():
        if len(tags) > 1:
            sys.exit(f"✗  customer {i}: its lines carry {len(tags)} different project tags - fix by hand")
        out[i] = next(iter(tags), None)
    return out


def scan(access: str, cid: str, src_id: str) -> tuple:
    """Every line on src_id across the cost documents (from the mirror) + the journal-entry lines."""
    rows, je = [], []
    for ent in ENTITIES:
        for t in qbo_api.query_all(access, cid, ent):
            for ln in t.get("Line") or []:
                dk = next((k for k in DETAILS if k in ln), None)
                if dk and (ln[dk].get("CustomerRef") or {}).get("value") == src_id:
                    rows.append({"entity": ent, "id": t["Id"], "doc": t.get("DocNumber") or "",
                                 "date": t.get("TxnDate") or "",
                                 "vendor": (t.get("VendorRef") or t.get("EntityRef") or {}).get("name") or "",
                                 "line": str(ln.get("Id") or ""), "code": _code(ln[dk]),
                                 "amount": float(ln.get("Amount") or 0),
                                 "doc_lines": len([x for x in t.get("Line") or [] if any(k in x for k in DETAILS)]),
                                 "desc": (ln.get("Description") or "")[:70]})
    for t in qbo_api.query_all(access, cid, "JournalEntry"):
        for ln in t.get("Line") or []:
            ent = ((ln.get("JournalEntryLineDetail") or {}).get("Entity") or {}).get("EntityRef") or {}
            if ent.get("value") == src_id:
                je.append({"id": t["Id"], "doc": t.get("DocNumber") or "", "date": t.get("TxnDate") or "",
                           "amount": float(ln.get("Amount") or 0), "desc": (ln.get("Description") or "")[:70]})
    rows.sort(key=lambda r: (r["date"], r["entity"], r["id"], r["line"]))
    return rows, je


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
        raise RuntimeError(f"QuickBooks refused the update ({r.status_code}): {msg} - nothing was changed")
    return r.json()[ent]


def _fingerprint(t: dict) -> tuple:
    return (round(float(t.get("TotalAmt") or 0), 2), round(float(t.get("Balance") or 0), 2),
            tuple(sorted((str(ln.get("Id")), round(float(ln.get("Amount") or 0), 2)) for ln in t.get("Line") or [])))


def item_map(access: str, cid: str, recode: dict) -> dict:
    """{old code -> new Item ref} for --recode; each new code must be exactly one QuickBooks item."""
    items = qbo_api.query_all(access, cid, "Item")
    out = {}
    for old, new in recode.items():
        hits = [i for i in items if (i.get("Name") or "").upper() == new]
        if len(hits) != 1:
            sys.exit(f"✗  cost code {new}: expected exactly one QuickBooks item, found {len(hits)}")
        out[old] = {"value": hits[0]["Id"], "name": hits[0].get("Name")}
    return out


def commit(access: str, cid: str, rows: list, src: dict, dst: dict, closed, recode: dict, dst_tag) -> list:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out = []
    docs = {}
    for r in rows:
        docs.setdefault((r["entity"], r["id"]), set()).add(r["line"])
    for (ent, tid), want in docs.items():
        live = _get(access, cid, ent, tid)
        if closed and live.get("TxnDate", "") <= closed:
            out.append((ent, tid, live.get("DocNumber") or "", "SKIP_CLOSED", live.get("TxnDate")))
            continue
        moved = 0
        for ln in live.get("Line") or []:
            dk = next((k for k in DETAILS if k in ln), None)
            if dk and str(ln.get("Id")) in want and (ln[dk].get("CustomerRef") or {}).get("value") == src["Id"]:
                ln[dk]["CustomerRef"] = {"value": dst["Id"]}
                if dst_tag:
                    ln["ProjectRef"] = {"value": dst_tag}
                elif "ProjectRef" in ln:
                    del ln["ProjectRef"]
                old = (ln[dk].get("ItemRef") or {}).get("name", "").upper()
                if old in recode:
                    ln[dk]["ItemRef"] = dict(recode[old])
                moved += 1
        if not moved:
            out.append((ent, tid, live.get("DocNumber") or "", "ALREADY_MOVED", ""))
            continue
        before = _fingerprint(live)
        (BACKUP_DIR / f"{stamp}_{ent}_{tid}.json").write_text(json.dumps(_get(access, cid, ent, tid), indent=1))
        body = {k: v for k, v in live.items() if k not in ("MetaData", "domain", "sparse")}
        after = _post(access, cid, ent, body)
        landed = sum(1 for ln in after.get("Line") or [] if str(ln.get("Id")) in want
                     and ((ln.get(next((k for k in DETAILS if k in ln), ""), {}) or {}).get("CustomerRef") or {}).get("value")
                     == dst["Id"])
        if landed != moved:
            out.append((ent, tid, live.get("DocNumber") or "", "NOT_MOVED", f"{landed} of {moved} line(s) landed"))
            print(f"✗  {ent} {live.get('DocNumber') or tid}: QuickBooks kept {moved - landed} line(s) on the old "
                  f"project - stopping. Before copy: {BACKUP_DIR}")
            break
        if _fingerprint(after) != before:
            out.append((ent, tid, live.get("DocNumber") or "", "CHANGED_MONEY", "STOPPED - check this document"))
            print(f"✗  {ent} {live.get('DocNumber') or tid}: totals moved after the write - stopping. "
                  f"Before copy: {BACKUP_DIR}")
            break
        out.append((ent, tid, live.get("DocNumber") or "", "MOVED", f"{moved} line(s)"))
        print(f"   moved {ent} {live.get('DocNumber') or tid}: {moved} line(s)")
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="Move every cost line from one project to another (line level).")
    ap.add_argument("src", help="project number to move OFF, e.g. RP7242-FTW")
    ap.add_argument("dst", help="project number to move ONTO, e.g. RP7242")
    ap.add_argument("--commit", action="store_true", help="write to QuickBooks (default: dry run)")
    ap.add_argument("--recode", default="", help="also swap cost codes on the moved lines, e.g. FW2=SL2,FW3=SL3")
    a = ap.parse_args()
    src_p, dst_p = a.src.upper().strip(), a.dst.upper().strip()

    access, cid = qbo_api._bearer_exchange()
    customers = qbo_api.query_all(access, cid, "Customer")
    src, dst = find_customer(customers, src_p), find_customer(customers, dst_p)
    if src["Id"] == dst["Id"]:
        sys.exit("✗  FROM and TO are the same customer")
    print(f"FROM {src.get('FullyQualifiedName')}\nTO   {dst.get('FullyQualifiedName')}\n")

    recode = item_map(access, cid, {o.strip().upper(): n.strip().upper() for o, n in
                                    (x.split("=") for x in a.recode.split(",") if "=" in x)})
    for o, n in recode.items():
        print(f"cost code {o} -> {n['name']} on the moved lines")
    tags = project_tags(access, cid, (src["Id"], dst["Id"]))
    if tags[src["Id"]] and not tags[dst["Id"]]:
        sys.exit(f"✗  {dst_p} has no lines yet, so its project tag is unknown - move one line by hand first")
    rows, je = scan(access, cid, src["Id"])
    prefs = qbo_api._api_get(f"/v3/company/{cid}/preferences", access).get("Preferences", {})
    closed = (prefs.get("AccountingInfoPrefs") or {}).get("BookCloseDate") or ""

    sign = {"VendorCredit": -1}
    print(f"{'type':<12} {'doc #':<14} {'date':<10} {'vendor':<28} {'code':<22} {'amount':>12}  lines on doc")
    for r in rows:
        d = dt.date.fromisoformat(r["date"]).strftime("%m/%d/%Y") if r["date"] else ""
        flag = "  CLOSED" if closed and r["date"] <= closed else ""
        print(f"{r['entity']:<12} {r['doc'][:14]:<14} {d:<10} {r['vendor'][:28]:<28} {(r['code'] + (' -> ' + recode[r['code'].upper()]['name'] if r['code'].upper() in recode else ''))[:22]:<22} "
              f"{sign.get(r['entity'], 1) * r['amount']:>12,.2f}  {r['doc_lines']}{flag}")
    net = sum(sign.get(r["entity"], 1) * r["amount"] for r in rows)
    print(f"\n{len(rows)} line(s) on {len({(r['entity'], r['id']) for r in rows})} document(s), net {net:,.2f}")
    if je:
        print(f"\n{len(je)} journal entry line(s) on {src_p} - NOT moved, fix by hand:")
        for j in je:
            print(f"   JE {j['doc'] or j['id']}  {j['date']}  {j['amount']:,.2f}  {j['desc']}")

    out_dir = paths.analysis_dir(f"move {src_p} to {dst_p}")
    out_dir.mkdir(parents=True, exist_ok=True)
    plan = out_dir / "lines.csv"
    with open(plan, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["type", "id", "doc #", "date", "vendor", "line", "cost code", "amount", "lines on doc", "description"])
        for r in rows:
            w.writerow([csv_cell(v) for v in (r["entity"], r["id"], r["doc"], r["date"], r["vendor"], r["line"], r["code"],
                                              r["amount"], r["doc_lines"], r["desc"])])
    print(f"\nList: {plan}")

    if not a.commit:
        print("\nDry run - nothing written. Add --commit to move these lines.")
        return
    res = commit(access, cid, [r for r in rows if not (closed and r["date"] <= closed)], src, dst, closed, recode, tags[dst["Id"]])
    with open(out_dir / "results.csv", "w", newline="") as f:
        csv.writer(f).writerows([("type", "id", "doc #", "result", "detail")] + [[csv_cell(v) for v in r] for r in res])
    print(f"\nDone: {sum(1 for r in res if r[3] == 'MOVED')} document(s) moved. Results: {out_dir / 'results.csv'}")


if __name__ == "__main__":
    main()
