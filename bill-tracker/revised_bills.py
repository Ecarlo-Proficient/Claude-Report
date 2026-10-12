#!/usr/bin/env python3
"""revised_bills.py - the REVISED-bill follow-up as a process (AP-13), 10/12/2026.

The owner: "did the clerk send the bill to get revised? when? have they responded? with the new
revised bill? is the new revised bill uploaded to QBO? that's the part where we can get in a pickle."

What it does, every AP run (and by hand):
  1. FIND every revised bill in the QuickBooks copy (shared/qbo_mirror): a bill whose memo carries
     the PM's `REVISED - <reason>` line (bill_rows.revised_reason - line 2, or line 1 for a bill
     entered under the QBO approval workflow). Nothing is read from QuickBooks directly.
  2. For each: the vendor's ORIGINAL total and WHEN the PM changed it (the mirror's change log:
     the last 'amount changed' edit), the open balance, the approval state, the project, and what
     the vendor's latest STATEMENT says about it (shared/statement_record - the reconciler's record:
     matched at the revised amount / still the old amount / not on the statement).
  3. FLAGS the pickles:  entered twice (another open bill from the same vendor with the same #, or
     the # plus a suffix, entered after the revision) · amount changed again after the clerk sent it
     · paid before the vendor re-issued · bill gone from QuickBooks.
  4. UPSERTS one row per bill on the Notion "Revised Bills" board (shared/revised_board, keyed by
     the QuickBooks bill id). The automation writes its own columns only; Stage and the four dates
     (Sent on · Replied on · Received on · Entered on) and Notes are the clerk's. Stage moves
     forward only, and only to "Done - statement agrees" or "Dropped" by this tool.

Runs from excel_bill_sync (the AP run, every 15 min on the office server) when ACB_REVISED_DS_ID is
set; `python3 revised_bills.py` previews (dry run), `--write` writes, `--since YYYY-MM-DD` widens the
scan (default: bills created or dated in the last 120 days - the REVISED follow-up began 09/16/2026).
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import bill_approval, qbo_api, qbo_mirror, revised_board as rb, statement_record as sr  # noqa: E402
from shared.notion_client import NotionClient, NotionError  # noqa: E402
sys.path.insert(0, str(Path(__file__).resolve().parent))
from bill_rows import revised_reason  # noqa: E402  (tool-local: the ONE reading of the REVISED memo line)

DEFAULT_DAYS = 120
SUFFIX_RE = re.compile(r"^(.+?)[\s\-_/]*(R|REV|REVISED|A|B|C|1|2|CC|CORR|CORRECTED)$", re.I)


def _ref_key(ref: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", (ref or "").upper())


def _iso_date(s: Optional[str]) -> str:
    return (s or "")[:10]


# ─────────────────────────── 1 + 2 + 3: find ───────────────────────────

def _memo_project(memo: str) -> str:
    m = qbo_api.PROJ_RE.search(memo or "")
    return m.group(1).upper() if m else ""


def _line_project(bill: dict) -> str:
    for ln in bill.get("Line") or []:
        det = ln.get("ItemBasedExpenseLineDetail") or ln.get("AccountBasedExpenseLineDetail") or {}
        name = ((det.get("CustomerRef") or {}).get("name")) or ""
        m = qbo_api.PROJ_RE.search(name)
        if m:
            return m.group(1).upper()
    return ""


def statement_says(vendor: str, ref: str, total: float, records: Optional[dict] = None) -> dict:
    """What the vendor's latest statement record shows for this bill #.
    -> {text, matched: bool|None, as_of}. matched=None when no record / not on the statement."""
    rec = None
    try:
        rec = (records or {}).get(sr.key(vendor)) if records is not None else sr.find(vendor)
    except Exception:                                    # noqa: BLE001 - share off = no answer, not a failure
        rec = None
    if not rec:
        return {"text": "no statement record for this vendor", "matched": None, "as_of": ""}
    as_of = sr.us_date(rec.get("as_of") or "")
    want = _ref_key(ref)
    hit = next((ln for ln in rec.get("lines") or [] if ln.get("ref") and _ref_key(ln["ref"]) == want), None)
    if not hit:
        return {"text": f"not on the {as_of} statement", "matched": None, "as_of": rec.get("as_of") or ""}
    amt = float(hit.get("amount") or 0)
    bucket = hit.get("bucket") or ""
    if bucket in ("matched", "approve", "checkqbo") and abs(amt - float(total or 0)) <= 0.01:
        return {"text": f"matched at ${amt:,.2f} on the {as_of} statement", "matched": True, "as_of": rec.get("as_of") or ""}
    if bucket == "lag":
        return {"text": f"paid in QBO, vendor still shows ${amt:,.2f} open on the {as_of} statement", "matched": False, "as_of": rec.get("as_of") or ""}
    return {"text": f"still ${amt:,.2f} on the {as_of} statement (QBO ${float(total or 0):,.2f})", "matched": False, "as_of": rec.get("as_of") or ""}


def find(con, since: Optional[dt.date] = None, today: Optional[dt.date] = None, records: Optional[dict] = None) -> List[dict]:
    """Every revised bill in the mirror, enriched. `records` = statement_record.read_all()["vendors"]
    style map (key -> full record) for tests; None reads the share per vendor."""
    today = today or dt.date.today()
    since = since or (today - dt.timedelta(days=DEFAULT_DAYS))
    s = since.isoformat()
    bills = {}
    by_vendor_refs: Dict[str, List[dict]] = {}
    for r in con.execute("SELECT id, txn_date, doc_number, total, created, last_updated, deleted, json FROM qbo_bill "
                         "WHERE (txn_date >= ? OR created >= ?)", (s, s)):
        try:
            j = qbo_mirror._unpack(r[7]) or {}
        except Exception:                                # noqa: BLE001
            continue
        vendor = ((j.get("VendorRef") or {}).get("name")) or ""
        rec = {"bill_id": str(r[0]), "txn_date": r[1] or "", "ref": r[2] or "", "total": float(r[3] or 0),
               "created": r[4] or "", "last_updated": r[5] or "", "deleted": bool(r[6]), "vendor": vendor,
               "memo": j.get("PrivateNote") or "", "balance": float(j.get("Balance") or 0), "json": j}
        by_vendor_refs.setdefault(sr.key(vendor), []).append(rec)
        created_d = _iso_date(rec["created"])
        reason = revised_reason(rec["memo"], allow_first_line=bool(created_d and created_d >= bill_approval.APPROVAL_WORKFLOW_START.isoformat()))
        if reason:
            bills[rec["bill_id"]] = dict(rec, reason=reason)
    # the change log: the vendor's original total + when the PM changed it (the LAST amount change)
    changes: Dict[str, List[dict]] = {}
    if bills:
        ids = list(bills)
        for i in range(0, len(ids), 500):
            chunk = ids[i:i + 500]
            for c in con.execute(f"SELECT rec_id, kind, changed_at, total_before, total_after, flags FROM mirror_change "
                                 f"WHERE entity='Bill' AND rec_id IN ({','.join('?' * len(chunk))}) ORDER BY changed_at", chunk):
                changes.setdefault(str(c[0]), []).append({"kind": c[1], "at": c[2] or "", "before": c[3], "after": c[4],
                                                          "flags": [f for f in (c[5] or "").split(",") if f]})
    out = []
    checked_at = dt.datetime.now().isoformat(timespec="minutes")
    for bid, b in bills.items():
        ch = changes.get(bid, [])
        amount_edits = [c for c in ch if c["kind"] == "edited" and "amount changed" in c["flags"]]
        first_total = next((c["after"] for c in ch if c["kind"] == "created"), None)
        original = amount_edits[0]["before"] if amount_edits else first_total
        revised_at = amount_edits[-1]["at"] if amount_edits else (b["last_updated"] or b["created"])
        approval = bill_approval.approval_state(b["memo"], _iso_date(b["created"]), b["balance"])
        project = _memo_project(b["memo"]) or _line_project(b["json"])
        st = statement_says(b["vendor"], b["ref"], b["total"], records)
        flags: List[str] = []
        if b["deleted"]:
            flags.append("bill gone from QuickBooks")
        # entered twice: another live bill from the same vendor, same # (or # + suffix), entered after the revision
        want = _ref_key(b["ref"])
        base = SUFFIX_RE.match(b["ref"] or "")
        for o in by_vendor_refs.get(sr.key(b["vendor"]), []):
            if o["bill_id"] == bid or o["deleted"]:
                continue
            ok = _ref_key(o["ref"])
            m2 = SUFFIX_RE.match(o["ref"] or "")
            same = want and (ok == want or (m2 and _ref_key(m2.group(1)) == want) or (base and _ref_key(base.group(1)) == ok))
            if same and _iso_date(o["created"]) >= _iso_date(revised_at):
                flags.append(f"entered twice? bill {o['ref']} ({o['txn_date']}, ${o['total']:,.2f}) entered {_iso_date(o['created'])}")
        if b["balance"] <= 0.005 and not b["deleted"] and st["matched"] is not True:
            flags.append("paid before the vendor re-issued")
        out.append({"bill_id": bid, "vendor": b["vendor"], "ref": b["ref"], "bill_date": b["txn_date"], "project": project,
                    "reason": b["reason"], "total": b["total"], "original": original, "open": b["balance"], "approval": approval,
                    "revised_at": _iso_date(revised_at), "revised_at_full": revised_at, "deleted": b["deleted"],
                    "amount_edits": [{"at": c["at"], "before": c["before"], "after": c["after"]} for c in amount_edits],
                    "statement_says": st["text"], "statement_matched": st["matched"], "statement_as_of": st["as_of"],
                    "flags": flags, "flags_text": "", "checked_at": checked_at})
    out.sort(key=lambda r: (r["vendor"].upper(), r["bill_date"]))
    return out


# ─────────────────────────── 4: the board ───────────────────────────

def plan(found: List[dict], prior: List[dict], today: Optional[dt.date] = None) -> List[dict]:
    """Pure: what to write for each found bill, given the board as it stands.
    -> [{action: create|update|keep, row, page_id, stage_to, why}]. Stage: forward only; the
    automation sets Done when the statement agrees, Dropped when the bill is gone; a stage the
    clerk set stays. 'amount changed again after sent' is judged here (needs the clerk's date)."""
    by_key = {p.get("key"): p for p in prior if p.get("key")}
    out = []
    for r in found:
        p = by_key.get(r["bill_id"])
        flags = list(r["flags"])
        stage_now = (p or {}).get("stage") or ""
        sent = (p or {}).get("sent_on") or ""
        if sent and any(_iso_date(e["at"]) > sent for e in r["amount_edits"]):
            flags.append(f"amount changed again in QuickBooks after the clerk sent it ({sent[5:7]}/{sent[8:10]})")
        row = dict(r, flags=flags, flags_text=rb.flags_text(flags))
        stage_to = None
        if not stage_now:
            stage_to = "To send"
        if r["deleted"] and stage_now not in rb.FINAL:
            stage_to = "Dropped"
        elif r["statement_matched"] is True and stage_now not in rb.FINAL:
            stage_to = "Done - statement agrees"
        if stage_to and stage_now and rb.STAGE_RANK.get(stage_to, 0) < rb.STAGE_RANK.get(stage_now, 0):
            stage_to = None                              # never backwards
        if p is None:
            out.append({"action": "create", "row": row, "page_id": None, "stage_to": stage_to or "To send", "why": "new on the board"})
        else:
            changed = any(str(p.get(rb._key(k)) or "") != str(row.get(src) or "") for k, src in
                          (("Statement says", "statement_says"), ("Flags", "flags_text"), ("Approval", "approval"), ("PM reason", "reason")))
            changed = changed or any(_num(p.get(rb._key(k))) != _num(row.get(src)) for k, src in
                                     (("QBO total", "total"), ("Open balance", "open"), ("Vendor original", "original")))
            out.append({"action": "update" if (changed or stage_to) else "keep", "row": row, "page_id": p["page_id"],
                        "stage_to": stage_to, "why": "stage -> " + stage_to if stage_to else ("columns changed" if changed else "no change")})
    return out


def _num(v):
    return None if v is None or v == "" else round(float(v), 2)


def run(write: bool = False, since: Optional[dt.date] = None, log=print) -> dict:
    """Find + plan + (write). Never raises past itself: the AP run must not fail on Notion."""
    ds = rb.ds_id()
    if not ds:
        log("  revised bills: board off (set ACB_REVISED_DS_ID) - skipped")
        return {"ok": False, "skipped": "no board id"}
    try:
        con = qbo_mirror.connect()
    except Exception as e:                               # noqa: BLE001
        log(f"  revised bills: no QuickBooks copy ({e}) - skipped")
        return {"ok": False, "skipped": str(e)}
    try:
        found = find(con, since=since)
    finally:
        con.close()
    client = prior = None
    try:
        client = NotionClient()
        prior = rb.read_rows(client, ds)
    except (NotionError, Exception) as e:                # noqa: BLE001
        log(f"  revised bills: cannot read the Notion board ({e}) - " + ("nothing written" if write else "preview without it"))
        if write:
            return {"ok": False, "found": found, "error": str(e)}
        prior = []
    acts = plan(found, prior or [])
    n_new = sum(1 for a in acts if a["action"] == "create")
    n_upd = sum(1 for a in acts if a["action"] == "update")
    log(f"  revised bills: {len(found)} revised bill(s) in QuickBooks · board has {len(prior or [])} · "
        f"{n_new} new · {n_upd} to update" + ("" if write else " (dry run)"))
    for a in acts:
        r = a["row"]
        st = rb.standing({"stage": a["stage_to"] or next((p.get("stage") for p in prior or [] if p.get("key") == r["bill_id"]), ""),
                          "revised_in_qbo": r["revised_at"]})
        log(f"    {a['action']:6s} {r['vendor'][:28]:28s} {r['ref']:14s} ${r['total']:>11,.2f}"
            + (f" (was ${r['original']:,.2f})" if r.get("original") not in (None, r["total"]) else "")
            + f"  {st['stage']} · {r['statement_says']}" + (f"  !! {r['flags_text']}" if r["flags_text"] else ""))
    written = 0
    if write and client is not None:
        for a in acts:
            if a["action"] == "keep":
                continue
            props = rb.props_auto(a["row"])
            if a["stage_to"]:
                props.update(rb.props_stage(a["stage_to"]))
            try:
                if a["action"] == "create":
                    client.create_page(ds, props)
                else:
                    client.update_page(a["page_id"], props)
                written += 1
            except NotionError as e:
                log(f"    Notion: {a['row']['vendor']} {a['row']['ref']}: {e}")
        log(f"  revised bills: {written} row(s) written to the Notion board")
    return {"ok": True, "found": found, "plan": acts, "written": written}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--write", action="store_true", help="write the Notion board (default: preview only)")
    ap.add_argument("--since", default="", metavar="YYYY-MM-DD", help=f"scan bills created or dated since (default: {DEFAULT_DAYS} days back)")
    ap.add_argument("--json", action="store_true", help="print the found bills as JSON")
    args = ap.parse_args()
    since = dt.date.fromisoformat(args.since) if args.since else None
    res = run(write=args.write, since=since)
    if args.json:
        print(json.dumps(res.get("found", []), indent=1, default=str))
    return 0 if res.get("ok") or res.get("skipped") else 1


if __name__ == "__main__":
    sys.exit(main())
