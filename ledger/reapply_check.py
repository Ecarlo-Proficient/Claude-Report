#!/usr/bin/env python3
"""
reapply_check.py - put a stripped check back on the bills it paid.

QBO sometimes strips a paid check: an edit or delete on one bill unapplies the
WHOLE payment, and the money floats (Company > Checks QBO changed in the ledger).
The mirror's change log kept the check as it was BEFORE the strip, listing every
bill it paid and how much. This script reads that before copy and writes those
lines back onto the check, so nobody has to pick 129 bills out of Pay Bills by hand.

    ledger/reapply_check.py 25745              dry run: what would go back on
    ledger/reapply_check.py 25745 --commit     write it to QBO

The ledger's Checks QBO changed page runs the same two steps from a check's card
(Re-apply in QuickBooks -> dry run -> "Are you sure?" -> write): dry_run() / commit().

Guards (a line is restored only when ALL hold, live from QBO at run time):
  - the bill (or vendor credit) still exists, same vendor as the check
  - its open balance still covers the amount the check paid on it
    (paid elsewhere since -> skipped and listed, never partially applied)
  - lines already back on the check are kept as they are, never doubled
  - the restored total never exceeds the check amount
  - the check is re-read right before the write (fresh SyncToken)
The live check is saved to ~/Library/Logs/Proficient/reapply-check/ before any write.

CLEAR COPY - the bulk fix (owner 2026-10-06: "a filter that the script can 100% GUARANTEE that it has a copy
of the check prior to qbo changing it ... be careful with credits and shortpays"). A check is a clear copy
only when every one of these is proven, from QuickBooks' own records, to the cent:
  - the change log holds the check as QuickBooks had it before the strip, and that copy is WHOLE: its bills
    less its credits add up to the check amount (nothing was floating even then)
  - every copy of the check on file agrees line for line (no hand re-apply in between that differs)
  - the check itself is untouched: same amount, date, vendor, bank account, pay type
  - nothing new is on the check now (QuickBooks did not drop its money on a later bill)
  - only bills and vendor credits - a journal entry line is fixed by hand
  - every bill / credit has its own copy from while this check was on it, the same total today, and after
    the fix its open balance lands EXACTLY where that copy had it. A short pay keeps the same open amount it
    had before; a credit goes back to the same balance (so it was not used on another check since)
  - after the fix the check is applied to the cent, nothing floating
The page filter (`annotate`) reads the mirror; the bulk dry run and the bulk write re-prove every check
LIVE from QuickBooks (`bulk_dry_run` / `bulk_commit`), one check at a time, and skip any that no longer prove.

    ledger/reapply_check.py --clear             every clear copy, dry run, live
    ledger/reapply_check.py --clear --commit    write every one that still proves
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import requests  # noqa: E402

from shared import qbo_api, qbo_mirror as mirror  # noqa: E402

EPS = 0.005
LOG_DIR = Path.home() / "Library" / "Logs" / "Proficient" / "reapply-check"
ENTITY = {"Bill": "Bill", "VendorCredit": "VendorCredit"}


def _money(x: float) -> str:
    return f"{x:,.2f}"


def _fmt(d: str | None) -> str:
    try:
        return dt.date.fromisoformat(str(d)[:10]).strftime("%m/%d/%Y")
    except ValueError:
        return str(d or "")


def _links(pay: dict) -> list:
    """[(TxnType, TxnId, amount)] - what the payment applies, line by line."""
    out = []
    for ln in pay.get("Line") or []:
        for lt in ln.get("LinkedTxn") or []:
            out.append((lt.get("TxnType"), str(lt.get("TxnId")), round(float(ln.get("Amount") or 0), 2)))
    return out


def _find_before(con, check: str, payment_id: str | None, change_id: int | None):
    """(payment id, before copy, change row): the fullest pre-strip copy the change log holds."""
    if payment_id is None:
        hits = [p for p in mirror.load("BillPayment", con=con) if str(p.get("DocNumber") or "") == check]
        if not hits:
            raise LookupError(f"check {check}: no bill payment with that number in the mirror")
        logged = {c["rec_id"] for c in mirror.changes(con, entity="BillPayment", limit=100000) if c.get("has_before")}
        if len(hits) > 1 and sum(p["Id"] in logged for p in hits) == 1:
            hits = [p for p in hits if p["Id"] in logged]   # check numbers repeat across years and vendors
        if len(hits) > 1:
            listing = "; ".join(f"id {p['Id']} {_fmt(p.get('TxnDate'))} {(p.get('VendorRef') or {}).get('name')}"
                                f" {_money(float(p.get('TotalAmt') or 0))}" for p in hits)
            raise LookupError(f"check {check} matches {len(hits)} payments ({listing}) - rerun with --payment-id")
        payment_id = hits[0]["Id"]
    rows = [c for c in mirror.changes(con, entity="BillPayment", limit=100000)
            if c["rec_id"] == payment_id and c.get("has_before")]
    if change_id is not None:
        rows = [c for c in rows if c["id"] == change_id]
    if not rows:
        snap = _from_snapshot(con, payment_id)
        if snap:
            return payment_id, snap[0], snap[1]
        raise LookupError(f"payment {payment_id}: no before copy in the change log (stripped before it started "
                          "09/23/2026) and no ledger snapshot whose bills add up to the check - re-apply by hand")
    best = max(rows, key=lambda c: (c.get("lines_before") or 0, c["changed_at"]))
    return payment_id, mirror.change_before(con, best["id"]), best


def _from_snapshot(con, payment_id: str):
    """A check stripped BEFORE the change log (09/23/2026): the bills it paid from a saved copy of the ledger
    (its `bill_payment_line` table, loaded from QBO while the check was still whole). Used only when that
    copy's lines add up to the check amount to the cent, and only copies saved BEFORE QBO last changed the check
    (a later copy holds the stripped state). Newest copy first."""
    pay = next((p for p in mirror.load("BillPayment", con=con) if p["Id"] == payment_id), None)
    if not pay:
        return None
    total = round(float(pay.get("TotalAmt") or 0), 2)
    changed = dt.datetime.fromisoformat(str((pay.get("MetaData") or {}).get("LastUpdatedTime")).replace("Z", "+00:00"))
    for f in sorted(mirror.db_path().parent.glob("ledger*.sqlite3*"), key=lambda f: f.stat().st_mtime, reverse=True):
        taken = dt.datetime.fromtimestamp(f.stat().st_mtime, dt.timezone.utc)
        if taken >= changed:                   # a copy saved AFTER QBO last changed the check shows the stripped state, not the evidence
            continue
        try:
            lc = sqlite3.connect(f"file:{f}?mode=ro", uri=True)
            try:
                got = lc.execute("SELECT bill_id, amount FROM bill_payment_line WHERE payment_id = ?", (payment_id,)).fetchall()
            finally:
                lc.close()
        except sqlite3.Error:
            continue
        if got and abs(round(sum(a for _, a in got), 2) - total) < EPS:
            when = dt.datetime.fromtimestamp(f.stat().st_mtime).isoformat(timespec="seconds")
            before = {"Line": [{"Amount": round(a, 2), "LinkedTxn": [{"TxnId": str(b), "TxnType": "Bill"}]} for b, a in got]}
            return before, {"changed_at": when, "source": f"ledger snapshot {f.name}"}
    return None


def _live(access: str, cid: str, typ: str, ids: list) -> dict:
    out = {}
    for i in range(0, len(ids), 40):
        chunk = ",".join(f"'{x}'" for x in ids[i:i + 40])
        q = f"SELECT * FROM {typ} WHERE Id IN ({chunk}) MAXRESULTS 1000"
        for r in qbo_api._api_get(f"/v3/company/{cid}/query", access, {"query": q}).get("QueryResponse", {}).get(typ, []):
            out[str(r["Id"])] = r
    return out


def _get_payment(access: str, cid: str, pid: str) -> dict:
    return qbo_api._api_get(f"/v3/company/{cid}/billpayment/{pid}", access)["BillPayment"]


def plan(access: str, cid: str, pid: str, before: dict, live: dict | None = None) -> dict:
    live = live or _get_payment(access, cid, pid)
    vendor = str((live.get("VendorRef") or {}).get("value") or "")
    total = round(float(live.get("TotalAmt") or 0), 2)
    now = {(t, i): a for t, i, a in _links(live)}
    want = [(t, i, a) for t, i, a in _links(before)]

    recs = {}
    for typ in ENTITY:
        ids = sorted({i for t, i, _ in want if t == typ})
        if ids:
            recs[typ] = _live(access, cid, typ, ids)

    restore, skipped, already = [], [], []
    for typ, tid, amt in want:
        if now.get((typ, tid), 0) >= amt - EPS:
            already.append((typ, tid, amt))
            continue
        if typ not in ENTITY:
            skipped.append((typ, tid, amt, f"{typ} line - re-apply by hand"))
            continue
        r = recs.get(typ, {}).get(tid)
        if not r:
            skipped.append((typ, tid, amt, "deleted in QBO"))
            continue
        doc = r.get("DocNumber") or tid
        if str((r.get("VendorRef") or {}).get("value") or "") != vendor:
            skipped.append((typ, doc, amt, "different vendor now"))
            continue
        need = round(amt - now.get((typ, tid), 0), 2)
        bal = float(r.get("Balance") or 0)
        if bal < need - EPS:
            skipped.append((typ, doc, amt, f"open balance {_money(bal)} < {_money(need)} (paid elsewhere since?)"))
            continue
        restore.append({"type": typ, "id": tid, "doc": doc, "date": r.get("TxnDate"), "amount": amt, "need": need,
                        "balance": bal, "doc_total": round(float(r.get("TotalAmt") or 0), 2),
                        "memo": (r.get("PrivateNote") or "").split("\n")[0][:60]})

    lines = [dict(ln) for ln in live.get("Line") or []]
    have = {(t, i) for t, i, _ in _links(live)}
    for x in restore:
        if (x["type"], x["id"]) in have:        # partially on already: raise that line to the before amount
            for ln in lines:
                if any(lt.get("TxnType") == x["type"] and str(lt.get("TxnId")) == x["id"] for lt in ln.get("LinkedTxn") or []):
                    ln["Amount"] = x["amount"]
        else:
            lines.append({"Amount": x["amount"], "LinkedTxn": [{"TxnId": x["id"], "TxnType": x["type"]}]})

    def net(ls):
        b = sum(float(ln.get("Amount") or 0) for ln in ls
                for lt in (ln.get("LinkedTxn") or [])[:1] if lt.get("TxnType") != "VendorCredit")
        c = sum(float(ln.get("Amount") or 0) for ln in ls
                for lt in (ln.get("LinkedTxn") or [])[:1] if lt.get("TxnType") == "VendorCredit")
        return round(b - c, 2)

    return {"live": live, "total": total, "applied_now": net(live.get("Line") or []), "applied_after": net(lines),
            "lines": lines, "restore": restore, "skipped": skipped, "already": already}


def post(access: str, cid: str, live: dict, lines: list) -> dict:
    body = {k: v for k, v in live.items() if k not in ("MetaData", "domain", "sparse")}
    body["Line"] = lines
    r = requests.post(f"{qbo_api.API_BASE}/v3/company/{cid}/billpayment",
                      headers={"Authorization": f"Bearer {qbo_api._CURRENT_ACCESS or access}",
                               "Content-Type": "application/json", "Accept": "application/json"},
                      params={"minorversion": qbo_api.MINOR_VERSION}, json=body, timeout=120)
    if r.status_code != 200:
        try:
            e = r.json()["Fault"]["Error"][0]
            msg = f"{e.get('Message')}: {e.get('Detail')}"
        except Exception:                                   # noqa: BLE001
            msg = r.text[:400]
        raise RuntimeError(f"QBO refused the update ({r.status_code}): {msg} - nothing was changed")
    return r.json()["BillPayment"]


def _json_plan(pid: str, before: dict, ch: dict, p: dict) -> dict:
    live = p["live"]
    return {"ok": True, "payment_id": pid, "check": live.get("DocNumber") or "",
            "vendor": (live.get("VendorRef") or {}).get("name") or "", "txn_date": live.get("TxnDate"),
            "total": p["total"], "applied_now": p["applied_now"], "applied_after": p["applied_after"],
            "floating_after": round(p["total"] - p["applied_after"], 2), "sync_token": live.get("SyncToken"),
            "before_from": ch["changed_at"], "before_source": ch.get("source") or "change log", "before_lines": len(_links(before)), "already": len(p["already"]),
            "restore": p["restore"], "restore_amt": round(sum(x["amount"] for x in p["restore"]), 2),
            "skipped": [{"type": t, "doc": d, "amount": a, "why": w} for t, d, a, w in p["skipped"]],
            "over": p["applied_after"] > p["total"] + EPS}


def dry_run(payment_id: str) -> dict:
    """What a re-apply would write, read live from QBO. Writes nothing."""
    con = mirror.connect()
    try:
        pid, before, ch = _find_before(con, "", str(payment_id), None)
    finally:
        con.close()
    access, cid = qbo_api.load_credentials()
    return _json_plan(pid, before, ch, plan(access, cid, pid, before))


def commit(payment_id: str, sync_token: str, n: int, amount: float) -> dict:
    """Write the re-apply - only if the plan is still exactly the one the owner confirmed
    (same check version, same number of bills, same amount)."""
    con = mirror.connect()
    try:
        pid, before, ch = _find_before(con, "", str(payment_id), None)
    finally:
        con.close()
    access, cid = qbo_api.load_credentials()
    p = plan(access, cid, pid, before)
    live = p["live"]
    amt = round(sum(x["amount"] for x in p["restore"]), 2)
    if live.get("SyncToken") != str(sync_token) or len(p["restore"]) != int(n) or abs(amt - float(amount)) > EPS:
        return {"ok": False, "error": "The check or its bills changed in QuickBooks since the dry run - run the dry run again."}
    if not p["restore"]:
        return {"ok": False, "error": "Nothing to put back."}
    if p["applied_after"] > p["total"] + EPS:
        return {"ok": False, "error": "Restoring would apply more than the check amount - nothing written."}
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    backup = LOG_DIR / f"{pid}-{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}.json"
    backup.write_text(json.dumps({"live_before": live, "before_copy": before, "restore": p["restore"],
                                  "skipped": p["skipped"]}, indent=1, default=str))
    done = post(access, cid, live, p["lines"])
    lk = _links(done)
    applied = round(sum(a for t, _, a in lk if t != "VendorCredit") - sum(a for t, _, a in lk if t == "VendorCredit"), 2)
    return {"ok": True, "payment_id": pid, "check": live.get("DocNumber") or "", "applied": applied,
            "lines": len(lk), "total": p["total"], "backup": str(backup)}


# ── Clear copy (the bulk fix) ──────────────────────────────────────────────────────────────────────────────

FIXABLE = ("Bill", "VendorCredit")
CREDIT_WINDOW = 1800          # seconds: a credit freed by the strip shows its own change within 30 minutes of it


def _net(links: list) -> float:
    return round(sum(a for t, _, a in links if t != "VendorCredit") - sum(a for t, _, a in links if t == "VendorCredit"), 2)


def _sig(rec: dict) -> tuple:
    return tuple(sorted(_links(rec)))


def _same_check(a: dict, b: dict) -> list:
    """What differs on the check itself between two copies (empty = the same check)."""
    def acct(r):
        return str(((r.get("CheckPayment") or r.get("CreditCardPayment") or {}).get("BankAccountRef")
                    or (r.get("CreditCardPayment") or {}).get("CCAccountRef") or {}).get("value") or "")
    out = []
    if abs(float(a.get("TotalAmt") or 0) - float(b.get("TotalAmt") or 0)) > EPS:
        out.append("amount")
    if str(a.get("TxnDate")) != str(b.get("TxnDate")):
        out.append("date")
    if str((a.get("VendorRef") or {}).get("value")) != str((b.get("VendorRef") or {}).get("value")):
        out.append("vendor")
    if acct(a) != acct(b):
        out.append("bank account")
    if str(a.get("PayType") or "") != str(b.get("PayType") or ""):
        out.append("pay type")
    return out


def _lists_payment(rec: dict, pid: str) -> bool:
    return any("BillPayment" in str(lt.get("TxnType")) and str(lt.get("TxnId")) == pid for lt in rec.get("LinkedTxn") or [])


class _Index:
    """The change log, read once per run (a bulk pass looks at every check)."""

    def __init__(self, con):
        self.con = con
        self.by = {}
        for ent in ("BillPayment",) + FIXABLE:
            d = {}
            for c in mirror.changes(con, entity=ent, limit=1_000_000):
                if c.get("has_before"):
                    d.setdefault(c["rec_id"], []).append(c)
            self.by[ent] = d
        self._before = {}

    def before(self, change_id: int) -> dict:
        if change_id not in self._before:
            self._before[change_id] = mirror.change_before(self.con, change_id) or {}
        return self._before[change_id]

    def rows(self, ent: str, rec_id: str) -> list:
        return self.by.get(ent, {}).get(str(rec_id), [])


def evidence(ix: _Index, pid: str, live: dict) -> dict:
    """The proof that the copy on file is the check exactly as it was, checked against `live` (the check as it
    is now - the mirror's for the page filter, QuickBooks' own for the bulk run). Reads the change log only.
    {"why": [] when proven, "before": the check copy, "change": its log row, "pre": {(type, id): copy facts}}"""
    pid = str(pid)
    out = {"why": [], "before": None, "change": None, "pre": {}}
    why = out["why"]
    rows = sorted(ix.rows("BillPayment", pid), key=lambda c: (c["changed_at"], c["id"]))
    if not rows:
        why.append("no copy of the check from before QuickBooks changed it (the change log started 09/23/2026)")
        return out
    whole = []
    for c in rows:
        b = ix.before(c["id"])
        if b and _links(b) and abs(_net(_links(b)) - round(float(b.get("TotalAmt") or 0), 2)) < EPS:
            whole.append((c, b))
    if not whole:
        why.append("the copy on file is not the whole check (it was not fully applied even then)")
        return out
    if len({_sig(b) for _, b in whole}) > 1:
        why.append("two different copies of the check on file - it was re-applied differently in between")
        return out
    ch, before = whole[-1]
    out["before"], out["change"] = before, ch
    diff = _same_check(before, live)
    if diff:
        why.append("the check itself was changed since (" + ", ".join(diff) + ")")
    want = {(t, i): a for t, i, a in _links(before)}
    for t, i, a in _links(live):
        if (t, i) not in want or a > want[(t, i)] + EPS:
            why.append(f"QuickBooks put part of it on another {'credit' if t == 'VendorCredit' else 'bill'} - take that off by hand first")
            break
    odd = sorted({t for t, _, _ in _links(before) if t not in FIXABLE})
    if odd:
        why.append(f"has a {', '.join(odd)} line - fix by hand")
    strip_at = mirror.parse_iso(ch["changed_at"])
    for t, i, a in _links(before):
        if t not in FIXABLE:
            continue
        cands = [c for c in ix.rows(t, i) if _lists_payment(ix.before(c["id"]), pid)]
        if not cands and t == "VendorCredit":    # a credit may not name the check: its own release, minutes from the strip, by this amount
            for c in ix.rows(t, i):
                try:
                    near = abs((mirror.parse_iso(c["changed_at"]) - strip_at).total_seconds()) <= CREDIT_WINDOW
                except (TypeError, ValueError):
                    near = False
                if near and c.get("balance_before") is not None and c.get("balance_after") is not None \
                        and abs(round(c["balance_after"] - c["balance_before"], 2) - a) < EPS:
                    cands.append(c)
        if not cands:
            why.append(f"no copy of {'credit' if t == 'VendorCredit' else 'bill'} {i} from while this check was on it")
            continue
        c = max(cands, key=lambda c: (c["changed_at"], c["id"]))
        b = ix.before(c["id"])
        out["pre"][(t, i)] = {"balance": round(float(b.get("Balance") or 0), 2),
                              "total": round(float(b.get("TotalAmt") or 0), 2),
                              "vendor": str((b.get("VendorRef") or {}).get("value") or ""), "doc": b.get("DocNumber") or i}
    return out


def _bal_proof(ev: dict, typ: str, tid: str, total_now: float, bal_now: float, need: float, vendor: str) -> str | None:
    """Why this bill/credit does NOT land exactly where it was, or None when it does."""
    pre = ev["pre"].get((typ, tid))
    word = "credit" if typ == "VendorCredit" else "bill"
    if not pre:
        return None                               # already reported by evidence()
    doc = pre["doc"]
    if abs(total_now - pre["total"]) > EPS:
        return f"{word} {doc} total changed {_money(pre['total'])} -> {_money(total_now)} since the check paid it"
    if pre["vendor"] and vendor and pre["vendor"] != vendor:
        return f"{word} {doc} is on a different vendor now"
    after = round(bal_now - need, 2)
    if abs(after - pre["balance"]) > EPS:
        return (f"{word} {doc} would end at {_money(after)} open, it was {_money(pre['balance'])} before - "
                f"{'used' if typ == 'VendorCredit' else 'paid'} or changed since")
    return None


def _proved(ev: dict, p: dict) -> list:
    """Live proof: the plan puts back EXACTLY the copy on file and every bill / credit lands where it was."""
    why = list(ev["why"])
    if why:
        return why
    before, live = ev["before"], p["live"]
    if p["skipped"]:
        why += [f"{d} {_money(a)} cannot go back ({w})" for _, d, a, w in p["skipped"]]
    if p["applied_after"] > p["total"] + EPS:
        why.append("would apply more than the check amount")
    if abs(p["applied_after"] - p["total"]) > EPS:
        why.append(f"would leave {_money(p['total'] - p['applied_after'])} floating")
    after = {}
    for ln in p["lines"]:
        for lt in (ln.get("LinkedTxn") or [])[:1]:
            after[(lt.get("TxnType"), str(lt.get("TxnId")))] = round(float(ln.get("Amount") or 0), 2)
    if after != {(t, i): a for t, i, a in _links(before)}:
        why.append("the check after the fix would not match the copy line for line")
    vendor = str((live.get("VendorRef") or {}).get("value") or "")
    for x in p["restore"]:
        w = _bal_proof(ev, x["type"], x["id"], x["doc_total"], x["balance"], x["need"], vendor)
        if w:
            why.append(w)
    return why


def annotate(rows: list) -> None:
    """The page filter: r["clear_copy"] on every floating check, proven against the MIRROR (QuickBooks as of the
    last 3-minute refresh). The bulk run proves each one again, live, before it writes."""
    todo = [r for r in rows if (r.get("floating") or 0) > 1]
    if not todo:
        return
    con = mirror.connect()
    try:
        ix = _Index(con)
        pays = {str(p["Id"]): p for p in mirror.load("BillPayment", con=con)}
        recs = {}
        for typ in FIXABLE:
            ids = sorted({i for r in todo for c in ix.rows("BillPayment", r["payment_id"])
                          for t, i, _ in _links(ix.before(c["id"])) if t == typ})
            for k in range(0, len(ids), 500):
                chunk = ids[k:k + 500]
                for rec in mirror.load(typ, where_sql=f"id IN ({','.join('?' * len(chunk))})", params=chunk, con=con):
                    recs[(typ, str(rec["Id"]))] = rec
        for r in todo:
            pid = str(r["payment_id"])
            live = pays.get(pid)
            if not live:
                r["clear_copy"] = {"ok": False, "why": ["the check is not in the mirror"]}
                continue
            ev = evidence(ix, pid, live)
            why = list(ev["why"])
            short = credits = 0
            if not why:
                now = {(t, i): a for t, i, a in _links(live)}
                for t, i, a in _links(ev["before"]):
                    need = round(a - now.get((t, i), 0), 2)
                    if need <= EPS:
                        continue
                    rec = recs.get((t, i))
                    if not rec:
                        why.append(f"{'credit' if t == 'VendorCredit' else 'bill'} {i} was deleted")
                        continue
                    w = _bal_proof(ev, t, i, round(float(rec.get("TotalAmt") or 0), 2),
                                   round(float(rec.get("Balance") or 0), 2), need,
                                   str((rec.get("VendorRef") or {}).get("value") or ""))
                    if w:
                        why.append(w)
                    if t == "VendorCredit":
                        credits += 1
                    elif ev["pre"].get((t, i), {}).get("balance", 0) > EPS:
                        short += 1
            r["clear_copy"] = {"ok": not why, "why": why[:6], "short_pays": short, "credits": credits,
                               "copy_from": ev["change"]["changed_at"] if ev["change"] else None}
    finally:
        con.close()


def _check_json(pid: str, ev: dict, p: dict, why: list) -> dict:
    live = p["live"]
    short, credits = [], []
    for x in p["restore"]:
        pre = ev["pre"].get((x["type"], x["id"])) or {}
        if x["type"] == "VendorCredit":
            credits.append({"doc": x["doc"], "id": x["id"], "amount": x["amount"], "balance_after": pre.get("balance")})
        elif (pre.get("balance") or 0) > EPS:
            short.append({"doc": x["doc"], "id": x["id"], "amount": x["amount"], "bill_total": x["doc_total"],
                          "stays_open": pre.get("balance")})
    return {"payment_id": pid, "check": live.get("DocNumber") or "", "vendor": (live.get("VendorRef") or {}).get("name") or "",
            "txn_date": live.get("TxnDate"), "total": p["total"], "applied_now": p["applied_now"],
            "applied_after": p["applied_after"], "sync_token": live.get("SyncToken"),
            "n": len(p["restore"]), "amount": round(sum(x["amount"] for x in p["restore"]), 2),
            "bills": sum(1 for x in p["restore"] if x["type"] == "Bill"), "short_pays": short, "credits": credits,
            "copy_from": (ev.get("change") or {}).get("changed_at"), "ok": not why, "why": why}


def bulk_dry_run(payment_ids: list) -> dict:
    """Every check proven again LIVE from QuickBooks. Writes nothing."""
    con = mirror.connect()
    try:
        ix = _Index(con)
        access, cid = qbo_api.load_credentials()
        out = []
        for pid in dict.fromkeys(str(x) for x in payment_ids if str(x).isdigit()):
            try:
                live = _get_payment(access, cid, pid)          # live check first; the copy is proven against it
                ev = evidence(ix, pid, live)
                p = plan(access, cid, pid, ev["before"] or {"Line": []}, live)
                out.append(_check_json(pid, ev, p, _proved(ev, p)))
            except Exception as e:                        # noqa: BLE001 - one bad check never sinks the run
                out.append({"payment_id": pid, "ok": False, "why": [f"could not read it: {e}"]})
    finally:
        con.close()
    ok = [c for c in out if c["ok"]]
    return {"ok": True, "checks": out, "ready": len(ok), "ready_amt": round(sum(c["amount"] for c in ok), 2)}


def bulk_commit(items: list) -> dict:
    """Write each confirmed check, but only one that STILL proves live and is still exactly what the dry run showed
    (same check version, same number of lines, same amount). Each is backed up first; a check that no longer
    proves is skipped and listed, never forced. One Touch ID for the batch, naming the count and the total."""
    items = [i for i in items if isinstance(i, dict)]
    if not items:
        return {"ok": False, "error": "nothing to write"}
    import presence                                       # noqa: PLC0415  (ledger-local: macOS' own Touch ID dialog)
    total = round(sum(float(i.get("amount") or 0) for i in items), 2)
    ok, why = presence.confirm(f"put {len(items)} stripped checks back on their bills in QuickBooks ({_money(total)})")
    if not ok:
        return {"ok": False, "error": f"not confirmed on this Mac ({why}) - nothing was written"}
    con = mirror.connect()
    try:
        ix = _Index(con)
        access, cid = qbo_api.load_credentials()
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        results = []
        for it in items:
            pid = str(it.get("payment_id") or "")
            res = {"payment_id": pid, "check": it.get("check") or "", "ok": False}
            try:
                live = _get_payment(access, cid, pid)
                ev = evidence(ix, pid, live)
                if not ev["before"]:
                    raise ValueError("; ".join(ev["why"]))
                p = plan(access, cid, pid, ev["before"], live)
                why = _proved(ev, p)
                amt = round(sum(x["amount"] for x in p["restore"]), 2)
                if why:
                    raise ValueError("no longer proves: " + "; ".join(why))
                if p["live"].get("SyncToken") != str(it.get("sync_token")) or len(p["restore"]) != int(it.get("n")) \
                        or abs(amt - float(it.get("amount"))) > EPS:
                    raise ValueError("changed in QuickBooks since the dry run")
                backup = LOG_DIR / f"{pid}-{stamp}-bulk.json"
                backup.write_text(json.dumps({"live_before": p["live"], "before_copy": ev["before"], "change": ev["change"],
                                              "restore": p["restore"], "proof": {f"{t}:{i}": v for (t, i), v in ev["pre"].items()}},
                                             indent=1, default=str))
                done = post(access, cid, p["live"], p["lines"])
                res.update(ok=True, check=p["live"].get("DocNumber") or "", applied=_net(_links(done)),
                           total=p["total"], lines=len(_links(done)), backup=str(backup))
            except Exception as e:                        # noqa: BLE001
                res["error"] = str(e)
            results.append(res)
    finally:
        con.close()
    return {"ok": True, "results": results, "written": sum(1 for r in results if r["ok"])}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("check", nargs="?", help="check number (the payment's DocNumber), e.g. 25745")
    ap.add_argument("--clear", action="store_true", help="every clear-copy check (the bulk fix), proven live")
    ap.add_argument("--payment-id", help="QBO BillPayment Id, when two payments share the check number")
    ap.add_argument("--change-id", type=int, help="use this change-log entry's before copy (default: the fullest one)")
    ap.add_argument("--list", action="store_true", help="print every bill that goes back on, not just the totals")
    ap.add_argument("--commit", action="store_true", help="write to QBO (default is a dry run)")
    a = ap.parse_args()
    if a.clear:
        return _main_clear(a.commit)
    if not a.check:
        ap.error("a check number, or --clear")

    con = mirror.connect()
    try:
        pid, before, ch = _find_before(con, a.check, a.payment_id, a.change_id)
    except LookupError as e:
        sys.exit(str(e))
    finally:
        con.close()
    access, cid = qbo_api.load_credentials()
    p = plan(access, cid, pid, before)
    live = p["live"]

    print(f"check {a.check}  {(live.get('VendorRef') or {}).get('name')}  {_fmt(live.get('TxnDate'))}  "
          f"total {_money(p['total'])}")
    src = ch.get("source") or "the change log"
    print(f"  what it paid, from {src} ({_fmt(ch['changed_at'])}): {len(_links(before))} lines")
    print(f"  applied now {_money(p['applied_now'])}  ->  after {_money(p['applied_after'])}"
          f"  (floating after: {_money(p['total'] - p['applied_after'])})")
    print(f"  already on the check: {len(p['already'])}   to put back: {len(p['restore'])} "
          f"({_money(sum(x['amount'] for x in p['restore']))})   skipped: {len(p['skipped'])}")
    if a.list:
        for x in sorted(p["restore"], key=lambda x: x["date"] or ""):
            print(f"    + {x['doc']:<14} {_fmt(x['date'])}  {_money(x['amount']):>12}  {x['memo']}")
    for t, doc, amt, why in p["skipped"]:
        print(f"    SKIP {t} {doc}  {_money(amt)}  - {why}")

    if p["applied_after"] > p["total"] + EPS:
        sys.exit("STOP: restoring would apply more than the check amount - nothing written")
    if not p["restore"]:
        print("nothing to put back")
        return 0
    if not a.commit:
        print("\ndry run - nothing written. Add --commit to write it to QBO.")
        return 0

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = LOG_DIR / f"{pid}-{stamp}.json"
    fresh = _get_payment(access, cid, pid)                  # fresh SyncToken, and nobody changed it meanwhile
    if fresh.get("SyncToken") != live.get("SyncToken"):
        sys.exit("STOP: the check changed in QBO while this ran - run it again")
    backup.write_text(json.dumps({"live_before": fresh, "before_copy": before, "restore": p["restore"],
                                  "skipped": p["skipped"]}, indent=1, default=str))
    try:
        done = post(access, cid, fresh, p["lines"])
    except RuntimeError as e:
        sys.exit(str(e))
    applied = sum(x[2] for x in _links(done) if x[0] != "VendorCredit") - \
        sum(x[2] for x in _links(done) if x[0] == "VendorCredit")
    print(f"\nwritten. QBO now shows {_money(applied)} applied on check {a.check} "
          f"({len(_links(done))} lines). Backup: {backup}")
    print("Refresh the mirror (ledger gear > sync, or ledger/refresh_mirror.py) so the ledger clears it.")
    return 0


def _main_clear(commit: bool) -> int:
    import check_drift                                     # local: the floating checks, read off the mirror
    rows = [r for r in check_drift.audit()["rows"] if r["floating"] > 1]
    annotate(rows)
    pids = [r["payment_id"] for r in rows if (r.get("clear_copy") or {}).get("ok")]
    print(f"{len(rows)} floating checks, {len(pids)} clear copies on the mirror - proving each live ...")
    d = bulk_dry_run(pids)
    for c in d["checks"]:
        tag = "READY" if c["ok"] else "SKIP "
        print(f"  {tag} {c.get('check', ''):<10} {_fmt(c.get('txn_date')):<10} {str(c.get('vendor', ''))[:28]:<28} "
              f"{_money(c.get('total') or 0):>12}  {c.get('n', 0)} lines"
              + (f"  short pays {len(c['short_pays'])}" if c.get("short_pays") else "")
              + (f"  credits {len(c['credits'])}" if c.get("credits") else "")
              + ("" if c["ok"] else "  - " + "; ".join(c["why"])))
    print(f"ready: {d['ready']} checks, {_money(d['ready_amt'])}")
    if not commit:
        print("\ndry run - nothing written. Add --commit to write every READY check to QBO.")
        return 0
    res = bulk_commit([c for c in d["checks"] if c["ok"]])
    for r in res["results"]:
        print(f"  {'written' if r['ok'] else 'NOT written'} {r['check']}: "
              + (f"{_money(r['applied'])} applied, backup {r['backup']}" if r["ok"] else r.get("error", "")))
    print(f"{res['written']} written. Refresh the mirror (ledger/refresh_mirror.py) so the ledger clears them.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
