"""
bill_marks.py - the bill-mark overlay: human edits the ledger dashboard writes and the
AP sync mirrors back to the workbook.

WHY THIS EXISTS
`ap_bill_line` is a MIRROR (full-replaced on every load), so a value typed on the website
into that table would be wiped on the next sync. Instead, per-bill web edits live here, in
a tiny overlay table keyed by the QBO **bill id** - which is exactly the workbook's hidden
`_Key`, so a mark (a) survives every `ap_bill_line` reload and (b) joins cleanly back to the
Bill Tracker workbook. Today the only mark is the **Lien tag** (Notice Sent / Lien Filed /
Released), the same manual escalation the workbook's Lien cell has always held.

FLOW (the owner chose "overlay + mirror on sync", 2026-08-17)
    website click ─▶ set_lien_mark()  ─▶  bill_mark table (this module)      ← instant, on the site
    next `sync-ap` ─▶ read_lien_marks() ─▶ excel_bill_sync writes the Lien cell ← reflects in the workbook
The dashboard NEVER edits the live OneDrive .xlsx directly (file locks / corruption); the AP
producer owns the workbook and applies the marks on its next run.

The overlay lives inside the ledger DB (the dashboard's existing write surface). Every reader
is ABSENT-SAFE: no ledger DB / no table → {}, so a tool without a ledger behaves exactly as
before. `shared/` is the only importable common code, so both the dashboard (writer) and
excel_bill_sync (reader) use this one module - no tool imports another tool.
"""
from __future__ import annotations

import sqlite3
from pathlib import Path

try:                                   # sibling shared module (package import)
    from . import paths
except ImportError:                    # run with the repo root on sys.path
    import paths                       # type: ignore

# Same resolution the dashboard uses for the ledger DB (env → machine.env → default).
LEDGER_DB: Path = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3")

# Settable lien tags - MATCH the workbook's manual escalation set exactly (excel_bill_sync
# uses these same literals). "" clears the mark (revert to the computed countdown).
LIEN_STATES = ("Notice Sent", "Lien Filed", "✓ Released")


def bill_id_from_link(qbo_link: str | None) -> str | None:
    """Pull the QBO bill id out of a bill deep link (…txnId=12345…). None if absent."""
    if not qbo_link:
        return None
    import re
    m = re.search(r"txnId(?:=|%3D)(\d+)", str(qbo_link), re.I)   # bare or URL-encoded '='
    return m.group(1) if m else None


def _ensure(con: sqlite3.Connection) -> None:
    con.execute("CREATE TABLE IF NOT EXISTS bill_mark ("
                "bill_id TEXT PRIMARY KEY, lien TEXT, updated_at TEXT NOT NULL)")


def read_lien_marks() -> dict:
    """{bill_id -> lien tag} for EVERY marked bill, including '' (explicitly cleared) so a
    clear can override a preserved workbook edit. {} if there is no ledger DB / table yet."""
    if not LEDGER_DB.exists():
        return {}
    try:
        con = sqlite3.connect(f"file:{LEDGER_DB}?mode=ro", uri=True)
    except sqlite3.OperationalError:
        return {}
    try:
        rows = con.execute("SELECT bill_id, lien FROM bill_mark").fetchall()
        return {str(b): (l or "") for b, l in rows}
    except sqlite3.OperationalError:      # table not created yet
        return {}
    finally:
        con.close()


def set_lien_mark(bill_id: str, lien: str | None, now: str) -> None:
    """Upsert a bill's lien tag. lien in LIEN_STATES sets it; '' / None clears it (the row is
    kept as '' so the AP sync knows to clear the workbook cell too). Opens the ledger writable
    - the dashboard's one write surface. Raises if the ledger DB is missing."""
    lien = (lien or "").strip()
    con = sqlite3.connect(str(LEDGER_DB))
    try:
        _ensure(con)
        con.execute(
            "INSERT INTO bill_mark (bill_id, lien, updated_at) VALUES (?,?,?) "
            "ON CONFLICT(bill_id) DO UPDATE SET lien=excluded.lien, updated_at=excluded.updated_at",
            (str(bill_id), lien, now))
        con.commit()
    finally:
        con.close()


def resolve_lien(bill_id: str | None, marks: dict, preserved: str | None, computed: str | None) -> str | None:
    """The final Lien value for a workbook row (used by excel_bill_sync on the next sync):
    a ledger mark wins (including '' = cleared → fall through to the computed countdown); with
    no mark, keep the value the clerk preserved in the workbook; else the computed bucket."""
    if bill_id is not None and bill_id in marks:
        tag = marks[bill_id]
        if tag in LIEN_STATES:
            return tag
        return computed                   # mark was cleared → recompute the countdown
    return preserved if preserved else computed


# ── Pay-run marks ────────────────────────────────────────────────────────────
# A SEPARATE overlay from the lien tag: which bills the owner has selected to pay in
# the current check run, and (optionally) a partial amount that overrides the bill's
# full open balance. This is a PLANNING WORKSHEET only - it records intent, never a
# payment. QBO stays the source of truth: the owner records the real payment in QBO,
# and the next `sync-ap` pulls the true pay status back (the bill then drops off the
# open list). So, unlike the lien tag, pay marks are NOT mirrored to the workbook and
# are NOT read by excel_bill_sync - the dashboard is their only reader/writer. Keyed by
# the same QBO bill id, absent-safe, in the same ledger DB (the dashboard's write surface).
def _ensure_pay(con: sqlite3.Connection) -> None:
    con.execute("CREATE TABLE IF NOT EXISTS pay_mark ("
                "bill_id TEXT PRIMARY KEY, amount REAL, updated_at TEXT NOT NULL)")


def read_pay_marks() -> dict:
    """{bill_id -> {'amount': float|None}} for every bill in the current pay run. A None
    amount means 'pay the full open balance' (the caller resolves it). {} if there is no
    ledger DB / table yet - absent-safe, so a tool without a ledger behaves as before."""
    if not LEDGER_DB.exists():
        return {}
    try:
        con = sqlite3.connect(f"file:{LEDGER_DB}?mode=ro", uri=True)
    except sqlite3.OperationalError:
        return {}
    try:
        rows = con.execute("SELECT bill_id, amount FROM pay_mark").fetchall()
        return {str(b): {"amount": a} for b, a in rows}
    except sqlite3.OperationalError:      # table not created yet
        return {}
    finally:
        con.close()


def set_pay_marks(items: list, now: str) -> int:
    """Persist a pay run in ONE transaction. Each item is {bill_id, amount, selected}: a
    selected bill is upserted (amount None → pay the full open balance), an unselected one
    is deleted. Returns the number of bills kept in the run. Opens the ledger writable."""
    con = sqlite3.connect(str(LEDGER_DB))
    try:
        _ensure_pay(con)
        kept = 0
        for it in items or []:
            bid = str(it.get("bill_id") or "").strip()
            if not bid:
                continue
            if it.get("selected"):
                amt = it.get("amount")
                amt = float(amt) if amt not in (None, "") else None
                con.execute(
                    "INSERT INTO pay_mark (bill_id, amount, updated_at) VALUES (?,?,?) "
                    "ON CONFLICT(bill_id) DO UPDATE SET amount=excluded.amount, updated_at=excluded.updated_at",
                    (bid, amt, now))
                kept += 1
            else:
                con.execute("DELETE FROM pay_mark WHERE bill_id=?", (bid,))
        con.commit()
        return kept
    finally:
        con.close()


def clear_pay_marks() -> int:
    """Empty the whole pay run (after the check run is done). Returns rows removed; 0 if the
    ledger DB / table is absent."""
    if not LEDGER_DB.exists():
        return 0
    con = sqlite3.connect(str(LEDGER_DB))
    try:
        _ensure_pay(con)
        n = con.execute("SELECT COUNT(*) FROM pay_mark").fetchone()[0]
        con.execute("DELETE FROM pay_mark")
        con.commit()
        return n
    finally:
        con.close()
