"""
shared/qbo_mirror.py - THE raw QBO mirror: every transaction and list entity,
as QBO returned it, in one SQLite file outside the repo. Every tool reads this
instead of pulling QBO itself (the owner 2026-09-17: "a full data pull of all
transactions in QBO ... back it up into a database").

  seed      full pull, once (296k records / ~300 pages / ~20 min measured)
  refresh   the change feed (/cdc) since the last stamp: adds, edits AND
            deletes, seconds on a normal day. QBO's feed reaches back 30 days;
            older than that the refresh falls back to LastUpdatedTime queries
            plus an id sweep (the query path cannot report deletions).
  reconcile COUNT(*) per entity vs the mirror; any drift -> id sweep.

Rows are never dropped: a deleted transaction is flagged `deleted=1` with the
date, so the mirror is also the history QBO does not keep. QBO stays the source
of truth; the mirror is a stamped copy and nothing writes to QBO from here.

Table per entity `qbo_<entity>`: id, sync_token, created, last_updated,
txn_date, doc_number, total, name, deleted, deleted_at, seen_at, json (zlib).
`load(entity)` hands a tool the parsed records - the replacement for
`query_all(access, cid, entity, where)`; filter in Python or via `where_sql`.

DB: ~/Library/Application Support/Proficient/qbo_mirror.sqlite3 (override
`ACB_QBO_MIRROR_DB`). Separate from ledger.sqlite3 on purpose: raw vs derived,
so the ledger can always be rebuilt from the mirror. Never print the realm.
"""
from __future__ import annotations

import datetime as dt
import json
import sqlite3
import zlib
from pathlib import Path
from typing import Callable, Dict, Iterable, Iterator, List, Optional, Tuple

from shared import paths
from shared import qbo_api

# Every entity QBO's change feed (/cdc) supports and the business uses.
# list=True: a name list whose query hides inactive rows unless asked.
ENTITIES: Dict[str, dict] = {
    "Account":       {"list": True},
    "Bill":          {"list": False},
    "BillPayment":   {"list": False},
    "Class":         {"list": True},
    "CreditMemo":    {"list": False},
    "Customer":      {"list": True},
    "Deposit":       {"list": False},
    "Estimate":      {"list": False},
    "Invoice":       {"list": False},
    "Item":          {"list": True},
    "JournalEntry":  {"list": False},
    "Payment":       {"list": False},
    "PaymentMethod": {"list": True},
    "Purchase":      {"list": False},
    "PurchaseOrder": {"list": False},
    "RefundReceipt": {"list": False},
    "SalesReceipt":  {"list": False},
    "Term":          {"list": True},
    "Transfer":      {"list": False},
    "Vendor":        {"list": True},
    "VendorCredit":  {"list": False},
}
PAGE = 1000            # QBO's max page
CDC_CAP = 1000         # /cdc returns at most this many per entity per call
CDC_MAX_DAYS = 29      # the feed reaches back 30 days; stay a day inside
OVERLAP = dt.timedelta(minutes=10)   # re-read a little; upserts are idempotent
FETCH_CHUNK = 100      # ids per `WHERE Id IN (...)`

DEFAULT_DB = Path.home() / "Library" / "Application Support" / "Proficient" / "qbo_mirror.sqlite3"

Progress = Optional[Callable[[str], None]]


# ───────────────────────── storage ─────────────────────────

def db_path() -> Path:
    return Path(paths.get_path("ACB_QBO_MIRROR_DB", DEFAULT_DB))


def table(entity: str) -> str:
    if entity not in ENTITIES:
        raise KeyError(f"not a mirrored entity: {entity}")
    return "qbo_" + entity.lower()


_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS {t} (
    id           TEXT PRIMARY KEY,
    sync_token   INTEGER,
    created      TEXT,
    last_updated TEXT,
    txn_date     TEXT,
    doc_number   TEXT,
    total        REAL,
    name         TEXT,
    deleted      INTEGER NOT NULL DEFAULT 0,
    deleted_at   TEXT,
    seen_at      TEXT,
    json         BLOB
);
CREATE INDEX IF NOT EXISTS {t}_txn_date ON {t}(txn_date);
CREATE INDEX IF NOT EXISTS {t}_updated  ON {t}(last_updated);
CREATE INDEX IF NOT EXISTS {t}_doc      ON {t}(doc_number);
"""
_META_SQL = """
CREATE TABLE IF NOT EXISTS mirror_meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS mirror_run (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    started  TEXT NOT NULL,
    finished TEXT,
    mode     TEXT NOT NULL,
    upserted INTEGER NOT NULL DEFAULT 0,
    deleted  INTEGER NOT NULL DEFAULT 0,
    note     TEXT
);
"""


def connect(path: Optional[Path | str] = None) -> sqlite3.Connection:
    """Open (and create) the mirror. ':memory:' for tests."""
    p = str(path) if path is not None else str(db_path())
    if p != ":memory:":
        Path(p).parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(p)
    con.row_factory = sqlite3.Row
    if p != ":memory:":
        con.execute("PRAGMA journal_mode=WAL")
    con.executescript(_META_SQL + "".join(_TABLE_SQL.format(t=table(e)) for e in ENTITIES))
    return con


# ── at rest: every record is AES-256-GCM encrypted (the owner 2026-09-17: "how do
# we keep the data safe so nobody can just steal the file"). The key lives in the
# ONE Keychain library (`MIRROR_KEY`, `shared/setup_qbo.py --rotate MIRROR_KEY`),
# so a copied file is unreadable off this Mac. Blob = b"ENC1" + 12-byte nonce +
# ciphertext(zlib(json)). A legacy plain zlib blob still reads (it starts with
# 0x78); `encrypt_existing` rewrites those once.
_MAGIC = b"ENC1"
_KEY = [None]


def _key() -> bytes:
    if _KEY[0] is None:
        import base64
        from shared import qbo_vault
        try:
            raw = base64.b64decode(qbo_vault.get("MIRROR_KEY"))
        except qbo_vault.SecretsError as e:
            raise RuntimeError("the mirror's MIRROR_KEY is not in the Keychain library - "
                               "run: python3 shared/setup_qbo.py --rotate MIRROR_KEY") from e
        if len(raw) != 32:
            raise RuntimeError("MIRROR_KEY must be 32 bytes (base64 of os.urandom(32))")
        _KEY[0] = raw
    return _KEY[0]


def _pack(rec: dict) -> bytes:
    import os
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    plain = zlib.compress(json.dumps(rec, separators=(",", ":")).encode("utf-8"))
    nonce = os.urandom(12)
    return _MAGIC + nonce + AESGCM(_key()).encrypt(nonce, plain, None)


def _unpack(blob: bytes) -> dict:
    blob = bytes(blob)
    if blob[:4] == _MAGIC:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        plain = AESGCM(_key()).decrypt(blob[4:16], blob[16:], None)
    else:
        plain = blob                       # legacy: plain zlib
    return json.loads(zlib.decompress(plain).decode("utf-8"))


def encrypt_existing(con, progress: Progress = print) -> int:
    """Rewrite every legacy (plain zlib) row encrypted. Idempotent."""
    say = progress or (lambda s: None)
    total = 0
    for e in ENTITIES:
        t = table(e)
        rows = con.execute(f"SELECT id, json FROM {t} WHERE substr(json,1,4) != ?", (_MAGIC,)).fetchall()
        if not rows:
            continue
        con.executemany(f"UPDATE {t} SET json=? WHERE id=?",
                        [(_pack(_unpack(r[1])), r[0]) for r in rows])
        con.commit()
        total += len(rows)
        say(f"  {e}: {len(rows):,} rows encrypted")
    return total


def encryption_status(con) -> Tuple[int, int]:
    """(encrypted rows, plain rows) across every table."""
    enc = plain = 0
    for e in ENTITIES:
        r = con.execute(f"SELECT SUM(substr(json,1,4)=?), SUM(substr(json,1,4)!=?) FROM {table(e)}",
                        (_MAGIC, _MAGIC)).fetchone()
        enc += r[0] or 0
        plain += r[1] or 0
    return enc, plain


def _meta_get(con, key: str) -> Optional[str]:
    r = con.execute("SELECT value FROM mirror_meta WHERE key=?", (key,)).fetchone()
    return r[0] if r else None


def _meta_set(con, key: str, value: str) -> None:
    con.execute("INSERT OR REPLACE INTO mirror_meta(key, value) VALUES (?, ?)", (key, value))


def now_utc() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0)


def iso_z(t: dt.datetime) -> str:
    return t.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_iso(s: str) -> dt.datetime:
    s = s.strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    return dt.datetime.fromisoformat(s)


def _summary(rec: dict) -> tuple:
    md = rec.get("MetaData") or {}
    # the name QBO sorts a list by: DisplayName (vendors/customers), else the
    # fully qualified name (accounts/items nest under parents), else Name (classes)
    name = rec.get("DisplayName") or rec.get("FullyQualifiedName") or rec.get("Name")
    total = rec.get("TotalAmt")
    if total is None:
        total = rec.get("Amount")
    st = rec.get("SyncToken")
    return (
        int(st) if str(st or "").isdigit() else None,
        md.get("CreateTime"), md.get("LastUpdatedTime"),
        rec.get("TxnDate"), rec.get("DocNumber"),
        float(total) if total not in (None, "") else None,
        name,
    )


def upsert_many(con, entity: str, recs: Iterable[dict], seen_at: str) -> int:
    """Write records as QBO returned them; a record that comes back is alive
    again (deleted=0). Last writer wins - the feed always hands us the latest."""
    t = table(entity)
    rows = []
    for rec in recs:
        rid = rec.get("Id")
        if not rid:
            continue
        rows.append((str(rid), *_summary(rec), seen_at, _pack(rec)))
    if not rows:
        return 0
    con.executemany(
        f"INSERT OR REPLACE INTO {t}(id, sync_token, created, last_updated, txn_date, "
        f"doc_number, total, name, deleted, deleted_at, seen_at, json) "
        f"VALUES (?,?,?,?,?,?,?,?,0,NULL,?,?)", rows)
    return len(rows)


def mark_deleted(con, entity: str, ids: Iterable[str], when: str) -> int:
    """Flag rows QBO no longer has. Rows are kept - that is the history."""
    ids = [str(i) for i in ids if i]
    if not ids:
        return 0
    t = table(entity)
    n = 0
    for i in range(0, len(ids), 500):
        chunk = ids[i:i + 500]
        cur = con.execute(
            f"UPDATE {t} SET deleted=1, deleted_at=? WHERE deleted=0 AND id IN "
            f"({','.join('?' * len(chunk))})", (when, *chunk))
        n += cur.rowcount
    return n


def _log_run(con, started: str, mode: str, upserted: int, deleted: int, note: str = "") -> None:
    con.execute("INSERT INTO mirror_run(started, finished, mode, upserted, deleted, note) "
                "VALUES (?,?,?,?,?,?)", (started, iso_z(now_utc()), mode, upserted, deleted, note))


# ───────────────────────── QBO reads ─────────────────────────

def _list_where(entity: str) -> str:
    return "Active IN (true, false)" if ENTITIES[entity]["list"] else ""


def _and(*parts: str) -> str:
    return " AND ".join(p for p in parts if p)


def _query_pages(access: str, cid: str, entity: str, where: str = "",
                 select: str = "*") -> Iterator[List[dict]]:
    """Page through SELECT <select> FROM <entity> [WHERE …], QBO's max page,
    the repo's retrying GET (token refresh, 5xx/429 backoff)."""
    start = 1
    while True:
        q = f"SELECT {select} FROM {entity}"
        if where:
            q += f" WHERE {where}"
        q += f" STARTPOSITION {start} MAXRESULTS {PAGE}"
        batch = (qbo_api._api_get(f"/v3/company/{cid}/query", access, {"query": q})
                 .get("QueryResponse", {}).get(entity, []))
        if not batch:
            return
        yield batch
        if len(batch) < PAGE:
            return
        start += PAGE


def count(access: str, cid: str, entity: str) -> int:
    q = f"SELECT COUNT(*) FROM {entity}"
    w = _list_where(entity)
    if w:
        q += f" WHERE {w}"
    return int(qbo_api._api_get(f"/v3/company/{cid}/query", access, {"query": q})
               .get("QueryResponse", {}).get("totalCount") or 0)


def _cdc_call(access: str, cid: str, entities: List[str], since: str) -> dict:
    return qbo_api._api_get(f"/v3/company/{cid}/cdc", access,
                            {"entities": ",".join(entities), "changedSince": since})


def parse_cdc(data: dict) -> Dict[str, dict]:
    """/cdc response -> {entity: {alive: [rec], deleted: [(id, when)], capped: bool,
    newest: iso|None}}. A deleted record arrives with status 'Deleted' and just
    Id + MetaData. `capped` = this entity hit the per-call cap, ask again from
    `newest`."""
    out: Dict[str, dict] = {}
    for block in data.get("CDCResponse", []) or []:
        for qr in block.get("QueryResponse", []) or []:
            for entity in ENTITIES:
                items = qr.get(entity)
                if items is None:
                    continue
                slot = out.setdefault(entity, {"alive": [], "deleted": [], "capped": False, "newest": None})
                for rec in items:
                    upd = (rec.get("MetaData") or {}).get("LastUpdatedTime")
                    if upd and (slot["newest"] is None or upd > slot["newest"]):
                        slot["newest"] = upd
                    if str(rec.get("status", "")).lower() == "deleted":
                        slot["deleted"].append((str(rec.get("Id")), upd))
                    else:
                        slot["alive"].append(rec)
                if len(items) >= CDC_CAP or int(qr.get("maxResults") or 0) >= CDC_CAP:
                    slot["capped"] = True
    return out


# ───────────────────────── the three verbs ─────────────────────────

def seed(con, access: str, cid: str, entities: Optional[List[str]] = None,
         progress: Progress = print) -> dict:
    """Full pull. Also a full reconcile: a row not seen by this pull is flagged
    deleted. Sets the refresh stamp to the pull's START so nothing edited
    mid-pull is skipped by the next refresh."""
    say = progress or (lambda s: None)
    ents = entities or list(ENTITIES)
    started = iso_z(now_utc())
    tot_up = tot_del = 0
    per: Dict[str, int] = {}
    for e in ents:
        expected = count(access, cid, e)
        pages = max(1, -(-expected // PAGE))
        say(f"  {e}: {expected:,} in QBO ({pages} page{'s' if pages != 1 else ''})")
        n = 0
        t0 = dt.datetime.now()
        for i, batch in enumerate(_query_pages(access, cid, e, _list_where(e)), 1):
            n += upsert_many(con, e, batch, started)
            con.commit()
            el = (dt.datetime.now() - t0).total_seconds()
            say(f"    page {i}/{pages}  {n:,} rows  {el:.0f}s")
        # anything this full pull did not see is gone from QBO
        cur = con.execute(f"UPDATE {table(e)} SET deleted=1, deleted_at=? "
                          f"WHERE deleted=0 AND (seen_at IS NULL OR seen_at < ?)", (started, started))
        tot_del += cur.rowcount
        tot_up += n
        per[e] = n
        con.commit()
    if not entities:
        _meta_set(con, "last_refresh", started)
        _meta_set(con, "seeded_at", started)
    _log_run(con, started, "seed", tot_up, tot_del, json.dumps(per))
    con.commit()
    return {"mode": "seed", "started": started, "upserted": tot_up, "deleted": tot_del, "per": per}


def refresh(con, access: str, cid: str, progress: Progress = print) -> dict:
    """Bring the mirror to now. No stamp -> seed. Stamp inside the feed's reach
    -> /cdc. Older -> LastUpdatedTime queries + id sweep (deletes)."""
    say = progress or (lambda s: None)
    last = _meta_get(con, "last_refresh")
    if not last:
        say("no refresh stamp - seeding (full pull)")
        return seed(con, access, cid, progress=progress)
    last_dt = parse_iso(last)
    started = now_utc()
    age = started - last_dt
    if age > dt.timedelta(days=CDC_MAX_DAYS):
        say(f"stamp is {age.days} days old - past the feed's reach; using LastUpdatedTime + id sweep")
        res = _refresh_by_updated(con, access, cid, last_dt - OVERLAP, say)
        mode = "updated+sweep"
    else:
        since = iso_z(last_dt - OVERLAP)
        say(f"change feed since {since}")
        res = _refresh_cdc(con, access, cid, since, say)
        mode = "cdc"
    _meta_set(con, "last_refresh", iso_z(started))
    _log_run(con, iso_z(started), mode, res["upserted"], res["deleted"], json.dumps(res.get("per", {})))
    con.commit()
    return {"mode": mode, "started": iso_z(started), **res}


def _refresh_cdc(con, access: str, cid: str, since: str, say) -> dict:
    seen_at = iso_z(now_utc())
    up = de = 0
    per: Dict[str, dict] = {}
    parsed = parse_cdc(_cdc_call(access, cid, list(ENTITIES), since))
    pending = list(parsed.items())
    while pending:
        entity, slot = pending.pop(0)
        n_up = upsert_many(con, entity, slot["alive"], seen_at)
        n_de = mark_deleted(con, entity, [i for i, _ in slot["deleted"]], seen_at)
        con.commit()
        p = per.setdefault(entity, {"upserted": 0, "deleted": 0})
        p["upserted"] += n_up
        p["deleted"] += n_de
        up += n_up
        de += n_de
        if slot["capped"] and slot["newest"]:
            # more than the cap changed: ask again for this entity from its newest stamp
            say(f"  {entity}: {CDC_CAP}+ changes, continuing from {slot['newest']}")
            more = parse_cdc(_cdc_call(access, cid, [entity], slot["newest"]))
            if entity in more and (more[entity]["alive"] or more[entity]["deleted"]):
                pending.append((entity, more[entity]))
    for entity, p in per.items():
        if p["upserted"] or p["deleted"]:
            say(f"  {entity}: {p['upserted']} changed, {p['deleted']} deleted")
    if not per:
        say("  nothing changed")
    return {"upserted": up, "deleted": de, "per": per}


def _refresh_by_updated(con, access: str, cid: str, since_dt: dt.datetime, say) -> dict:
    seen_at = iso_z(now_utc())
    since = iso_z(since_dt)
    up = 0
    per: Dict[str, dict] = {}
    for e in ENTITIES:
        n = 0
        where = _and(f"MetaData.LastUpdatedTime > '{since}'", _list_where(e))
        for batch in _query_pages(access, cid, e, where):
            n += upsert_many(con, e, batch, seen_at)
            con.commit()
        if n:
            say(f"  {e}: {n} updated since {since}")
        per[e] = {"upserted": n, "deleted": 0}
        up += n
    sw = sweep_deleted(con, access, cid, list(ENTITIES), say)
    for e, d in sw["per"].items():
        per.setdefault(e, {"upserted": 0, "deleted": 0})
        per[e]["deleted"] += d["deleted"]
        per[e]["upserted"] += d["fetched"]
    return {"upserted": up + sw["fetched"], "deleted": sw["deleted"], "per": per}


def sweep_deleted(con, access: str, cid: str, entities: List[str], say=None) -> dict:
    """Compare ids: in the mirror but not in QBO -> deleted; in QBO but not in
    the mirror -> fetched. `SELECT Id` pages are tiny, so this is cheap."""
    say = say or (lambda s: None)
    seen_at = iso_z(now_utc())
    tot_del = tot_fetch = 0
    per: Dict[str, dict] = {}
    for e in entities:
        qbo_ids = set()
        for batch in _query_pages(access, cid, e, _list_where(e), select="Id"):
            qbo_ids.update(str(r.get("Id")) for r in batch if r.get("Id"))
        mine = {r[0] for r in con.execute(f"SELECT id FROM {table(e)} WHERE deleted=0")}
        gone = mine - qbo_ids
        missing = sorted(qbo_ids - mine, key=lambda s: int(s) if s.isdigit() else 0)
        n_del = mark_deleted(con, e, gone, seen_at)
        n_fetch = 0
        for i in range(0, len(missing), FETCH_CHUNK):
            chunk = missing[i:i + FETCH_CHUNK]
            where = _and("Id IN (" + ",".join(f"'{x}'" for x in chunk) + ")", _list_where(e))
            for batch in _query_pages(access, cid, e, where):
                n_fetch += upsert_many(con, e, batch, seen_at)
        con.commit()
        if n_del or n_fetch:
            say(f"  {e}: {n_del} deleted, {n_fetch} fetched")
        per[e] = {"deleted": n_del, "fetched": n_fetch}
        tot_del += n_del
        tot_fetch += n_fetch
    return {"deleted": tot_del, "fetched": tot_fetch, "per": per}


def reconcile(con, access: str, cid: str, progress: Progress = print) -> dict:
    """COUNT(*) per entity vs the mirror's alive rows; drift -> id sweep."""
    say = progress or (lambda s: None)
    started = iso_z(now_utc())
    drift = []
    for e in ENTITIES:
        q = count(access, cid, e)
        m = con.execute(f"SELECT COUNT(*) FROM {table(e)} WHERE deleted=0").fetchone()[0]
        flag = "" if q == m else "   <- drift"
        say(f"  {e:14s} QBO {q:>7,}  mirror {m:>7,}{flag}")
        if q != m:
            drift.append(e)
    res = sweep_deleted(con, access, cid, drift, say) if drift else {"deleted": 0, "fetched": 0, "per": {}}
    _log_run(con, started, "reconcile", res["fetched"], res["deleted"], json.dumps(drift))
    con.commit()
    return {"mode": "reconcile", "drift": drift, **res}


# ───────────────────────── the read side (what tools call) ─────────────────────────

def load(entity: str, where_sql: str = "", params: Tuple = (), include_deleted: bool = False,
         con: Optional[sqlite3.Connection] = None) -> List[dict]:
    """Parsed records for a tool - the stand-in for `query_all(...)`. `where_sql`
    is SQL over the summary columns (txn_date, doc_number, total, name,
    last_updated), e.g. "txn_date >= ?"."""
    own = con is None
    con = con or connect()
    try:
        t = table(entity)
        parts = [] if include_deleted else ["deleted=0"]
        if where_sql:
            parts.append(f"({where_sql})")
        sql = f"SELECT json FROM {t}"
        if parts:
            sql += " WHERE " + " AND ".join(parts)
        # QBO's own result order, so a tool that takes the FIRST candidate
        # (draw matching, PO index) picks the same row it always did (bill
        # tracker diff 2026-09-17): transactions newest-edited first, name lists
        # alphabetical (case-insensitive).
        if ENTITIES[entity]["list"]:
            sql += " ORDER BY lower(name), CAST(id AS INTEGER)"
        else:
            sql += " ORDER BY last_updated DESC, CAST(id AS INTEGER) DESC"
        return [_unpack(r[0]) for r in con.execute(sql, params)]
    finally:
        if own:
            con.close()


def get(entity: str, rec_id: str, con: Optional[sqlite3.Connection] = None) -> Optional[dict]:
    own = con is None
    con = con or connect()
    try:
        r = con.execute(f"SELECT json FROM {table(entity)} WHERE id=?", (str(rec_id),)).fetchone()
        return _unpack(r[0]) if r else None
    finally:
        if own:
            con.close()


def status(con: Optional[sqlite3.Connection] = None) -> dict:
    own = con is None
    con = con or connect()
    try:
        per = {}
        for e in ENTITIES:
            r = con.execute(
                f"SELECT SUM(deleted=0), SUM(deleted=1), MAX(last_updated) FROM {table(e)}").fetchone()
            per[e] = {"alive": r[0] or 0, "deleted": r[1] or 0, "newest": r[2]}
        p = db_path()
        return {"db": str(p), "size_mb": round(p.stat().st_size / 1e6, 1) if p.exists() else 0.0,
                "last_refresh": _meta_get(con, "last_refresh"),
                "seeded_at": _meta_get(con, "seeded_at"), "entities": per}
    finally:
        if own:
            con.close()


# ───────────────────────── serving query_all from the mirror ─────────────────────────
# `shared.qbo_api.query_all(access, cid, entity, where)` routes here for every
# mirrored entity, so every reader in the repo flips at once (the owner
# 2026-09-17: "rewrite ALL codes"). The QBO WHERE grammar the tools actually use
# is small and AND-only: `Field op value` with op in = != < <= > >= IN LIKE, over
# TxnDate / Balance / Active / Id / DocNumber / *Ref / DisplayName / AccountType /
# Classification / MetaData.LastUpdatedTime. Anything else raises - never a
# silent mismatch. `ACB_QBO_LIVE=1` sends every query to QBO (the parity check,
# or a machine without the mirror); a missing or empty mirror also falls back.

import os as _os
import re as _re

MAX_AGE_HOURS = 12                         # older than this -> loud warning
_STAMP_SHOWN = [False]
_TERM_RE = _re.compile(r"^\s*([A-Za-z_.]+)\s*(>=|<=|!=|=|>|<|IN|LIKE)\s*(.+?)\s*$", _re.I)
_PUSHDOWN = {"txndate": "txn_date", "id": "id", "docnumber": "doc_number"}


class WhereError(ValueError):
    """A QBO WHERE clause the mirror does not understand - route it live."""


def _split_and(where: str) -> List[str]:
    """Split on AND outside quotes/parentheses."""
    parts, buf, q, depth = [], [], False, 0
    i, s = 0, where
    while i < len(s):
        c = s[i]
        if c == "'":
            q = not q
        elif not q and c == "(":
            depth += 1
        elif not q and c == ")":
            depth -= 1
        if not q and depth == 0 and s[i:i + 5].upper() == " AND ":
            parts.append("".join(buf))
            buf = []
            i += 5
            continue
        buf.append(c)
        i += 1
    parts.append("".join(buf))
    return [p.strip() for p in parts if p.strip()]


def _parse_value(tok: str):
    tok = tok.strip()
    if tok.startswith("(") and tok.endswith(")"):
        inner, items, buf, q = tok[1:-1], [], [], False
        for c in inner:
            if c == "'":
                q = not q
            if c == "," and not q:
                items.append("".join(buf))
                buf = []
            else:
                buf.append(c)
        items.append("".join(buf))
        return [_parse_value(x) for x in items if x.strip()]
    if len(tok) >= 2 and tok[0] == "'" and tok[-1] == "'":
        return tok[1:-1].replace("''", "'")
    low = tok.lower()
    if low in ("true", "false"):
        return low == "true"
    try:
        return float(tok)
    except ValueError:
        raise WhereError(f"unreadable value {tok!r}")


def parse_where(where: str) -> List[Tuple[str, str, object]]:
    """'TxnDate >= '2026-01-01' AND Balance > '0'' -> [(field, OP, value), ...]"""
    out = []
    for term in _split_and(where or ""):
        m = _TERM_RE.match(term)
        if not m:
            raise WhereError(f"unreadable term {term!r}")
        out.append((m.group(1), m.group(2).upper(), _parse_value(m.group(3))))
    return out


def _field(rec: dict, name: str):
    cur = rec
    for part in name.split("."):
        if not isinstance(cur, dict):
            return None
        cur = cur.get(part)
    if isinstance(cur, dict) and "value" in cur:      # a Ref
        return cur.get("value")
    return cur


def _norm(actual, expected):
    """Bring both sides to one type: numbers when either is numeric-looking,
    bools, else strings (dates/ids/doc numbers compare as text, like QBO)."""
    if isinstance(expected, bool):
        return bool(actual) if actual is not None else False, expected
    if isinstance(actual, bool):
        return actual, str(expected).lower() == "true"
    if isinstance(actual, (int, float)):
        try:
            return float(actual), float(expected)
        except (TypeError, ValueError):
            return str(actual), str(expected)
    if isinstance(expected, float):
        try:
            return float(actual), expected
        except (TypeError, ValueError):
            return actual, expected
    return ("" if actual is None else str(actual)), str(expected)


def _like(actual: str, pattern: str) -> bool:
    """QBO's LIKE is case-insensitive (parity check 2026-09-17: '%Concrete%'
    returned 98 vendors live, 53 case-sensitively)."""
    rx = "^" + ".*".join(_re.escape(p) for p in pattern.split("%")) + "$"
    return _re.match(rx, actual, _re.S | _re.I) is not None


def match_where(rec: dict, terms: List[Tuple[str, str, object]]) -> bool:
    for field, op, expected in terms:
        actual = _field(rec, field)
        if op == "IN":
            vals = expected if isinstance(expected, list) else [expected]
            if not any(_norm(actual, v)[0] == _norm(actual, v)[1] for v in vals):
                return False
            continue
        if op == "LIKE":
            if actual is None or not _like(str(actual), str(expected)):
                return False
            continue
        a, e = _norm(actual, expected)
        if actual is None and op in (">", ">=", "<", "<="):
            return False
        ok = {"=": a == e, "!=": a != e, ">": a > e, ">=": a >= e, "<": a < e, "<=": a <= e}[op]
        if not ok:
            return False
    return True


def _pushdown(terms) -> Tuple[str, list]:
    """Terms on TxnDate / Id / DocNumber also run as SQL on the summary columns
    so a 100k-row entity is not decompressed for a one-week window."""
    sql, params = [], []
    for field, op, expected in terms:
        col = _PUSHDOWN.get(field.lower())
        if not col:
            continue
        if op == "IN" and isinstance(expected, list):
            vals = [str(v) for v in expected]
            sql.append(f"{col} IN ({','.join('?' * len(vals))})")
            params.extend(vals)
        elif op in ("=", "!=", ">", ">=", "<", "<=") and not isinstance(expected, list):
            sql.append(f"{col} {op} ?")
            params.append(str(expected))
    return " AND ".join(sql), params


def stamp(con: Optional[sqlite3.Connection] = None) -> Optional[dt.datetime]:
    own = con is None
    con = con or connect()
    try:
        s = _meta_get(con, "last_refresh")
        return parse_iso(s) if s else None
    finally:
        if own:
            con.close()


def serves(entity: str) -> bool:
    """Should `query_all` answer this entity from the mirror? No when the owner
    forces live (ACB_QBO_LIVE=1), the entity is not mirrored, or the mirror has
    never been seeded on this machine."""
    if _os.environ.get("ACB_QBO_LIVE", "").strip() in ("1", "true", "yes"):
        return False
    if entity not in ENTITIES or not db_path().exists():
        return False
    return stamp() is not None


def announce(con: Optional[sqlite3.Connection] = None) -> None:
    """Print the mirror's stamp once per process; loud when stale."""
    if _STAMP_SHOWN[0]:
        return
    _STAMP_SHOWN[0] = True
    t = stamp(con)
    if t is None:
        return
    age = now_utc() - t
    hours = age.total_seconds() / 3600
    local = t.astimezone().strftime("%m/%d/%Y %I:%M %p")
    print(f"   QBO mirror as of {local} ({hours:.1f} h ago)")
    if hours > MAX_AGE_HOURS:
        print(f"   WARNING: mirror is older than {MAX_AGE_HOURS} h - run sync-all "
              f"(or ledger/refresh_mirror.py) before trusting these numbers")


def query(entity: str, where: str = "", con: Optional[sqlite3.Connection] = None) -> List[dict]:
    """The mirror's answer to `SELECT * FROM <entity> [WHERE …]`. Same rows QBO
    would return: deleted rows excluded; a name list hides inactive rows unless
    the WHERE mentions Active (QBO's own default)."""
    terms = parse_where(where)
    own = con is None
    con = con or connect()
    try:
        announce(con)
        sql, params = _pushdown(terms)
        rows = load(entity, sql, tuple(params), con=con)
    finally:
        if own:
            con.close()
    if ENTITIES[entity]["list"] and not any(f.lower() == "active" for f, _, _ in terms):
        rows = [r for r in rows if r.get("Active", True) is not False]
    return [r for r in rows if match_where(r, terms)]
