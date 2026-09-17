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


def _pack(rec: dict) -> bytes:
    return zlib.compress(json.dumps(rec, separators=(",", ":")).encode("utf-8"))


def _unpack(blob: bytes) -> dict:
    return json.loads(zlib.decompress(blob).decode("utf-8"))


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
    name = rec.get("DisplayName") or rec.get("Name") or rec.get("FullyQualifiedName")
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
