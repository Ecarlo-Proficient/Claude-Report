"""shared/statement_record.py - the vendor statement RECORD the ledger reads (10/09/2026).

The owner: "make the vendor ledger have a record copy of vendor statements so it can
show, when paying, reconciled as of ... bring it into the Vendor Center; when opened we
see the vendor statement and the highlighted ones with the legend floating as you scroll".

Two tools share this file, so it lives in shared/ (repo rule: a file a second tool needs
moves here, never a cross-tool import):

  statement-reconciler/   WRITES one record per vendor at the end of every live run -
                          `<Vendor>/.statement-record.json` beside `.reconciler.json`, and the
                          marked-up statement pages as PNGs under `<Vendor>/.marked/`
                          (the same images that go into the Excel's "Statement (marked)" sheet).
  ledger/                 READS every record (the Vendor Center bubble, the pay run's
                          "reconciled as of" standing) and one record in full (the Statement
                          view: pages + legend + every line with its band on the page).

The record is the reconciler's answer frozen at the moment it ran - the ledger never
re-derives a bucket, never reads a statement itself, never writes here. The Excel in
`<Vendor>/Current` stays the filed evidence; Notion stays where the clerk works
(notes, Re-check); this record is what the ledger shows.

Record (version 1):
  vendor, vendor_id, folder            QBO vendor name, its id, the share folder name
  status, standing                     as on the Notion board (All entered / Not entered /
                                       Unreadable / No statement) + the sentence
  as_of, checked, checked_at           statement as-of (ISO), when checked (mm/dd/yyyy hh:mm),
                                       and the same as ISO
  clean, tieout, gap, amount_due, merged_total, matched, counts {kind: n}
  statements  [{as_of, label, kind, file}]      every statement the vendor sent, newest first
  files       ["Current/<statement>"]           the statement(s) the standing is built from
  excel       "Current/Reconciliation - ..."    the vendor's one Excel
  notion_url
  pages       [{file, width, height, source}]  the marked pages, in order, across the
                                                current statements (file under the vendor dir)
  lines       [{ref, date, amount, bucket, bill_id, job, note, qbo_ref, qbo_amount,
                qbo_date, file, page, top, bottom}]   every statement line as banded; page
                is an index into `pages` (None = read but not found on the page); top /
                bottom are image pixels. Also every QBO bill not on the statement
                (bucket "unlisted", no page).
  unread      [{page, top, bottom, text}]     rows on the page that look like a bill with no
                                              band - the tool may have missed them
  not_found   ["ref · date · amount"]         lines read but not located on any page
  located, total                             the "did it get everything" count

Vendor matching (ledger side): `key()` keeps only letters and digits, upper-cased, so
"BURNCO TEXAS, LLC" in the Bill Tracker meets "Burnco Texas LLC" from QuickBooks and
"R.C.I." meets "RCI". The JS twin is stmtKey() in ledger/static/app.js - keep them equal.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
from pathlib import Path
from typing import Dict, List, Optional

from shared import paths

VERSION = 1
RECORD_FILE = ".statement-record.json"
MARKED_DIR = ".marked"
INBOX_NAMES = ("inbox", "_inbox", "statement inbox", "statements inbox")

# bucket key -> label (the SAME words as statement_markup.BUCKETS, repeated here so the
# ledger, which never imports the reconciler, can label a legend) and its colour
BUCKETS: Dict[str, tuple] = {
    "matched":  ("Matched - in QBO, same amount",               (76, 175, 80)),
    "approve":  ("Matched, NOT APPROVED (memo) - chase the PM",  (205, 220, 57)),
    "checkqbo": ("Matched, approval is in QBO - check there",   (0, 172, 193)),
    "mismatch": ("Amount differs from QBO",                     (255, 152, 0)),
    "tax":      ("Sales tax added",                             (255, 112, 67)),
    "lag":      ("Paid in QBO, vendor still shows it open",     (66, 165, 245)),
    "enter":    ("Not in QBO - enter it",                       (229, 57, 53)),
    "unlisted": ("Open in QBO, not on the statement",           (121, 85, 72)),
    "other":    ("Payment / credit / balance forward",          (149, 117, 205)),
    "skipped":  ("Another customer's line - left out on purpose", (158, 158, 158)),
    "unread":   ("On the statement, not read - check by hand",  (198, 40, 40)),
}
STATUSES = ("All entered", "Not entered", "Unreadable", "No statement")


def root() -> Path:
    """The Vendor Statements root on the Accounting share (one vendor folder each).
    Override: ACB_STATEMENTS_ROOT (a test fixture, or the SharePoint move later)."""
    env = os.environ.get("ACB_STATEMENTS_ROOT", "").strip()
    if env:
        return Path(env).expanduser()
    return paths.accounting_base() / "Accounts Payable" / "Vendor Statements"


def key(name: str) -> str:
    """Fold a vendor name so the Bill Tracker's spelling meets QuickBooks' and the folder's."""
    return re.sub(r"[^A-Z0-9]+", "", (name or "").upper())


def us_date(iso: str) -> str:
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    return f"{m.group(2)}/{m.group(3)}/{m.group(1)}" if m else (iso or "")


# ─────────────────────────── writing (the reconciler) ───────────────────────────

def write(vendor_dir: Path, record: dict, pages: Optional[List[object]] = None) -> Path:
    """Save one vendor's record + its marked pages. `pages` = PIL images in the order
    record["pages"] names them (each entry's `file` is set here). Old marked pages go
    first, so a statement replaced leaves no stale picture behind. Atomic on the JSON."""
    vendor_dir.mkdir(parents=True, exist_ok=True)
    marked = vendor_dir / MARKED_DIR
    if pages is not None:
        if marked.is_dir():
            for f in marked.iterdir():
                if f.suffix.lower() == ".png":
                    try:
                        f.unlink()
                    except OSError:
                        pass
        marked.mkdir(exist_ok=True)
        metas = record.setdefault("pages", [])
        for i, img in enumerate(pages):
            name = f"page-{i + 1:02d}.png"
            img.save(marked / name, format="PNG", optimize=True)
            meta = metas[i] if i < len(metas) else {}
            if i >= len(metas):
                metas.append(meta)
            meta["file"] = f"{MARKED_DIR}/{name}"
            meta["width"], meta["height"] = img.size
    record["version"] = VERSION
    record.setdefault("written", dt.datetime.now().isoformat(timespec="seconds"))
    body = json.dumps(record, indent=1, sort_keys=True, default=str)
    tmp = vendor_dir / (RECORD_FILE + ".tmp")
    tmp.write_text(body)
    tmp.replace(vendor_dir / RECORD_FILE)
    return vendor_dir / RECORD_FILE


def build(result: dict, *, status: str, standing: str, notion_url: str = "",
          markups: Optional[list] = None, now: Optional[dt.datetime] = None) -> dict:
    """One vendor's reconcile `result` (statement_reconciler) + the marked statements ->
    the record dict (pages' files are filled by write()). Pure: no I/O.

    result["_record"] = {"rows": [...], "files": [...], "excel": "..."} is what
    reconcile_vendor stashes for this: each row = one ReconRow as a dict (category,
    bucket, stmt_ref, qbo_ref, stmt_date, qbo_date, stmt_amount, qbo_amount, qbo_bill_id,
    job, note).
    markups = [(statement file name, Result)] from statement_markup.mark - Result.bands
    carry where each line sits on which page."""
    now = now or dt.datetime.now()
    rec_in = result.get("_record") or {}
    rows = rec_in.get("rows") or []
    by_ref: Dict[str, dict] = {}
    for r in rows:
        ref = key(r.get("stmt_ref") or "")
        if ref and ref not in by_ref:
            by_ref[ref] = r
    pages: List[dict] = []
    lines: List[dict] = []
    unread: List[dict] = []
    not_found: List[str] = []
    located = total = 0
    seen_refs = set()
    offset = 0
    for name, res in (markups or []):
        res_pages = getattr(res, "pages", []) or []
        for _img in res_pages:
            pages.append({"source": name})
        n_lines = len(getattr(res, "lines", []) or [])
        band_by_line: Dict[int, dict] = {}
        for b in getattr(res, "bands", []) or []:
            band_by_line.setdefault(int(b["line"]), b)     # the first band = the line's own row
        for i, ln in enumerate(getattr(res, "lines", []) or []):
            ref = getattr(ln, "ref", "") or ""
            amt = float(getattr(ln, "amount", 0.0) or 0.0)
            row = by_ref.get(key(ref)) if ref else None
            band = band_by_line.get(i)
            bucket = (band or {}).get("bucket") or (row or {}).get("bucket") or "other"
            if ref and amt >= 0 and row is None and bucket == "other":
                bucket = "other"
            entry = {
                "ref": ref, "date": getattr(ln, "date", "") or "", "amount": round(amt, 2),
                "bucket": bucket, "file": name,
                "page": (offset + int(band["page"])) if band else None,
                "top": band.get("top") if band else None, "bottom": band.get("bottom") if band else None,
                "bill_id": (row or {}).get("qbo_bill_id") or "", "job": (row or {}).get("job") or "",
                "note": (row or {}).get("note") or "", "qbo_ref": (row or {}).get("qbo_ref") or "",
                "qbo_amount": (row or {}).get("qbo_amount"), "qbo_date": (row or {}).get("qbo_date") or "",
            }
            lines.append(entry)
            if ref:
                seen_refs.add(key(ref))
        for u in getattr(res, "unread", []) or []:
            unread.append({"page": offset + int(u["page"]), "top": u.get("top"), "bottom": u.get("bottom"),
                           "text": (u.get("text") or "")[:120]})
        not_found += [f"{name}: {t}" for t in (getattr(res, "not_found", []) or [])]
        located += int(getattr(res, "located", 0) or 0)
        total += int(getattr(res, "total", 0) or 0) if n_lines else 0
        offset += len(res_pages)
    # statement rows the markup never saw (no markups at all - e.g. an Excel statement that
    # could not be rendered) still make the legend, without a page
    for r in rows:
        ref = r.get("stmt_ref") or ""
        if r.get("category") == "MISSING_ON_STATEMENT":
            lines.append({"ref": r.get("qbo_ref") or "", "date": r.get("qbo_date") or "", "amount": round(float(r.get("qbo_amount") or 0), 2),
                          "bucket": "unlisted", "file": "", "page": None, "top": None, "bottom": None,
                          "bill_id": r.get("qbo_bill_id") or "", "job": r.get("job") or "", "note": r.get("note") or "",
                          "qbo_ref": r.get("qbo_ref") or "", "qbo_amount": r.get("qbo_amount"), "qbo_date": r.get("qbo_date") or ""})
        elif ref and key(ref) not in seen_refs:
            lines.append({"ref": ref, "date": r.get("stmt_date") or "", "amount": round(float(r.get("stmt_amount") or 0), 2),
                          "bucket": r.get("bucket") or "other", "file": "", "page": None, "top": None, "bottom": None,
                          "bill_id": r.get("qbo_bill_id") or "", "job": r.get("job") or "", "note": r.get("note") or "",
                          "qbo_ref": r.get("qbo_ref") or "", "qbo_amount": r.get("qbo_amount"), "qbo_date": r.get("qbo_date") or ""})
            seen_refs.add(key(ref))
    counts: Dict[str, int] = {}
    for ln in lines:
        counts[ln["bucket"]] = counts.get(ln["bucket"], 0) + 1
    if unread:
        counts["unread"] = len(unread)
    checked = result.get("checked") or now.strftime("%m/%d/%Y %I:%M %p")
    return {
        "version": VERSION,
        "vendor": result.get("vendor") or "", "vendor_id": str(result.get("vendor_id") or ""),
        "folder": result.get("vendor_folder") or result.get("vendor") or "",
        "status": status, "standing": standing,
        "as_of": result.get("as_of") or "", "checked": checked,
        "checked_at": now.isoformat(timespec="seconds"),
        "clean": bool(result.get("clean")), "tieout": bool(result.get("tieout", True)),
        "gap": round(float(result.get("gap") or 0), 2),
        "amount_due": result.get("amount_due"), "merged_total": result.get("merged_total"),
        "matched": result.get("matched", 0), "counts": counts,
        "open_items": result.get("open_items") or {},
        "statements": result.get("statements") or [],
        "files": rec_in.get("files") or [], "excel": rec_in.get("excel") or "",
        "notion_url": notion_url or "",
        "pages": pages, "lines": lines, "unread": unread, "not_found": not_found,
        "located": located, "total": total,
    }


# ─────────────────────────── reading (the ledger) ───────────────────────────

def vendor_dirs(base: Optional[Path] = None) -> List[Path]:
    """Every vendor folder under the root (the Inbox and dot-folders left out)."""
    base = base or root()
    if not base.is_dir():
        return []
    out = []
    for d in sorted(base.iterdir()):
        if not d.is_dir() or d.name.startswith(".") or d.name.lower() in INBOX_NAMES:
            continue
        out.append(d)
    return out


def read(vendor_dir: Path) -> Optional[dict]:
    try:
        rec = json.loads((vendor_dir / RECORD_FILE).read_text())
    except (OSError, ValueError):
        return None
    if not isinstance(rec, dict):
        return None
    rec["_dir"] = str(vendor_dir)
    return rec


def summary(rec: dict) -> dict:
    """The short form the Vendor Center and the pay run show (no lines, no pages)."""
    return {k: rec.get(k) for k in ("vendor", "vendor_id", "folder", "status", "standing", "as_of", "checked",
                                    "checked_at", "clean", "tieout", "gap", "amount_due", "merged_total",
                                    "matched", "counts", "open_items", "notion_url", "located", "total")} | {
        "pages": len(rec.get("pages") or []), "lines": len(rec.get("lines") or []),
        "unread": len(rec.get("unread") or [])}


def read_all(base: Optional[Path] = None) -> dict:
    """{"mounted": bool, "root": str, "vendors": {key: summary}} - every vendor folder, with a
    record when the reconciler has written one, else {"status": "No record", ...} so the
    ledger can say the folder exists but the tool has not run on it since this record began."""
    base = base or root()
    out: Dict[str, dict] = {}
    mounted = base.is_dir()
    for d in vendor_dirs(base) if mounted else []:
        rec = read(d)
        if rec:
            s = summary(rec)
            s["dir"] = d.name
        else:
            s = {"vendor": d.name, "folder": d.name, "status": "No record", "standing": "Folder on the share, not reconciled since the record began",
                 "as_of": "", "checked": "", "counts": {}, "pages": 0, "lines": 0, "dir": d.name}
        for k in {key(s.get("vendor") or ""), key(d.name)}:
            if k and (k not in out or out[k].get("status") == "No record"):
                out[k] = s
    return {"mounted": mounted, "root": str(base), "vendors": out}


def find(name: str, base: Optional[Path] = None) -> Optional[dict]:
    """The full record for a vendor named as the ledger names it (QBO name), or by folder."""
    base = base or root()
    want = key(name)
    if not want or not base.is_dir():
        return None
    for d in vendor_dirs(base):
        if key(d.name) == want:
            rec = read(d)
            if rec:
                return rec
    for d in vendor_dirs(base):
        rec = read(d)
        if rec and key(rec.get("vendor") or "") == want:
            return rec
    return None


def resolve_file(rec: dict, rel: str) -> Optional[Path]:
    """A path named INSIDE the record (a page, a statement, the Excel) -> the file on disk,
    only when it sits under that vendor's folder. The request never names a path."""
    base = Path(rec.get("_dir") or "")
    if not rel or not base.is_dir():
        return None
    target = (base / rel).resolve()
    try:
        target.relative_to(base.resolve())
    except ValueError:
        return None
    return target if target.is_file() else None
