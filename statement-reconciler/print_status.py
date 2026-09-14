"""Print-status verification for the statement reconciler  (ties AP-03 -> AP-01).

WHY
    AP intake (vault process AP-01): a vendor bill lands in the billings inbox,
    the bill clerk PRINTS it, tags the email with the "printed" category, and
    enters it in QBO. In that pipeline the email category *is* the bill's state -
    and the vault's own note is blunt about the hole: "an untagged email is an
    invisible bill - no queue, no aging, no count, and no way to notice one that
    never got printed."

    This module closes that hole from the reconciliation side. For a vendor
    STATEMENT we already parse (every invoice the vendor says is open), it asks
    the billings inbox: was each of those invoices ever received-and-printed?
    A statement invoice with no printed email is a bill that slipped intake.

NAMES POLICY
    The printed category label is a person's handle in the mail system
    ("<name> Printed"). It is NEVER hard-coded here - it comes from config
    (PRINTED_EMAIL_CATEGORY in machine.env). The person who prints is "the bill
    clerk" everywhere in text.

EMAIL ACCESS IS PLUGGABLE  (settle this before wiring in - see the two loaders)
    PrintedIndex.from_export_csv(path)
        An Outlook-exported CSV of the billings folder (Subject + Categories +
        Received columns). Zero infrastructure, works today; the bill clerk (or
        a saved Outlook view) exports it. Good enough to ship the sheet.
    PrintedIndex.from_graph(mailbox, ...)
        Microsoft Graph / Exchange Online - the live path. Needs an Azure AD app
        registration with Mail.Read (admin consent) and the mailbox address, with
        the client-secret going in the ONE automation-qbo Keychain blob (key
        library rule). STUBBED until we settle auth (raises NotImplementedError).
"""
from __future__ import annotations

import csv
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set


# Printed = the email's category marks the bill printed. The billing inbox uses
# per-clerk labels ("ITZ PRINTED", "GISELLE PRINTED"), and "PRINT PENDING" is
# NOT printed - so match the whole word "printed", never the substring "print".
# No personal name is hard-coded: the keyword is generic (override with
# PRINTED_CATEGORY_KEYWORD if the convention ever changes).
PRINTED_KEYWORD = os.environ.get("PRINTED_CATEGORY_KEYWORD", "printed").strip()
_PRINTED_RE = re.compile(rf"\b{re.escape(PRINTED_KEYWORD)}\b", re.I)


def _is_printed_category(categories: str) -> bool:
    """True if any of the email's categories marks it printed (whole-word match,
    so 'ITZ PRINTED'/'GISELLE PRINTED' count but 'PRINT PENDING' does not)."""
    return bool(categories) and bool(_PRINTED_RE.search(categories))


def _norm_ref(ref: str) -> str:
    """Normalize an invoice identifier for matching: upper, strip non-alnum.
    'INV #38,486.' and 'inv38486' both -> 'INV38486'. Statement refs and the
    invoice number an email carries rarely agree on punctuation/case."""
    return re.sub(r"[^A-Za-z0-9]", "", (ref or "")).upper()


def _digit_core(ref: str) -> str:
    """The longest run of digits in an identifier - the part vendor prefixes and
    punctuation vary around. Statement 'INV38486' and email '#38486' share the
    core '38486'; 'INV_SUN15251' and 'SUN15251' share '15251'. Secondary key, so
    a full-token match wins first (digit-only carries a small collision risk)."""
    runs = re.findall(r"\d+", ref or "")
    return max(runs, key=len) if runs else ""


def _digits_all(s: str) -> str:
    """Every digit concatenated, leading zeros stripped - the last-resort key for
    zero-padded encodings. Sunbelt's attachment names carry '0001879912780001'
    for invoice '187991278-0001'; both reduce to '1879912780001'."""
    return re.sub(r"\D", "", s or "").lstrip("0")


# Invoice-number-ish tokens: any run of ref-ish chars carrying 4+ digits
# ('INV_SUN15251', '38486', '16018K', 'D80025'). Fed from subject + body +
# attachment filenames (filename separators are pre-split to spaces so each
# segment tokenizes on its own).
_SUBJECT_REF_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9_\-]*\d{4,}[A-Za-z0-9_\-]*|\d{4,}")

_TAG_RE = re.compile(r"<[^>]+>")
_ENT_RE = re.compile(r"&(?:#\d+|#x[0-9a-fA-F]+|[a-zA-Z]+);")


def _html_to_text(content: str) -> str:
    """Flatten an HTML email body to plain text for token scanning. Vendors like
    White Cap put the invoice # in the body TABLE, below the 255-char bodyPreview
    cutoff - so we index the full body, tags stripped. Cheap and dependency-free:
    the goal is only to expose the digit runs, not to render."""
    if not content:
        return ""
    t = _TAG_RE.sub(" ", content)
    t = _ENT_RE.sub(" ", t)
    return t


def _extract_tokens(subject: str, body: str, filenames) -> tuple:
    """The invoice-ref tokens in an email's subject + body + attachment filenames.
    Filename separators (_ - .) are pre-split to spaces so each segment tokenizes
    on its own. This is the ONE place raw tokens are pulled - both the live keying
    and the cache reload go through it, so a cached index keys identically."""
    text = " ".join([subject or "", body or "",
                     *[re.sub(r"[_\-.]", " ", fn) for fn in (filenames or ())]])
    return tuple(m.group(0) for m in _SUBJECT_REF_RE.finditer(text))


@dataclass
class PrintedEmail:
    """One billings-inbox message that carries a printed bill."""
    subject: str
    received: str                       # ISO-ish date string, best effort
    categories: str
    body: str = ""                      # full body text (invoice # for Cowtown/White Cap)
    filenames: tuple = ()               # attachment names (invoice # for CMC/Sunbelt)
    msg_id: str = ""                    # Graph message id (cache dedupe / incremental key)
    tokens: tuple = ()                  # raw invoice-ref tokens (cached, so reload re-keys)
    refs: Set[str] = field(default_factory=set)   # normalized invoice tokens


@dataclass
class PrintedIndex:
    """Every printed invoice identifier found in the billings inbox, plus the
    email each came from, so the sheet can show the proof (date/subject)."""
    by_id: Dict[str, PrintedEmail] = field(default_factory=dict)         # msg id -> email (authoritative)
    by_ref: Dict[str, PrintedEmail] = field(default_factory=dict)        # full normalized token (derived)
    by_core: Dict[str, PrintedEmail] = field(default_factory=dict)       # longest digit run (derived)
    by_alldigits: Dict[str, PrintedEmail] = field(default_factory=dict)  # all digits, 0-stripped (derived)
    printed_count: int = 0
    max_lastmod: str = ""               # max lastModifiedDateTime seen (incremental floor)
    source: str = ""                    # where this index came from (for the sheet)

    # ---- loaders -------------------------------------------------------
    @classmethod
    def from_export_csv(cls, csv_path: Path,
                        subject_col: str = "Subject",
                        categories_col: str = "Categories",
                        received_col: str = "Received") -> "PrintedIndex":
        """Build from an Outlook-exported CSV of the billings folder. Only rows
        whose category marks them printed are indexed. Column names match a
        default Outlook export; override if the export differs."""
        idx = cls(source=f"Outlook export: {csv_path.name}")
        with open(csv_path, newline="", encoding="utf-8-sig") as fh:
            for row in csv.DictReader(fh):
                cats = (row.get(categories_col) or "").strip()
                if not _is_printed_category(cats):
                    continue
                idx.printed_count += 1
                idx._add(PrintedEmail(
                    subject=(row.get(subject_col) or "").strip(),
                    received=(row.get(received_col) or "").strip(),
                    categories=cats))
        return idx

    @classmethod
    def from_graph(cls, mailbox: Optional[str] = None,
                   creds: Optional[dict] = None,
                   since_days: int = 550,
                   prior: Optional["PrintedIndex"] = None) -> "PrintedIndex":
        """Live read via Microsoft Graph (Exchange Online). Reads GRAPH_TENANT_ID
        / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET / GRAPH_BILLING_MAILBOX from the
        automation-qbo Keychain blob (one Touch ID) unless `creds` is passed.
        Pulls each message's subject + FULL body + attachment names and indexes
        the printed ones. Retries transient timeouts/5xx/429.

        Incremental: pass `prior` (a cached index). Only messages CHANGED since the
        prior pull are fetched (lastModifiedDateTime floor) - a new bill AND an old
        email that was just tagged printed both bump lastModifiedDateTime, so both
        are caught; an email whose printed tag was removed is dropped. Without
        `prior` it does a full pull back to the receivedDateTime horizon."""
        import requests
        import time
        if creds is None:
            import sys as _sys
            _sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
            from shared import qbo_vault as kc
            creds = kc.get_all()
        mailbox = mailbox or creds.get("GRAPH_BILLING_MAILBOX", "")
        if not mailbox:
            raise ValueError("no mailbox: set GRAPH_BILLING_MAILBOX or pass mailbox=")
        access = _graph_token(creds)

        idx = cls(source=f"Graph: {mailbox}")
        # Seed from the prior cache so an incremental pull only has to apply the
        # delta on top of what we already knew.
        if prior is not None:
            idx.by_id = {i: e for i, e in prior.by_id.items()}
            idx.max_lastmod = prior.max_lastmod

        # No server-side category filter: the inbox uses several printed labels
        # (per clerk), so we page messages and keep the printed ones by the
        # whole-word "printed" test. A receivedDateTime floor bounds the horizon
        # (statements only list recent-enough invoices, and it cuts cross-year
        # invoice-number collisions); an incremental run also floors on
        # lastModifiedDateTime so it fetches only what changed.
        import datetime as _dt
        from urllib.parse import quote
        floor = (_dt.datetime.utcnow() - _dt.timedelta(days=since_days)).strftime("%Y-%m-%dT00:00:00Z")
        filt = f"receivedDateTime ge {floor}"
        if prior is not None and prior.max_lastmod:
            filt += f" and lastModifiedDateTime gt {prior.max_lastmod}"
        # Full body (invoice # lives in the body table for White Cap/Cowtown, past
        # the 255-char bodyPreview) + attachment NAMES ($expand $select name =
        # metadata only, NOT the PDF bytes, so CMC/Sunbelt decode from the filename
        # without ever downloading the attachment). id + lastModifiedDateTime drive
        # the cache dedupe / incremental floor.
        url = (f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages"
               f"?$select=id,subject,categories,receivedDateTime,lastModifiedDateTime,body"
               f"&$expand=attachments($select=name)&$top=50"
               f"&$filter={quote(filt)}")
        headers = {"Authorization": f"Bearer {access}"}

        def _get(u):                       # retry transient timeouts / 5xx / 429
            for attempt in range(5):
                try:
                    r = requests.get(u, headers=headers, timeout=120)
                    if r.status_code in (429, 500, 502, 503, 504):
                        raise requests.exceptions.RequestException(f"HTTP {r.status_code}")
                    r.raise_for_status()
                    return r.json()
                except requests.exceptions.RequestException:
                    if attempt == 4:
                        raise
                    time.sleep(2 ** attempt)   # 1,2,4,8,16s

        pages = 0
        while url and pages < 800:            # hard cap ~40k messages
            page = _get(url)
            for msg in page.get("value", []):
                mid = msg.get("id") or ""
                lm = msg.get("lastModifiedDateTime") or ""
                if lm > idx.max_lastmod:
                    idx.max_lastmod = lm
                cats = ", ".join(msg.get("categories") or [])
                if not _is_printed_category(cats):
                    idx.by_id.pop(mid, None)   # tag cleared since last pull -> drop it
                    continue
                atts = msg.get("attachments") or []
                bodytext = _html_to_text((msg.get("body") or {}).get("content", ""))
                fns = tuple(a.get("name", "") for a in atts if a.get("name"))
                em = PrintedEmail(
                    subject=(msg.get("subject") or "").strip(),
                    received=(msg.get("receivedDateTime") or "")[:10],
                    categories=cats, body=bodytext, filenames=fns, msg_id=mid,
                    tokens=_extract_tokens((msg.get("subject") or ""), bodytext, fns))
                em.body = ""               # tokens are extracted; don't carry the HTML
                idx.by_id[mid] = em
            url = page.get("@odata.nextLink")
            pages += 1
        idx.rebuild()
        return idx

    # ---- keying --------------------------------------------------------
    def _key(self, em: "PrintedEmail") -> None:
        """Index one email under every invoice-ref token it carries - three keys
        each (full token, longest digit run, all-digits-zero-stripped) so a
        statement ref matches however the vendor encodes it. First writer wins on
        each key (setdefault), so a more-specific earlier email is not clobbered."""
        for tok in em.tokens:
            r = _norm_ref(tok)
            if r:
                em.refs.add(r)
                self.by_ref.setdefault(r, em)
            core = _digit_core(tok)
            if core:
                self.by_core.setdefault(core, em)
            ad = _digits_all(tok)
            if len(ad) >= 5:              # >=5 digits: cut date/short-number noise
                self.by_alldigits.setdefault(ad, em)

    def rebuild(self) -> None:
        """Rebuild the three derived lookup maps from by_id (the authoritative
        store). Called after any add/remove so the maps never carry a stale email."""
        self.by_ref.clear()
        self.by_core.clear()
        self.by_alldigits.clear()
        for em in self.by_id.values():
            em.refs = set()
            self._key(em)
        self.printed_count = len(self.by_id)

    def _add(self, em: "PrintedEmail") -> None:
        """Append one printed email (used by the CSV loader). Assigns a synthetic
        id when the source has none, stores it, extracts tokens, and keys it."""
        if not em.msg_id:
            em.msg_id = f"row{len(self.by_id)}"
        if not em.tokens:
            em.tokens = _extract_tokens(em.subject, em.body, em.filenames)
        self.by_id[em.msg_id] = em
        self._key(em)
        self.printed_count = len(self.by_id)

    # ---- cache (de)serialization --------------------------------------
    def to_dict(self) -> dict:
        """Serialize to a small JSON-able dict for the on-disk cache. Only the ref
        tokens and the sheet's display fields are kept - never the raw body/HTML."""
        return {
            "source": self.source,
            "max_lastmod": self.max_lastmod,
            "emails": [{"id": e.msg_id, "subject": e.subject, "received": e.received,
                        "categories": e.categories, "filenames": list(e.filenames),
                        "tokens": list(e.tokens)} for e in self.by_id.values()],
        }

    @classmethod
    def from_dict(cls, d: dict, received_floor: str = "") -> "PrintedIndex":
        """Rebuild an index from a cached dict. Emails received before
        `received_floor` (YYYY-MM-DD) are dropped so the horizon/collision guard
        holds as time moves forward even though the cache is never evicted."""
        idx = cls(source=d.get("source", ""))
        idx.max_lastmod = d.get("max_lastmod", "")
        for e in d.get("emails", []):
            rcv = e.get("received", "")
            if received_floor and rcv and rcv < received_floor:
                continue
            idx.by_id[e.get("id", "")] = PrintedEmail(
                subject=e.get("subject", ""), received=rcv,
                categories=e.get("categories", ""),
                filenames=tuple(e.get("filenames", [])),
                msg_id=e.get("id", ""), tokens=tuple(e.get("tokens", [])))
        idx.rebuild()
        return idx

    # ---- lookup --------------------------------------------------------
    def status_for(self, ref: str) -> Optional[PrintedEmail]:
        """The printed email for this statement invoice ref, or None if none is
        found (may never have been intake-printed). Tries full token, then the
        longest digit run, then all-digits-zero-stripped - most precise first."""
        em = self.by_ref.get(_norm_ref(ref))
        if em is not None:
            return em
        core = _digit_core(ref)
        em = self.by_core.get(core) if core else None
        if em is not None:
            return em
        ad = _digits_all(ref)
        return self.by_alldigits.get(ad) if len(ad) >= 5 else None


def _graph_token(creds: dict) -> str:
    """One client-credentials access token for the Graph mailbox reads."""
    import requests
    r = requests.post(
        f"https://login.microsoftonline.com/{creds['GRAPH_TENANT_ID']}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials",
              "client_id": creds["GRAPH_CLIENT_ID"],
              "client_secret": creds["GRAPH_CLIENT_SECRET"],
              "scope": "https://graph.microsoft.com/.default"},
        timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


# Live-search fallback state: a reused token + a per-invoice result cache so a
# statement (and a whole sweep) never searches the same number twice.
_SEARCH: dict = {"token": None, "creds": None, "cache": {}}


def live_search_printed(ref: str, creds: Optional[dict] = None,
                        mailbox: Optional[str] = None) -> Optional["PrintedEmail"]:
    """Fallback for an invoice the pre-built index missed: ask Graph whether ANY
    printed email contains this number. Graph's $search reads INSIDE PDF
    attachments server-side, so it catches invoices whose number appears only in
    the PDF - e.g. a bundled email with a generic filename (Croell 'Croell
    Invocies' -> 'Proficient Concrete Invoices 1.pdf'). We never download the PDF.

    Returns the printed email (subject/date for the sheet) or None. Per-number
    cached; a transient failure returns None WITHOUT caching, so a later line can
    retry. So a "NOT PRINTED" that survives this check means the number is in no
    printed email at all - the genuine finding, not a reader blind spot."""
    import requests
    # Search the longest alphanumeric SEGMENT, not digits-only: Graph tokenizes
    # "RW786083" and "188673772-0001" as whole words, so a digits-only term
    # ("786083" / "1886737720001") never matches. The longest segment
    # ("RW786083", "188673772") is how the number is actually indexed.
    segs = [s for s in re.split(r"[^A-Za-z0-9]+", ref or "") if s]
    term = max(segs, key=len) if segs else ""
    if len(term) < 5:
        term = _norm_ref(ref)
    if not term or len(term) < 5:
        return None
    if term in _SEARCH["cache"]:
        return _SEARCH["cache"][term]
    try:
        if creds is None:
            if _SEARCH["creds"] is None:
                import sys as _sys
                _sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
                from shared import qbo_vault as kc
                _SEARCH["creds"] = kc.get_all()
            creds = _SEARCH["creds"]
        mailbox = mailbox or creds.get("GRAPH_BILLING_MAILBOX", "")
        if _SEARCH["token"] is None:
            _SEARCH["token"] = _graph_token(creds)
        H = {"Authorization": f"Bearer {_SEARCH['token']}", "ConsistencyLevel": "eventual"}
        u = (f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages"
             f'?$search="{term}"&$select=subject,categories,receivedDateTime&$top=10')
        r = requests.get(u, headers=H, timeout=60)
        if r.status_code == 401:                 # token aged out mid-run -> refresh once
            _SEARCH["token"] = _graph_token(creds)
            H["Authorization"] = f"Bearer {_SEARCH['token']}"
            r = requests.get(u, headers=H, timeout=60)
        r.raise_for_status()
        hit = None
        for m in r.json().get("value", []):
            cats = ", ".join(m.get("categories") or [])
            if _is_printed_category(cats):
                hit = PrintedEmail(subject=(m.get("subject") or "").strip(),
                                   received=(m.get("receivedDateTime") or "")[:10],
                                   categories=cats)
                break
        _SEARCH["cache"][term] = hit             # cache the definite answer (hit or None)
        return hit
    except Exception:
        return None                              # transient: don't cache, allow a retry


_INDEX_CACHE: dict = {}


def _cache_path() -> Path:
    """On-disk printed-index cache. App state, not a log, so it lives under
    Application Support/Proficient (same home as the ledger DB) - override with
    PRINT_STATUS_CACHE_DIR. Never inside the repo (Claude-visible/synced)."""
    base = Path(os.environ.get(
        "PRINT_STATUS_CACHE_DIR",
        str(Path.home() / "Library" / "Application Support" / "Proficient" / "print-status")))
    base.mkdir(parents=True, exist_ok=True)
    return base / "printed_index.json"


def printed_index(force: bool = False, since_days: int = 550) -> Optional["PrintedIndex"]:
    """The printed index, process-cached (ONE build per run, reused across every
    statement) AND disk-cached (so a run only pulls what changed since last time,
    not the whole mailbox - the full-body pull is far too slow to repeat).

    Flow: load the disk cache -> incremental Graph refresh (only messages modified
    since the cached floor) -> save. `force=True` ignores the disk cache and does a
    full rebuild. Returns None only if Graph is unreachable AND there is no cache
    to fall back on (reason in _INDEX_CACHE['err']); callers then skip the sheet."""
    if not force and "v" in _INDEX_CACHE:
        return _INDEX_CACHE["v"]
    import datetime as _dt
    import json
    floor = (_dt.datetime.utcnow() - _dt.timedelta(days=since_days)).strftime("%Y-%m-%d")
    prior = None
    try:
        cp = _cache_path()
        if not force and cp.exists():
            try:
                prior = PrintedIndex.from_dict(json.loads(cp.read_text()), received_floor=floor)
            except Exception:
                prior = None            # corrupt/old cache -> full rebuild
        idx = PrintedIndex.from_graph(since_days=since_days, prior=prior)
        try:
            cp.write_text(json.dumps(idx.to_dict()))
        except Exception:
            pass                        # a cache we can't persist is not fatal
        _INDEX_CACHE["v"] = idx
    except Exception as e:
        # Graph failed: better to serve the (possibly slightly stale) cache than
        # to drop the sheet entirely.
        _INDEX_CACHE["v"] = prior
        _INDEX_CACHE["err"] = str(e)
    return _INDEX_CACHE["v"]


@dataclass
class PrintRow:
    """One statement invoice's print status, for the sheet."""
    date: str
    ref: str
    amount: float
    printed: bool
    email_date: str = ""
    email_subject: str = ""


def build_print_rows(statement_lines, index: PrintedIndex,
                     search_fn=None) -> List[PrintRow]:
    """Cross-reference every parsed statement invoice against the printed index.
    Lines with no ref (e.g. a Balance-forward lump) are skipped - nothing to
    match. Credits/payments (negative) are skipped too; they aren't printed bills.

    `search_fn` (e.g. `live_search_printed`): a fallback called only for invoices
    the pre-built index misses. It reads inside PDF attachments server-side, so a
    bill printed as a generically-named PDF is still found. With it, a remaining
    "NOT PRINTED" means the number is in no printed email at all."""
    out: List[PrintRow] = []
    for l in statement_lines:
        if not getattr(l, "ref", "") or getattr(l, "amount", 0) <= 0:
            continue
        em = index.status_for(l.ref)
        if em is None and search_fn is not None:
            em = search_fn(l.ref)
        out.append(PrintRow(
            date=getattr(l, "date", ""), ref=l.ref, amount=l.amount,
            printed=em is not None,
            email_date=em.received if em else "",
            email_subject=em.subject if em else "",
        ))
    return out


def write_print_status_sheet(wb, statement_lines, index: PrintedIndex,
                             title_font=None, header_fill=None,
                             search_fn=None) -> None:
    """Add a plain 'Print Status' sheet to an open workbook. Follows the repo's
    plain-Excel rule (white/black, label+amount on a row). Green/red is state,
    not decoration, so a single Printed? column is fine as text for now.
    `search_fn` is the live PDF-content fallback (see build_print_rows)."""
    from openpyxl.styles import Font, Alignment
    rows = build_print_rows(statement_lines, index, search_fn=search_fn)
    # Whole-statement 0% match: this vendor almost certainly carries the invoice #
    # only INSIDE the PDF (e.g. Ellis) - subject/body/filename never expose it. Do
    # NOT cry "NOT PRINTED" on every line (a wall of false alarms for the clerk);
    # mark them "unverified" and say so. A mixed result is trusted line by line.
    pdf_only = bool(rows) and len(rows) >= 3 and all(not r.printed for r in rows)
    ws = wb.create_sheet("Print Status")
    ws.sheet_view.showGridLines = False
    headers = ["Date", "Invoice #", "Amount", "Printed?", "Email date", "Email subject"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        status = "Yes" if r.printed else ("unverified" if pdf_only else "NOT PRINTED")
        ws.append([r.date, r.ref, r.amount, status, r.email_date, r.email_subject])
    not_printed = sum(1 for r in rows if not r.printed)
    ws.append([])
    if pdf_only:
        ws.append([f"{len(rows)} invoices  |  none matched an email - this vendor's invoice # "
                   f"likely appears only inside the PDF (not the subject/body/filename); open the "
                   f"PDF to verify before treating these as unprinted.  |  source: {index.source}"])
    else:
        ws.append([f"{len(rows)} invoices  |  {not_printed} NOT printed  |  source: {index.source}"])
    for col, w in zip("ABCDEF", (12, 16, 14, 14, 14, 60)):
        ws.column_dimensions[col].width = w
    return ws
