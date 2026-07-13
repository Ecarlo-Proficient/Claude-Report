"""
Invoice sync — QBO → Notion Invoice Tracker (Res/Com) + Invoice Tracker (MFD).

Semantics:
  Each QBO invoice is upserted into one of two Notion DBs based on its
  Project # prefix:
    - MFD###  → Invoice Tracker (MFD)
    - RP####  → Invoice Tracker (Res/Com), Division="RP"
    - CP###   → Invoice Tracker (Res/Com), Division="CP"
    - other / no project # → skipped + logged (cannot route)

  Match key (upsert): Invoice ID (QBO Id, hidden field on the page).

  Sync ONLY writes the QBO-sourced fields. Human-owned fields are never
  touched on update:
    - Quick Status, Next Follow-Up, Owner, Last Action Date
    - Page body (Collection Log)

Filter strategy (open-only, with delta-based paid detection):
  - QBO query is restricted to OPEN invoices (Balance > 0). This is the
    smallest possible payload and dodges Notion rate limits.
  - For every open QBO invoice: upsert into the routed Notion DB.
  - For every Notion page currently marked open (Status != Paid): if its
    Invoice ID is NO LONGER in the QBO open set, the invoice was paid
    out-of-band — flip Status to Paid and zero the Open balance. This
    is a delta detection, not a paid-history fetch.
  - Notion retains paid invoices for 12 months from TxnDate, then archives.
    (Old paid invoices the system never observed as open will not appear
    in Notion. To seed history, run with --backfill-paid once.)

Customer relation:
  Each Notion page links to a customer page via the Customer relation.

  Resolving the parent customer:
    QBO Invoice.CustomerRef.name returns just the immediate customer's
    display name — for sub-customers (projects like 'RP7038-FTW') it's
    NOT the parent path. We pull the full Customer entity hierarchy at
    sync start and walk ParentRef chains to resolve each invoice's
    customer to its top-level parent (the actual GC / builder).

  Matching to Notion customer list:
    Names drift between QBO and Notion ("Riseland Homes LLC" vs
    "Riseland Homes"). Exact match would fail in those cases, leading
    users to create duplicate customer entries. Instead: normalize both
    sides (strip "LLC", "Inc", punctuation, etc.), tokenize, and find
    the Notion customer with the highest keyword overlap above a
    confidence threshold.

  On miss:
    Customer (raw) text still populates so the invoice imports cleanly;
    sync logs a warning so the user can add the customer to the Notion
    list. Next sync auto-links.

Per-invoice errors are caught, logged, counted. One bad row never stops
the run. If ANY row errors, the caller treats the run as partial.
"""
from __future__ import annotations

import datetime as dt
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from notion_client import NotionClient, NotionError
import qbo_client
from state import StateStore
from teams_notify import notify_invoice_event


# State key for the Change Data Capture deletion pass (last clean check time).
CDC_FLOW = "invoice_cdc_deletions"


log = logging.getLogger("automation_worker.invoice_sync")

FLOW_NAME = "qbo_invoices_to_notion"


# ───────────────── parsing constants ─────────────────

# Matches MFD###, CP###, or RP#### (with optional -FTW suffix), case-insensitive.
# Same shape as bill-tracker/qbo_bill_tracker.py.
PROJECT_NUM_RE = re.compile(r"\b((?:MFD|CP|RP)\d+(?:-FTW)?)\b", re.IGNORECASE)

# Matches a leading project number prefix in a memo string, like
# "CP582 - 1350 E. Miller Rd" or "MFD192-FTW - Foundation". Used to strip
# the internal project tag from Memo before it reaches a customer-facing
# email (the project # is already its own column in Notion).
LEADING_PROJECT_RE = re.compile(
    r"^(?:MFD|CP|RP)\d+(?:-FTW)?\s*[-–—]\s*",
    re.IGNORECASE,
)


def _clean_memo(memo: str) -> str:
    """Strip leading project # prefix from memo, if present."""
    if not memo:
        return ""
    return LEADING_PROJECT_RE.sub("", memo, count=1).strip()


# ───────────────── division → DB routing ─────────────────

DIVISION_MFD = "MFD"
DIVISION_RP = "RP"
DIVISION_CP = "CP"

# Notion Status option names (must match the Status select options in both DBs).
STATUS_UNPAID = "Unpaid"
STATUS_PARTIALLY_PAID = "Partially Paid"
STATUS_PAID = "Paid"

# Aging Bucket option names (must match the Aging Bucket select in both DBs).
AGING_CURRENT = "Current"
AGING_1_30 = "1-30"
AGING_31_60 = "31-60"
AGING_61_90 = "61-90"
AGING_90_PLUS = "90+"


# ───────────────── data shapes ─────────────────

@dataclass
class InvoiceRecord:
    """Parsed QBO invoice ready for Notion upsert."""
    qbo_id: str                  # match key
    invoice_num: str             # display title (DocNumber)
    project_num: str             # MFD118, RP4521, etc.
    division: str                # MFD / RP / CP
    parent_customer: str         # resolved from Customer hierarchy walk
    customer_raw: str            # full QBO CustomerRef.name (audit fallback)
    txn_date: Optional[dt.date]
    due_date: Optional[dt.date]
    total_amt: float
    balance: float
    memo: str                    # QBO PrivateNote — internal memo, hidden from customer statement
    status: str                  # one of STATUS_*
    net_terms: str               # QBO SalesTermRef resolved name ("Net 30", "Net 45", etc.) or empty
    line_items: List[Dict[str, Any]] = field(default_factory=list)  # positive lines: {description, amount}


@dataclass
class InvoiceSyncSummary:
    db_label: str
    invoices_seen: int = 0
    created: int = 0
    updated: int = 0
    flipped_to_paid: int = 0       # Notion was open, QBO no longer reports it open
    skipped_no_project: int = 0
    customer_match_failures: int = 0
    archived_paid_old: int = 0
    archived_deleted: int = 0      # QBO-confirmed deleted → removed from Notion + Excel
    errors: int = 0
    error_examples: List[str] = field(default_factory=list)

    @property
    def had_errors(self) -> bool:
        return self.errors > 0

    def as_dict(self) -> dict:
        return {
            "db": self.db_label,
            "seen": self.invoices_seen,
            "created": self.created,
            "updated": self.updated,
            "flipped_to_paid": self.flipped_to_paid,
            "skipped_no_project": self.skipped_no_project,
            "customer_unmatched": self.customer_match_failures,
            "archived_paid_old": self.archived_paid_old,
            "archived_deleted": self.archived_deleted,
            "errors": self.errors,
        }


# ───────────────── parsing ─────────────────

def _parse_iso_date(s: Any) -> Optional[dt.date]:
    """Parse 'YYYY-MM-DD' from QBO. None on missing/invalid."""
    if not s:
        return None
    try:
        return dt.date.fromisoformat(str(s)[:10])
    except ValueError:
        return None


def _extract_project_num(customer_ref_name: str) -> Optional[str]:
    """Pull MFD###/CP###/RP#### from QBO CustomerRef.name. Searches anywhere."""
    if not customer_ref_name:
        return None
    m = PROJECT_NUM_RE.search(customer_ref_name)
    return m.group(1).upper() if m else None


def _extract_division(project_num: Optional[str]) -> Optional[str]:
    if not project_num:
        return None
    p = project_num.upper()
    if p.startswith("MFD"):
        return DIVISION_MFD
    if p.startswith("CP"):
        return DIVISION_CP
    if p.startswith("RP"):
        return DIVISION_RP
    return None


def _resolve_parent_customer(
    qbo_inv: dict,
    customer_hierarchy: Dict[str, str],
) -> str:
    """
    Resolve the invoice's customer to its TOP-LEVEL parent display name.

    QBO Invoice.CustomerRef.name returns the immediate customer's display
    name. For sub-customers (project-level customers like 'RP7038-FTW'),
    that's just the project string — not the actual GC / builder name.

    customer_hierarchy is the {customer_id → root_parent_name} map built
    once per sync run by qbo_client.fetch_customer_hierarchy().
    """
    customer_ref = qbo_inv.get("CustomerRef") or {}
    cust_id = customer_ref.get("value")
    if cust_id:
        resolved = customer_hierarchy.get(str(cust_id))
        if resolved:
            return resolved
    # Fallback: pre-colon part of CustomerRef.name (rarely needed once
    # hierarchy is in place, but handles edge cases like missing/deleted
    # customers in the hierarchy map).
    name = customer_ref.get("name") or ""
    return name.split(":", 1)[0].strip() if ":" in name else name.strip()


def _compute_status(balance: float, total_amt: float) -> str:
    """Map Balance vs Total → Unpaid / Partially Paid / Paid."""
    # QBO returns floats; tolerate cent-level rounding noise.
    if balance <= 0.005:
        return STATUS_PAID
    if balance + 0.005 >= total_amt:
        return STATUS_UNPAID
    return STATUS_PARTIALLY_PAID


def _compute_aging_bucket(due_date: Optional[dt.date]) -> Optional[str]:
    """
    Map Due Date → aging bucket string. Returns None when due date is missing
    (Notion will leave the Aging Bucket select empty).

    Days to Due = due_date - today.
      >= 0   → Current (not yet due)
      -1..-30 → 1-30 days overdue
      -31..-60 → 31-60
      -61..-90 → 61-90
      < -90  → 90+
    """
    if due_date is None:
        return None
    days_to_due = (due_date - dt.date.today()).days
    if days_to_due >= 0:
        return AGING_CURRENT
    if days_to_due >= -30:
        return AGING_1_30
    if days_to_due >= -60:
        return AGING_31_60
    if days_to_due >= -90:
        return AGING_61_90
    return AGING_90_PLUS


def _positive_line_items(qbo_inv: dict) -> List[Dict[str, Any]]:
    """
    Return the invoice's POSITIVE line items as [{description, amount}, …].

    For MFD Teams cards we show what was billed this draw (e.g. the draw
    description, or "Retainage Billed") with a +$ amount. Negative lines
    (prior-billing offsets, retainage held back, deductions) are excluded by
    design, as are subtotal lines and zero/blank amounts.
    """
    out: List[Dict[str, Any]] = []
    for line in qbo_inv.get("Line") or []:
        if line.get("DetailType") == "SubTotalLineDetail":
            continue
        try:
            amt = float(line.get("Amount"))
        except (TypeError, ValueError):
            continue
        if amt <= 0:
            continue
        desc = (line.get("Description") or "").strip()
        if not desc:
            sid = line.get("SalesItemLineDetail") or {}
            desc = ((sid.get("ItemRef") or {}).get("name") or "Line item").strip()
        out.append({"description": desc, "amount": amt})
    return out


def _parse_invoice(
    qbo_inv: dict,
    customer_hierarchy: Dict[str, str],
    term_map: Dict[str, str],
) -> Optional[InvoiceRecord]:
    """
    Convert raw QBO Invoice JSON into our InvoiceRecord. Returns None and
    logs if the invoice can't be routed (missing project #).
    """
    qbo_id = str(qbo_inv.get("Id") or "")
    if not qbo_id:
        return None

    customer_ref = qbo_inv.get("CustomerRef") or {}
    customer_raw = customer_ref.get("name") or ""
    private_note = qbo_inv.get("PrivateNote") or ""

    # Look in CustomerRef.name first (the standard place — sub-customer projects).
    # Fall back to PrivateNote for legacy invoices that pre-date QBO Projects
    # adoption and only have the project # in the memo.
    project_num = _extract_project_num(customer_raw) or _extract_project_num(private_note)
    division = _extract_division(project_num)
    if not project_num or not division:
        return None  # caller logs as skipped_no_project

    total_amt = float(qbo_inv.get("TotalAmt") or 0.0)
    balance = float(qbo_inv.get("Balance") or 0.0)

    sales_term_ref = qbo_inv.get("SalesTermRef") or {}
    term_id = sales_term_ref.get("value")
    net_terms = term_map.get(str(term_id), "") if term_id else ""

    return InvoiceRecord(
        qbo_id=qbo_id,
        invoice_num=str(qbo_inv.get("DocNumber") or qbo_id),
        project_num=project_num,
        division=division,
        parent_customer=_resolve_parent_customer(qbo_inv, customer_hierarchy),
        customer_raw=customer_raw,
        txn_date=_parse_iso_date(qbo_inv.get("TxnDate")),
        due_date=_parse_iso_date(qbo_inv.get("DueDate")),
        total_amt=total_amt,
        balance=balance,
        status=_compute_status(balance, total_amt),
        memo=_clean_memo(qbo_inv.get("PrivateNote") or ""),
        net_terms=net_terms,
        line_items=_positive_line_items(qbo_inv),
    )


# ───────────────── customer cache + fuzzy matching ─────────────────

# Business-entity suffixes stripped during name normalization. Removing
# these means "Riseland Homes LLC" matches "Riseland Homes". Order matters:
# longer / more-specific suffixes first.
_BUSINESS_SUFFIXES = (
    " l l c", " l.l.c.", " llc",
    " incorporated", " inc",
    " limited", " ltd",
    " corporation", " corp",
    " company", " co",
    " group",
    " holdings",
    " enterprises",
)

# Tokens removed from the keyword set because they're too generic to be
# distinguishing on their own. "Riseland Homes" vs "Anderson Homes" should
# NOT match just because both end in "homes".
_STOPWORDS = frozenset({
    "the", "a", "an", "of", "and", "&",
    # Industry / business-descriptor words
    "homes", "home", "builders", "builder", "build", "building", "buildings",
    "construction", "constructions", "contractor", "contractors", "contracting",
    "developments", "development", "developers", "developer",
    "properties", "property",
    "investments", "investment",
    "real", "estate",
    "residential", "commercial", "industrial",
    "design", "designs", "designer", "designers",
    "studios", "studio",
    "services", "service",
    "solutions",
    "ventures", "venture",
    "associates",
    "partners", "partnership",
    "customs", "custom",  # generic in construction/automotive customer names
    "group", "groups",
    "concrete",  # we ARE the concrete company; GC names shouldn't lean on it
})


def _normalize_name(name: str) -> str:
    """Lowercase, strip punctuation, strip business suffixes, collapse spaces."""
    s = (name or "").lower().strip()
    if not s:
        return ""
    # Replace any non-alnum with space first — this turns "Inc." into "inc"
    # so the suffix matcher below can catch it.
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    # Pad with spaces so suffix replace doesn't merge with surrounding tokens.
    padded = f" {s} "
    for sfx in _BUSINESS_SUFFIXES:
        padded = padded.replace(f"{sfx} ", " ")
    s = re.sub(r"\s+", " ", padded).strip()
    return s


def _keywords(name: str) -> frozenset:
    """Distinctive tokens used for matching. Stopwords + 1-char tokens dropped."""
    norm = _normalize_name(name)
    if not norm:
        return frozenset()
    tokens = [t for t in norm.split() if len(t) > 1 and t not in _STOPWORDS]
    return frozenset(tokens)


def _generic_tokens(name: str) -> frozenset:
    """
    All meaningful tokens KEEPING stopwords (suffixes stripped, 1-char dropped).
    Used only for the all-stopword fallback: some GC names are made entirely of
    generic words (e.g. "Development & Construction Services"), so their
    distinctive keyword set is empty and normal matching can't see them.
    """
    norm = _normalize_name(name)
    if not norm:
        return frozenset()
    return frozenset(t for t in norm.split() if len(t) > 1)


def _tokens_near(a: str, b: str) -> bool:
    """
    True if two tokens look like the same word with a plural-S, single-char
    typo, or minor spelling drift. Used as a fallback inside the matcher when
    the exact-keyword path finds NO match — never overrides exact matches.

    Guards:
      - Both tokens must be ≥ 5 chars (avoids short-token false positives
        like "land" being a substring of "lakeland").
      - Either token must be a prefix of the other (plural-S, singular,
        possessive 's), OR string similarity must be ≥ 0.85 (single-char
        typos like "richmnd"/"richmond" or "alford"/"allford").

    Catches:
      'richmond'   ↔ 'richmonds'    (prefix — plural-S, the case that
                                     burned us with 'Richmonds Builders, LLC')
      'alford'     ↔ 'allford'      (single-char insertion typo)
    Does NOT match:
      'alford'     ↔ 'stafford'     (ratio ~0.71)
      'lonestar'   ↔ 'lone'         (one token < 5 chars — too generic)
      'green'      ↔ 'red'          (different roots)
    """
    if not a or not b:
        return False
    if a == b:
        return True
    if len(a) < 5 or len(b) < 5:
        return False
    if a.startswith(b) or b.startswith(a):
        return True
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio() >= 0.85


def _near_set_equal(a: frozenset, b: frozenset) -> bool:
    """
    True if two token sets are the same up to plural-S / minor spelling drift —
    every token on each side has an exact-or-near counterpart on the other.
    Requires ≥ 2 tokens per side so a single generic word ("construction") can't
    bridge two different companies. Used ONLY for the all-stopword fallback.
    """
    if len(a) < 2 or len(b) < 2:
        return False

    def _covered(x: str, ys: frozenset) -> bool:
        return any(x == y or _tokens_near(x, y) for y in ys)

    return all(_covered(x, b) for x in a) and all(_covered(y, a) for y in b)


def _compressed_form(name: str) -> str:
    """
    Compressed match form — alphanumerics only, joined, lowercase, business
    suffixes stripped. Used by the matcher as an exact-match pre-check, BEFORE
    the keyword/jaccard fuzzy logic runs.

    Catches the class of "same name, different spacing/punctuation" cases that
    keyword matching trips on:
      - 'LONESTAR GREEN HOMES'  ↔  'LONE STAR GREEN HOMES'  → both → 'lonestargreenhomes'
      - 'R.A. RAMOS …'          ↔  'RA Ramos …'             → identical compressed form
      - 'Richmond Builders, INC' ↔ 'Richmond Builders, LLC' → both → 'richmondbuilders'

    Only fires on EXACT compressed match. If no entry matches, falls through to
    the original fuzzy logic unchanged — so existing fuzzy matches aren't disturbed.
    """
    norm = _normalize_name(name)  # already lowercase + suffix-stripped
    return re.sub(r"[^a-z0-9]", "", norm)


@dataclass
class CustomerCacheEntry:
    page_id: str
    raw_name: str
    keywords: frozenset
    compressed: str = ""   # compressed alphanumeric form for exact-match pre-check


def _build_customer_cache(
    notion: NotionClient,
    customer_ds_id: str,
    title_prop: str = "Client",
) -> List[CustomerCacheEntry]:
    """
    Preload all customer pages from a customer DB into a list of cache
    entries. Each entry holds the page id, original display name, and the
    distinctive keyword set used for fuzzy matching.
    """
    entries: List[CustomerCacheEntry] = []
    for page in notion.query_data_source(customer_ds_id, page_size=100):
        page_id = page.get("id")
        if not page_id:
            continue
        title_prop_data = (page.get("properties") or {}).get(title_prop) or {}
        title_arr = title_prop_data.get("title") or []
        if not title_arr:
            continue
        # Notion title is a list of rich text blocks; concat plain_text.
        name = "".join(t.get("plain_text", "") for t in title_arr).strip()
        if not name:
            continue
        entries.append(CustomerCacheEntry(
            page_id=page_id,
            raw_name=name,
            keywords=_keywords(name),
            compressed=_compressed_form(name),
        ))
    return entries


def _lookup_customer_id(
    qbo_name: str,
    cache: List[CustomerCacheEntry],
    *,
    min_overlap: int = 1,
) -> Optional["CustomerMatch"]:
    """
    Find the best Notion customer match for a QBO parent customer name.

    Strategy: normalize both sides, extract distinctive keyword sets,
    pick the cache entry with the largest intersection. Returns None
    if no entry shares at least `min_overlap` distinctive keywords.

    Tie-breaker: highest Jaccard similarity. Returns CustomerMatch with
    the page id, the matched raw name, and the score so callers can log
    fuzzy matches for review.

    Pre-check: before running fuzzy logic, try an exact compressed-form match
    (alphanumerics joined, suffixes stripped). If a cache entry's compressed
    form equals QBO's, that's treated as an exact match — covers space and
    punctuation differences that the keyword tokenizer can't see through
    (e.g., LONESTAR ↔ LONE STAR, R.A. RAMOS ↔ RA Ramos). Pre-check ONLY fires
    on exact equality; otherwise the original fuzzy logic runs unchanged.
    """
    qbo_compressed = _compressed_form(qbo_name)
    qbo_kw = _keywords(qbo_name)

    # Compressed-form exact-match pre-check
    if qbo_compressed:
        for entry in cache:
            if entry.compressed and entry.compressed == qbo_compressed:
                # Real keyword overlap for logging consistency (may be 0 if the
                # name consists entirely of stopwords — pre-check still wins).
                overlap = len(qbo_kw & entry.keywords) if qbo_kw else 0
                return CustomerMatch(
                    page_id=entry.page_id,
                    matched_name=entry.raw_name,
                    overlap=overlap,
                    jaccard=1.0,   # marks is_exact=True → no "Fuzzy-matched" log
                )

    if not qbo_kw:
        # All-stopword name (e.g. "Development & Construction Services LLC"):
        # no distinctive keywords, so the passes below can't see it. Match it
        # ONLY against other all-stopword entries whose generic word set is
        # near-equal (same words up to plural/typo drift). Narrow by design —
        # a distinctively-named customer (non-empty keywords) is never a
        # candidate here, so generic words can't bleed a wrong match onto it.
        qbo_generic = _generic_tokens(qbo_name)
        for entry in cache:
            if entry.keywords:
                continue  # distinctively-named — not an all-stopword candidate
            if _near_set_equal(qbo_generic, _generic_tokens(entry.raw_name)):
                return CustomerMatch(
                    page_id=entry.page_id,
                    matched_name=entry.raw_name,
                    overlap=len(qbo_generic),
                    jaccard=0.9,   # < 1.0 → logs as a fuzzy match for review
                )
        return None
    best: Optional[CustomerMatch] = None
    for entry in cache:
        if not entry.keywords:
            continue
        overlap = len(qbo_kw & entry.keywords)
        if overlap < min_overlap:
            continue
        union = len(qbo_kw | entry.keywords)
        jaccard = overlap / union if union else 0.0
        if best is None or overlap > best.overlap or (
            overlap == best.overlap and jaccard > best.jaccard
        ):
            best = CustomerMatch(
                page_id=entry.page_id,
                matched_name=entry.raw_name,
                overlap=overlap,
                jaccard=jaccard,
            )

    # Near-token fallback — runs ONLY when the exact-keyword pass found nothing.
    # Catches plural-S and minor-spelling-drift cases that the strict set
    # intersection misses (the bug that was hitting 'Richmond' vs 'Richmonds').
    # Never overrides an exact match; never broadens an existing match.
    if best is None:
        for entry in cache:
            if not entry.keywords:
                continue
            # Count QBO tokens that have an exact OR near match in the entry.
            near = 0
            for q in qbo_kw:
                if q in entry.keywords or any(_tokens_near(q, e) for e in entry.keywords):
                    near += 1
            if near < min_overlap:
                continue
            union = len(qbo_kw | entry.keywords)
            jaccard = near / union if union else 0.0
            if best is None or near > best.overlap or (
                near == best.overlap and jaccard > best.jaccard
            ):
                best = CustomerMatch(
                    page_id=entry.page_id,
                    matched_name=entry.raw_name,
                    overlap=near,
                    jaccard=jaccard,
                )
    return best


@dataclass
class CustomerMatch:
    page_id: str
    matched_name: str
    overlap: int        # number of shared distinctive keywords
    jaccard: float      # |A∩B| / |A∪B|

    @property
    def is_exact(self) -> bool:
        # Jaccard of 1.0 → identical keyword sets after normalization.
        return self.jaccard >= 0.999


# ───────────────── Notion property builders ─────────────────

def _ts_prop_rich_text(value: str) -> dict:
    return {"rich_text": [{"text": {"content": value or ""}}]}


def _ts_prop_title(value: str) -> dict:
    return {"title": [{"text": {"content": value or ""}}]}


def _ts_prop_date(value: Optional[dt.date]) -> dict:
    if value is None:
        return {"date": None}
    return {"date": {"start": value.isoformat()}}


def _ts_prop_number(value: float) -> dict:
    return {"number": value}


def _ts_prop_select(value: Optional[str]) -> dict:
    if not value:
        return {"select": None}
    return {"select": {"name": value}}


def _ts_prop_relation(page_id: Optional[str]) -> dict:
    if not page_id:
        return {"relation": []}
    return {"relation": [{"id": page_id}]}


def _build_synced_properties(
    rec: InvoiceRecord,
    customer_page_id: Optional[str],
    include_division: bool,
) -> Dict[str, Any]:
    """
    Build the dict of Notion properties for the QBO-sourced fields ONLY.
    Excludes all human-owned fields (Quick Status, Next Follow-Up, Owner,
    Last Action Date) — those are never touched by sync.
    """
    today_iso = dt.date.today().isoformat()
    aging_bucket = _compute_aging_bucket(rec.due_date)
    qbo_link = f"https://app.qbo.intuit.com/app/invoice?txnId={rec.qbo_id}"
    props: Dict[str, Any] = {
        "Invoice #": _ts_prop_title(rec.invoice_num),
        "Invoice ID": _ts_prop_rich_text(rec.qbo_id),
        "Project #": _ts_prop_rich_text(rec.project_num),
        "Customer (raw)": _ts_prop_rich_text(rec.customer_raw),
        "Date": _ts_prop_date(rec.txn_date),
        "Due Date": _ts_prop_date(rec.due_date),
        "Total Amount": _ts_prop_number(rec.total_amt),
        "Open balance": _ts_prop_number(rec.balance),
        "Status": _ts_prop_select(rec.status),
        "Aging Bucket": _ts_prop_select(aging_bucket),
        "Memo": _ts_prop_rich_text(rec.memo),
        "QBO Link": {"url": qbo_link},
        "Last Synced": {"date": {"start": today_iso}},
    }
    # Net Terms: write only when QBO has a term value. Preserves manual edits
    # if QBO invoice has no term. Notion auto-creates new select options on
    # write, so non-standard term names from QBO (e.g., custom terms) don't
    # require pre-defining options in Notion.
    if rec.net_terms:
        props["Net Terms"] = _ts_prop_select(rec.net_terms)
    # Customer relation: ONLY write when sync found a match. If no match,
    # leave the field alone — preserves any manual link the user added in
    # Notion to fix a missing customer (Notion's PATCH semantics: keys
    # absent from the request are left untouched).
    if customer_page_id:
        props["Customer"] = _ts_prop_relation(customer_page_id)
    if include_division:
        props["Division"] = _ts_prop_select(rec.division)
    return props


# ───────────────── invoice page cache ─────────────────

@dataclass
class InvoiceCacheEntry:
    """
    Snapshot of an existing Notion invoice page used to detect Status transitions
    on the next sync. Captures the page id plus the prior values of Status and
    Open balance so the upsert can compare new vs. old and fire payment-event
    Teams notifications when something actually changed.
    """
    page_id: str
    prior_status: str         # "" for new invoices not yet in Notion
    prior_open_balance: float
    invoice_num: str = ""     # human DocNumber (e.g. "34075") for readable logs


def _build_invoice_cache(
    notion: NotionClient,
    ds_id: str,
) -> Dict[str, InvoiceCacheEntry]:
    """
    Pre-fetch all existing invoice pages from a DB and return
    {Invoice ID → InvoiceCacheEntry}. One paginated query per DB instead of
    one per QBO invoice — keeps us under Notion's rate limit.

    The InvoiceCacheEntry captures prior Status + Open balance so the upsert
    step can detect transitions (Unpaid → Partially Paid, Unpaid → Paid, etc.)
    and route Teams payment notifications accordingly.
    """
    cache: Dict[str, InvoiceCacheEntry] = {}
    for page in notion.query_data_source(ds_id, page_size=100):
        page_id = page.get("id")
        if not page_id:
            continue
        props = page.get("properties") or {}
        inv_id_prop = props.get("Invoice ID") or {}
        rich = inv_id_prop.get("rich_text") or []
        if not rich:
            continue
        inv_id = "".join(t.get("plain_text", "") for t in rich).strip()
        if not inv_id:
            continue
        prior_status = ((props.get("Status") or {}).get("select") or {}).get("name") or ""
        prior_open_balance = float((props.get("Open balance") or {}).get("number") or 0.0)
        inv_num_arr = (props.get("Invoice #") or {}).get("title") or []
        invoice_num = "".join(t.get("plain_text", "") for t in inv_num_arr).strip()
        cache[inv_id] = InvoiceCacheEntry(
            page_id=page_id,
            prior_status=prior_status,
            prior_open_balance=prior_open_balance,
            invoice_num=invoice_num,
        )
    return cache


# ───────────────── per-DB upsert pass ─────────────────

def _upsert_one(
    notion: NotionClient,
    ds_id: str,
    rec: InvoiceRecord,
    customer_page_id: Optional[str],
    include_division: bool,
    invoice_cache: Dict[str, InvoiceCacheEntry],
    summary: InvoiceSyncSummary,
    dry_run: bool,
    teams_webhook_url: str = "",
) -> None:
    props = _build_synced_properties(rec, customer_page_id, include_division)
    entry = invoice_cache.get(rec.qbo_id)
    existing_page_id = entry.page_id if entry else None
    prior_status = entry.prior_status if entry else ""
    prior_open_balance = entry.prior_open_balance if entry else 0.0

    if existing_page_id:
        if dry_run:
            log.info("[dry-run] UPDATE %s/%s → page=%s",
                     rec.division, rec.invoice_num, existing_page_id)
        else:
            notion.update_page(existing_page_id, props)
        summary.updated += 1
    else:
        if dry_run:
            log.info("[dry-run] CREATE %s/%s", rec.division, rec.invoice_num)
        else:
            created = notion.create_page(ds_id, props)
            # Track the new page in the cache so a duplicate QBO invoice
            # (shouldn't happen, but defensive) won't double-create.
            new_id = created.get("id")
            if new_id:
                invoice_cache[rec.qbo_id] = InvoiceCacheEntry(
                    page_id=new_id,
                    prior_status=rec.status,
                    prior_open_balance=rec.balance,
                )
        summary.created += 1

    # Short-pay Teams notification (MFD only, Phase 1 scope).
    # Fires when an EXISTING Unpaid invoice transitions to Partially Paid —
    # i.e., customer made a partial payment since last sync. New invoices that
    # arrive partial-from-the-start don't fire (no prior state to compare).
    # Full-Paid transitions are handled by the flip-to-paid sweep, not here.
    if (
        not dry_run
        and existing_page_id
        and teams_webhook_url
        and rec.division == DIVISION_MFD
        and prior_status == STATUS_UNPAID
        and rec.status == STATUS_PARTIALLY_PAID
        and rec.balance < prior_open_balance
    ):
        partial_amount = max(prior_open_balance - rec.balance, 0.0)
        qbo_link = f"https://app.qbo.intuit.com/app/invoice?txnId={rec.qbo_id}"
        notify_invoice_event(
            teams_webhook_url,
            event_type="short_pay",
            division=rec.division,
            invoice_num=rec.invoice_num,
            customer=rec.parent_customer or "(unknown customer)",
            amount=rec.balance,
            project=rec.project_num,
            qbo_link=qbo_link,
            short_pay_amount=partial_amount,
            line_items=rec.line_items,
        )


# ───────────────── flip-to-paid sweep ─────────────────

def _flip_open_to_paid_when_qbo_no_longer_open(
    notion: NotionClient,
    ds_id: str,
    qbo_open_invoice_ids: set,
    payment_dates: Dict[str, str],
    summary: InvoiceSyncSummary,
    dry_run: bool,
    qbo_creds=None,
    customer_hierarchy: Optional[Dict[str, str]] = None,
    teams_webhook_url: str = "",
    is_mfd: bool = False,
) -> None:
    """
    Find Notion invoice pages currently marked NOT Paid whose Invoice ID
    is no longer in the QBO open set, and resolve each one:

      - Has a QBO Payment behind it           → flip to Paid (normal collection).
      - No Payment, but QBO still has it       → flip to Paid (voided / zero-
                                                 balance / written-off; stays on
                                                 file, treated as closed).
      - No Payment AND QBO confirms it's gone  → DELETED → archive the Notion
                                                 page so it disappears from Notion
                                                 views and the Excel mirror.

    The delete check (qbo_client.invoice_exists) is only consulted for the
    no-Payment case, and we archive ONLY when QBO positively confirms the
    invoice is gone. If the existence check can't be made (API error) or QBO
    still returns the invoice, we never archive — uncertainty is left for the
    next run. `qbo_creds` is required to run that check; if None, the delete
    path is skipped and behavior falls back to the original flip-to-Paid.

    Teams notification (MFD only, Phase 1): when this sweep flips an MFD
    invoice to Paid, fire a notification to the configured Workflow webhook.
    Best-effort — notifier swallows failures so the flip still records.
    """
    # Filter: Status is "Unpaid" or "Partially Paid" — i.e., not Paid.
    # Notion doesn't have "not equals" for select cleanly, so use OR of
    # the open statuses.
    filter_body = {
        "or": [
            {"property": "Status", "select": {"equals": STATUS_UNPAID}},
            {"property": "Status", "select": {"equals": STATUS_PARTIALLY_PAID}},
        ]
    }
    for page in notion.query_data_source(ds_id, filter_body=filter_body, page_size=100):
        page_id = page.get("id")
        if not page_id:
            continue
        props = page.get("properties") or {}
        inv_id_prop = props.get("Invoice ID") or {}
        rich = inv_id_prop.get("rich_text") or []
        inv_id = "".join(t.get("plain_text", "") for t in rich).strip()
        if not inv_id:
            # No QBO Invoice ID set — can't reconcile. Probably a manual
            # entry; leave alone.
            continue
        if inv_id in qbo_open_invoice_ids:
            continue  # still open in QBO, no flip

        # No longer open in QBO. Distinguish a genuine closure (paid / voided /
        # zero-balance — invoice still on file) from a DELETION (invoice removed
        # entirely, e.g. created by accident and deleted).
        has_payment = inv_id in payment_dates
        if not has_payment and qbo_creds is not None:
            # No Payment object behind this closure. Could be deleted, voided,
            # written off, or paid before the payment-lookback window. Read back
            # from QBO to find out which. Only a POSITIVE "gone" answer archives.
            try:
                still_exists = qbo_client.invoice_exists(qbo_creds, inv_id)
            except Exception as e:
                log.warning(
                    "Could not verify QBO existence for invoice %s — leaving "
                    "as-is this run (no flip, no archive): %s", inv_id, e
                )
                continue
            if not still_exists:
                # DELETED in QBO → remove from Notion (archive → Trash, restorable
                # 30 days) so it drops out of Notion views AND the Excel mirror.
                if dry_run:
                    log.info("[dry-run] ARCHIVE-DELETED inv=%s page=%s "
                             "(deleted in QBO → would remove from Notion + Excel)",
                             inv_id, page_id)
                    summary.archived_deleted += 1
                    continue
                try:
                    notion.archive_page(page_id)
                except NotionError as e:
                    log.warning("Failed to archive deleted invoice %s: %s", page_id, e)
                    summary.errors += 1
                    if len(summary.error_examples) < 5:
                        summary.error_examples.append(f"archive-deleted {page_id}: {e}")
                    continue
                log.info("Archived invoice %s — deleted in QBO; removed from "
                         "Notion + Excel.", inv_id)
                summary.archived_deleted += 1
                continue
            # else: still exists in QBO (voided / zero-balance) → fall through
            # and close it as Paid below, same as before.

        # Genuine closure → flip to Paid. Use QBO's actual payment date when
        # available, else today (fallback for voided/written-off invoices that
        # drop out of open without a Payment object).
        today_iso = dt.date.today().isoformat()
        paid_date_iso = payment_dates.get(inv_id) or today_iso
        if dry_run:
            log.info("[dry-run] FLIP-TO-PAID inv=%s page=%s paid_date=%s",
                     inv_id, page_id, paid_date_iso)
        else:
            try:
                # Paid Date is written ONLY at the moment of the flip — once
                # the page is Status=Paid, the filter above excludes it from
                # subsequent sweeps so this value won't be overwritten.
                # If a user manually flipped Paid back to Unpaid then it gets
                # re-flipped here, that's an acceptable corner case.
                notion.update_page(page_id, {
                    "Status": _ts_prop_select(STATUS_PAID),
                    "Open balance": _ts_prop_number(0.0),
                    "Paid Date": {"date": {"start": paid_date_iso}},
                    "Last Synced": {"date": {"start": today_iso}},
                })
            except NotionError as e:
                log.warning("Failed to flip %s to Paid: %s", page_id, e)
                summary.errors += 1
                if len(summary.error_examples) < 5:
                    summary.error_examples.append(f"flip {page_id}: {e}")
                continue

            # Teams paid notification — MFD only in Phase 1. Read invoice
            # details off the page (we already have them from the query).
            if teams_webhook_url and is_mfd:
                def _title_text(prop: dict) -> str:
                    arr = (prop or {}).get("title") or []
                    return "".join(t.get("plain_text", "") for t in arr).strip()

                def _rich_text(prop: dict) -> str:
                    arr = (prop or {}).get("rich_text") or []
                    return "".join(t.get("plain_text", "") for t in arr).strip()

                inv_num = _title_text(props.get("Invoice #")) or inv_id
                # Fallback customer = the raw sub-customer (e.g. "MFD183"); the
                # resolved PARENT (GC / developer) below overrides it when we can
                # read the invoice from QBO.
                customer = _rich_text(props.get("Customer (raw)")) or "(unknown customer)"
                project = _rich_text(props.get("Project #")) or ""
                total_amt = float((props.get("Total Amount") or {}).get("number") or 0.0)
                qbo_link = (props.get("QBO Link") or {}).get("url") or \
                    f"https://app.qbo.intuit.com/app/invoice?txnId={inv_id}"

                # Pull the invoice from QBO for (a) its positive line items and
                # (b) the resolved PARENT customer (the GC / developer, not the
                # project-# sub-customer). Best-effort: a paid invoice still
                # exists in QBO, but if the fetch fails we keep the fallbacks.
                paid_line_items: List[Dict[str, Any]] = []
                if qbo_creds is not None:
                    try:
                        full_inv = qbo_client.fetch_invoice(qbo_creds, inv_id)
                        if full_inv:
                            paid_line_items = _positive_line_items(full_inv)
                            parent = _resolve_parent_customer(
                                full_inv, customer_hierarchy or {})
                            if parent:
                                customer = parent
                    except Exception as e:
                        log.warning("Teams paid: couldn't read QBO detail for "
                                    "invoice #%s: %s", inv_num, e)

                notify_invoice_event(
                    teams_webhook_url,
                    event_type="paid",
                    division=DIVISION_MFD,
                    invoice_num=inv_num,
                    customer=customer,
                    amount=total_amt,
                    project=project,
                    qbo_link=qbo_link,
                    line_items=paid_line_items,
                )
        summary.flipped_to_paid += 1


# ───────────────── CDC deletion pass ─────────────────

def _archive_cdc_deleted_invoices(
    notion: NotionClient,
    qbo_creds,
    caches: List[Tuple[Dict[str, "InvoiceCacheEntry"], "InvoiceSyncSummary", bool]],
    state: StateStore,
    dry_run: bool,
) -> None:
    """
    Catch invoices DELETED in QBO that the open-set sweep can't see — i.e. ones
    already marked Paid in Notion when they were deleted (the void→delete path).

    Uses QBO Change Data Capture: one call returns everything deleted since the
    last clean check. For each reported-deleted Invoice Id we (a) confirm it's
    really gone via invoice_exists (safety gate — never archive on a CDC quirk),
    then (b) archive the matching Notion page from the prebuilt invoice caches,
    regardless of its Notion status. Counts toward `archived_deleted`.

    `caches` is a list of (invoice_cache, summary, is_mfd) — one per DB. The
    changedSince timestamp comes from StateStore and is clamped to QBO's 30-day
    limit. State is advanced to "now" only on a fully clean pass (no archive
    errors), so a failed archive is retried next run.
    """
    now = dt.datetime.now(dt.timezone.utc)
    floor = now - dt.timedelta(days=30)  # QBO CDC won't accept older than 30 days
    last = state.get_last_run(CDC_FLOW)
    since_dt = floor
    if last:
        try:
            parsed = dt.datetime.strptime(last, "%Y-%m-%dT%H:%M:%SZ").replace(
                tzinfo=dt.timezone.utc
            )
            since_dt = max(parsed, floor)
        except ValueError:
            since_dt = floor
    changed_since = since_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

    try:
        deleted_ids = qbo_client.fetch_deleted_invoice_ids(qbo_creds, changed_since)
    except Exception as e:
        log.warning("CDC deletion check failed — skipped this run (state unchanged): %s", e)
        return

    had_error = False
    matched = 0       # deleted in QBO AND present in a Notion tracker
    not_tracked = 0   # deleted in QBO but never in Notion → nothing to archive

    for inv_id in deleted_ids:
        # Find the Notion page first — most QBO deletions are invoices we never
        # tracked (no project #, equipment leases, etc.), so skip those without
        # spending a QBO existence call.
        entry = summary = None
        for cache, db_summary, _is_mfd in caches:
            hit = cache.get(inv_id)
            if hit:
                entry, summary = hit, db_summary
                break
        if entry is None:
            not_tracked += 1
            continue

        matched += 1
        inv_label = entry.invoice_num or f"id {inv_id}"

        # Safety gate: confirm the invoice is actually gone before archiving.
        try:
            if qbo_client.invoice_exists(qbo_creds, inv_id):
                log.info("CDC: invoice #%s reported deleted but still exists in QBO "
                         "— skipped.", inv_label)
                continue
        except Exception as e:
            log.warning("CDC: could not confirm deletion of invoice #%s — skipped: %s",
                        inv_label, e)
            had_error = True
            continue

        if dry_run:
            log.info("[dry-run] CDC-ARCHIVE invoice #%s — deleted in QBO "
                     "(would remove from Notion + Excel).", inv_label)
            summary.archived_deleted += 1
        else:
            try:
                notion.archive_page(entry.page_id)
                summary.archived_deleted += 1
                log.info("CDC-ARCHIVE invoice #%s — deleted in QBO; removed from "
                         "Notion + Excel.", inv_label)
            except NotionError as e:
                log.warning("CDC archive failed for invoice #%s: %s", inv_label, e)
                summary.errors += 1
                had_error = True

    log.info("CDC: %d deleted in QBO since %s — %d matched a Notion invoice "
             "(%s), %d not tracked in Notion (skipped).",
             len(deleted_ids), changed_since, matched,
             "to archive" if dry_run else "archived", not_tracked)

    # Advance the watermark only on a clean pass, so a failed archive is retried.
    if not dry_run and not had_error:
        state.set_last_run(CDC_FLOW, now)


# ───────────────── 12-month cleanup ─────────────────

def _archive_paid_invoices_older_than(
    notion: NotionClient,
    ds_id: str,
    cutoff: dt.date,
    summary: InvoiceSyncSummary,
    dry_run: bool,
) -> None:
    """
    Soft-delete (archive) paid invoices whose TxnDate is older than `cutoff`.
    Uses a Notion-side filter so we only fetch the ones that actually qualify.
    """
    filter_body = {
        "and": [
            {"property": "Status", "select": {"equals": STATUS_PAID}},
            {"property": "Date", "date": {"before": cutoff.isoformat()}},
        ]
    }
    for page in notion.query_data_source(ds_id, filter_body=filter_body, page_size=100):
        page_id = page.get("id")
        if not page_id:
            continue
        if dry_run:
            log.info("[dry-run] ARCHIVE old-paid page=%s", page_id)
        else:
            try:
                notion.archive_page(page_id)
            except NotionError as e:
                log.warning("Failed to archive %s: %s", page_id, e)
                summary.errors += 1
                if len(summary.error_examples) < 5:
                    summary.error_examples.append(f"archive {page_id}: {e}")
                continue
        summary.archived_paid_old += 1


# ───────────────── main flow ─────────────────

@dataclass(frozen=True)
class InvoiceSyncTargets:
    """Notion data source IDs the sync writes to / reads customer relations from."""
    res_com_invoice_ds_id: str
    mfd_invoice_ds_id: str
    customer_list_ds_id: str       # for Res/Com customer lookup
    mfd_client_list_ds_id: str     # for MFD customer lookup


def sync_qbo_invoices_to_notion(
    *,
    notion: NotionClient,
    targets: InvoiceSyncTargets,
    paid_retention_months: int = 12,
    dry_run: bool = False,
    teams_webhook_mfd_paid: str = "",
    state_dir: Optional[Path] = None,
) -> Tuple[InvoiceSyncSummary, InvoiceSyncSummary]:
    """
    One full sync pass. Returns (res_com_summary, mfd_summary).

    Steps:
      1. Auth to QBO (one Touch ID)
      2. Pull all open invoices + paid invoices in the retention window
      3. Preload both Notion customer caches
      4. Per invoice: parse, route, upsert
      5. Cleanup: archive old paid invoices outside the retention window
    """
    res_com_summary = InvoiceSyncSummary(db_label="Res/Com")
    mfd_summary = InvoiceSyncSummary(db_label="MFD")

    # Teams Workflow webhook for MFD payment notifications (paid / short-pay).
    # Passed in from config (Keychain-first, env fallback — see
    # config._get_teams_webhook). Optional — empty string silently disables.
    # Phase 1 covers MFD only; CP and RP routes have webhook="" hardcoded below.
    teams_webhook_mfd_paid = (teams_webhook_mfd_paid or "").strip()
    if teams_webhook_mfd_paid:
        log.info("Teams MFD-paid webhook configured — paid/short-pay events will route to channel.")
    else:
        log.debug("TEAMS_WEBHOOK_MFD_PAID not set — Teams notifications disabled.")

    log.info("Authenticating to QBO…")
    qbo = qbo_client.load_qbo_credentials()

    log.info("Loading QBO customer hierarchy…")
    customer_hierarchy = qbo_client.fetch_customer_hierarchy(qbo)
    log.info("Customer hierarchy loaded: %d QBO customers indexed", len(customer_hierarchy))

    log.info("Loading QBO term map…")
    term_map = qbo_client.fetch_term_map(qbo)
    log.info("Term map loaded: %d QBO terms indexed", len(term_map))

    # Payment dates per invoice — used by flip-to-paid sweep to stamp the
    # actual QBO payment TxnDate (not "today") on Notion's Paid Date field.
    # Critical when the sync has been offline (Mac sleep, FDA outage, etc.)
    # and multiple payments accumulated in QBO during the gap.
    log.info("Loading QBO payment dates (last 6 months)…")
    payment_dates = qbo_client.fetch_payment_dates(qbo, lookback_months=6)
    log.info("Payment dates loaded: %d invoices have payments in window", len(payment_dates))

    cutoff = dt.date.today() - dt.timedelta(days=paid_retention_months * 30)
    # Open-only fetch from QBO. Paid invoices are detected as a delta
    # against Notion's open set in the flip-to-paid sweep below — this
    # cuts the QBO payload from ~1500 rows to ~150 and avoids hammering
    # Notion with paid invoices we already know about.
    log.info("Fetching QBO open invoices (Balance > 0)…")
    raw_invoices = qbo_client.query_all(qbo, "Invoice", where="Balance > '0'")
    qbo_open_invoice_ids = {
        str(inv.get("Id")) for inv in raw_invoices if inv.get("Id")
    }
    log.info("QBO returned %d open invoices", len(raw_invoices))

    log.info("Loading Notion customer caches…")
    res_com_cache = _build_customer_cache(notion, targets.customer_list_ds_id)
    mfd_cache = _build_customer_cache(notion, targets.mfd_client_list_ds_id)
    log.info(
        "Customer cache loaded: %d Res/Com customers, %d MFD clients",
        len(res_com_cache), len(mfd_cache),
    )

    log.info("Loading existing Notion invoice pages…")
    res_com_invoice_cache = _build_invoice_cache(notion, targets.res_com_invoice_ds_id)
    mfd_invoice_cache = _build_invoice_cache(notion, targets.mfd_invoice_ds_id)
    log.info(
        "Invoice page cache loaded: %d Res/Com invoices, %d MFD invoices already in Notion",
        len(res_com_invoice_cache), len(mfd_invoice_cache),
    )

    total_invoices = len(raw_invoices)
    PROGRESS_EVERY = 25  # log a progress line every N invoices
    for idx, raw in enumerate(raw_invoices, start=1):
        if idx == 1 or idx == total_invoices or idx % PROGRESS_EVERY == 0:
            log.info(
                "Progress: %d/%d invoices processed (Res/Com: %d seen, %d unmatched | MFD: %d seen, %d unmatched)",
                idx, total_invoices,
                res_com_summary.invoices_seen, res_com_summary.customer_match_failures,
                mfd_summary.invoices_seen, mfd_summary.customer_match_failures,
            )
        try:
            rec = _parse_invoice(raw, customer_hierarchy, term_map)
            if rec is None:
                # Could be a no-DocNumber row, or one we can't route.
                # Distinguish: if there's no project #, count as no-project.
                customer_raw = (raw.get("CustomerRef") or {}).get("name") or ""
                if not _extract_project_num(customer_raw):
                    # Pick a summary to attribute it to — both, since we don't
                    # know division. Attribute to res_com by convention.
                    res_com_summary.skipped_no_project += 1
                    log.debug("Skipping invoice %s — no project # in '%s'",
                              raw.get("Id"), customer_raw)
                continue

            if rec.division == DIVISION_MFD:
                target_ds = targets.mfd_invoice_ds_id
                match = _lookup_customer_id(rec.parent_customer, mfd_cache)
                summary = mfd_summary
                include_division = False
                invoice_cache = mfd_invoice_cache
            else:
                target_ds = targets.res_com_invoice_ds_id
                match = _lookup_customer_id(rec.parent_customer, res_com_cache)
                summary = res_com_summary
                include_division = True
                invoice_cache = res_com_invoice_cache

            summary.invoices_seen += 1
            customer_id: Optional[str] = None
            if match is None:
                if rec.parent_customer:
                    summary.customer_match_failures += 1
                    log.warning(
                        "[%s] No Notion customer match for '%s' (invoice %s) — "
                        "Customer relation will be empty. Add to %s list and "
                        "the next sync will link.",
                        rec.division, rec.parent_customer, rec.invoice_num,
                        "MFD Client" if rec.division == DIVISION_MFD else "Customer",
                    )
            else:
                customer_id = match.page_id
                if not match.is_exact:
                    log.info(
                        "[%s] Fuzzy-matched '%s' → '%s' (overlap=%d, jaccard=%.2f) "
                        "for invoice %s",
                        rec.division, rec.parent_customer, match.matched_name,
                        match.overlap, match.jaccard, rec.invoice_num,
                    )

            _upsert_one(
                notion=notion,
                ds_id=target_ds,
                rec=rec,
                customer_page_id=customer_id,
                include_division=include_division,
                invoice_cache=invoice_cache,
                summary=summary,
                dry_run=dry_run,
                teams_webhook_url=teams_webhook_mfd_paid,
            )

        except (NotionError, Exception) as e:  # noqa: BLE001 — flow-level catch
            # Attribute to whichever summary we know about; default Res/Com.
            summary_to_blame = res_com_summary
            try:
                if rec is not None and rec.division == DIVISION_MFD:  # type: ignore[has-type]
                    summary_to_blame = mfd_summary
            except UnboundLocalError:
                pass
            summary_to_blame.errors += 1
            if len(summary_to_blame.error_examples) < 5:
                summary_to_blame.error_examples.append(
                    f"invoice {raw.get('Id')}: {type(e).__name__}: {e}"
                )
            log.exception("Per-invoice error on QBO Id=%s: %s", raw.get("Id"), e)

    # Flip-to-paid sweep — detect QBO→paid transitions by finding Notion
    # invoices marked open whose Invoice ID dropped out of the QBO open set.
    log.info("Sweep: flipping Notion-open invoices not in QBO open set → Paid…")
    _flip_open_to_paid_when_qbo_no_longer_open(
        notion=notion,
        ds_id=targets.res_com_invoice_ds_id,
        qbo_open_invoice_ids=qbo_open_invoice_ids,
        payment_dates=payment_dates,
        summary=res_com_summary,
        dry_run=dry_run,
        qbo_creds=qbo,
        customer_hierarchy=customer_hierarchy,
        teams_webhook_url="",   # Phase 1: RP/CP not in Teams notification scope
        is_mfd=False,
    )
    _flip_open_to_paid_when_qbo_no_longer_open(
        notion=notion,
        ds_id=targets.mfd_invoice_ds_id,
        qbo_open_invoice_ids=qbo_open_invoice_ids,
        payment_dates=payment_dates,
        summary=mfd_summary,
        dry_run=dry_run,
        qbo_creds=qbo,
        customer_hierarchy=customer_hierarchy,
        teams_webhook_url=teams_webhook_mfd_paid,
        is_mfd=True,
    )

    # CDC deletion pass — catch invoices deleted in QBO that the open-set sweep
    # can't see (already Paid in Notion when deleted, e.g. void→delete). One QBO
    # call returns what was deleted since the last clean run; each is confirmed
    # gone, then the matching Notion page is archived regardless of status.
    if state_dir is not None:
        log.info("CDC: checking for QBO-deleted invoices to archive…")
        cdc_state = StateStore(Path(state_dir) / "invoice_cdc_state.json")
        _archive_cdc_deleted_invoices(
            notion=notion,
            qbo_creds=qbo,
            caches=[
                (res_com_invoice_cache, res_com_summary, False),
                (mfd_invoice_cache, mfd_summary, True),
            ],
            state=cdc_state,
            dry_run=dry_run,
        )
    else:
        log.debug("CDC deletion pass skipped (no state_dir provided).")

    # Cleanup pass — archive old paid invoices (12-month retention).
    log.info("Cleanup: archiving paid invoices older than %s…", cutoff.isoformat())
    _archive_paid_invoices_older_than(
        notion=notion,
        ds_id=targets.res_com_invoice_ds_id,
        cutoff=cutoff,
        summary=res_com_summary,
        dry_run=dry_run,
    )
    _archive_paid_invoices_older_than(
        notion=notion,
        ds_id=targets.mfd_invoice_ds_id,
        cutoff=cutoff,
        summary=mfd_summary,
        dry_run=dry_run,
    )

    return res_com_summary, mfd_summary
