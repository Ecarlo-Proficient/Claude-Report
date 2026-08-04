#!/usr/bin/env python3
"""
proposals.py — read the contract price off a signed bid-proposal PDF.

Lives in `shared/` because two tools need it (repo rule: tools never import
tools). The CP WIP reader uses it for pre-draw jobs; the RP side has its own
older copy in `one-offs/rp_schedule_wip_preview.pdf_subtotal` that should fold
into this one when that tool is next touched.

Why the PDF and not the takeoff (the user 2026-08-04): the proposal PDF is the
SIGNED document — it is what the customer agreed to pay. The takeoff is an
internal working file, and its template ships with several proposal tabs, so it
can be ambiguous about which total is the contract. Order is therefore
draw → proposal PDF → takeoff.

Extraction is deliberately defensive: pdfplumber often breaks digits apart —
a total prints as "$ 1 23,456" with a stray space mid-number — so whitespace
inside a number is stripped before parsing.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Tuple

# "<Qualifier> TOTAL: $ 1 05,815.00" — the qualifier group tells a section
# total ("Foundation Total") from the overall one ("TOTAL"). Tabs/spaces are
# allowed INSIDE the number but never a newline, so a match can't run past the
# end of its line.
_TOTAL_CENTS = re.compile(
    r"(?:(?P<qual>[A-Za-z][A-Za-z&/ ]{0,24}?)[ \t]+)?"
    r"TOTAL[ \t]*:?[ \t]*\$?[ \t]*(?P<amt>\d[\d,\t ]*\.\d{2})", re.I)
_TOTAL_PLAIN = re.compile(
    r"(?:(?P<qual>[A-Za-z][A-Za-z&/ ]{0,24}?)[ \t]+)?"
    r"TOTAL[ \t]*:?[ \t]*\$[ \t]*(?P<amt>\d[\d,\t ]*)", re.I)

_SKIP_PDF = ("INVOICE", "DIAGRAM", "PLAN", "SUBMITTAL", "LIEN", "W-9", "COI")


def _amount(raw: str) -> Optional[float]:
    try:
        return float(re.sub(r"[,\s]", "", raw))
    except ValueError:
        return None


def pdf_text(path: Path) -> str:
    """All text in the PDF, or '' if it can't be read."""
    try:
        import pdfplumber
    except ImportError:
        return ""
    try:
        with pdfplumber.open(path) as pdf:
            return "\n".join(pg.extract_text() or "" for pg in pdf.pages)
    except Exception:
        return ""


def grand_total(path: Path) -> Tuple[Optional[float], str]:
    """The overall contract total in a proposal PDF → (amount, how_we_got_it).

    An UNQUALIFIED 'TOTAL:' is the overall figure; 'Foundation Total' and
    'Site Total' are section subtotals. When the unqualified total equals the
    sum of the qualified ones, that is stated in the note — a free cross-check
    that the number really is the whole proposal and not one page of it.
    """
    txt = pdf_text(path)
    if not txt:
        return None, "PDF unreadable (or pdfplumber missing)"
    hits: List[Tuple[Optional[str], float]] = []
    for rx in (_TOTAL_CENTS, _TOTAL_PLAIN):
        for m in rx.finditer(txt):
            amt = _amount(m.group("amt"))
            if amt:
                hits.append(((m.group("qual") or "").strip() or None, amt))
        if hits:
            break
    if not hits:
        return None, "no 'TOTAL: $…' line in the PDF"

    plain = [a for q, a in hits if not q]
    parts = [a for q, a in hits if q]
    if not plain:
        # ONLY section subtotals ('Eartwork Total', 'Foundation Total'). A
        # section is not the contract — returning the largest one silently cut
        # CP783's contract from $364k to $56k (2026-08-04). Refuse instead.
        return None, (f"PDF has only section totals "
                      f"({', '.join(q for q, _a in hits if q)}) — no overall "
                      f"'TOTAL:' line, so it isn't the whole contract")
    total = plain[-1]
    note = f"proposal PDF 'TOTAL:' ${total:,.0f}"
    if parts and abs(sum(parts) - total) < 1.0:
        note += f" (= sum of its {len(parts)} section totals)"
    return total, note


def find_proposal_pdf(folder: Path) -> Tuple[Optional[Path], str]:
    """The bid-proposal PDF in a project folder → (path, note).

    Anything that is plainly not a proposal (invoices, plans, submittals) is
    skipped. With several candidates the NEWEST by modified time wins — that is
    the revision the customer signed last — and the note says so, so a wrong
    pick is visible rather than silent.
    """
    if folder is None or not folder.exists():
        return None, "no project folder"
    cands = []
    for f in folder.iterdir():
        if f.suffix.lower() != ".pdf" or f.name.startswith("~$"):
            continue
        up = f.name.upper()
        if any(s in up for s in _SKIP_PDF):
            continue
        if "PROPOSAL" in up or "BID" in up:
            cands.append(f)
    if not cands:
        return None, "no proposal PDF in the folder"
    if len(cands) == 1:
        return cands[0], f"proposal PDF '{cands[0].name}'"
    # SEVERAL proposal PDFs — do NOT guess. They are usually different SCOPES,
    # not revisions of one price: CP783 carries a main proposal, a breakout, a
    # dirt/utilities proposal and a revised dirt proposal. Picking the newest
    # chose the $56k dirt scope over the $364k job. The caller falls back to
    # the takeoff and this note names the candidates.
    return None, ("several proposal PDFs, none can be assumed to be THE "
                  "contract: " + ", ".join(sorted(p.name for p in cands)[:4]))


def contract_from_folder(folder: Path) -> Tuple[Optional[float], str]:
    """Contract price from the folder's proposal PDF → (amount, note)."""
    pdf, note = find_proposal_pdf(folder)
    if pdf is None:
        return None, note
    amt, how = grand_total(pdf)
    return amt, f"{how} — {pdf.name}" if amt is not None else f"{how} ({pdf.name})"
