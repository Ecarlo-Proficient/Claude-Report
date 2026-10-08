"""statement_markup.py - a marked-up COPY of the vendor's statement (tool-local).

Owner 10/08/2026: "highlight each bill row with the color of its bucket so the
clerk can verify you grabbed everything; leave the statement in the folder
untouched, mark up the copy in the Excel".

Each page of the statement is rendered to an image; every line the parser read
is found on the page and its row gets a band in its bucket's colour. The check
runs both ways:
  * a line the parser read but could not be found on the page   -> listed
  * a row on the page that looks like a bill (a date + an amount) with no band
    -> outlined in red and listed: the parser may have missed it.

Where the words are:
  * PDF with text      -> pdfplumber word boxes (exact)
  * scanned PDF / image -> Tesseract OCR word boxes
Nothing here writes to the statement file; it only reads it.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

DPI = 150
SCALE = DPI / 72

# bucket key -> (label, RGB). Colour encodes the bucket only.
BUCKETS: Dict[str, Tuple[str, Tuple[int, int, int]]] = {
    "matched":  ("Matched - in QBO, same amount",            (76, 175, 80)),
    "approve":  ("Matched, NOT APPROVED (memo) - chase the PM", (205, 220, 57)),
    "checkqbo": ("Matched, approval is in QBO - check there",  (0, 172, 193)),
    "mismatch": ("Amount differs from QBO",                   (255, 152, 0)),
    "tax":      ("Sales tax added",                           (255, 112, 67)),
    "lag":      ("Paid in QBO, vendor still shows it open",   (66, 165, 245)),
    "enter":    ("Not in QBO - enter it",                     (229, 57, 53)),
    "other":    ("Payment / credit / balance forward",        (149, 117, 205)),
    "skipped":  ("Another customer's line - left out on purpose", (158, 158, 158)),
}
MISSED_RGB = (198, 40, 40)

DATE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")
MONEY_RE = re.compile(r"^\(?-?\$?[\d,]*\.\d{2}\)?-?$")


@dataclass
class Word:
    text: str
    x0: float
    x1: float
    top: float
    bottom: float


@dataclass
class Row:                     # a visual line of words on one page
    top: float
    bottom: float
    words: List[Word]

    @property
    def text(self) -> str:
        return " ".join(w.text for w in sorted(self.words, key=lambda w: w.x0))


@dataclass
class Page:
    image: object              # PIL.Image
    rows: List[Row]
    source: str                # "text" | "ocr"
    alts: List[List[Row]] = field(default_factory=list)   # OCR: every reading tried


@dataclass
class Result:
    pages: List[object] = field(default_factory=list)          # marked PIL images
    counts: Dict[str, int] = field(default_factory=dict)       # bucket -> lines banded
    not_found: List[str] = field(default_factory=list)         # read but not on the page
    unread_rows: List[str] = field(default_factory=list)       # on the page, no band
    located: int = 0
    total: int = 0
    source: str = ""


# ─────────────────────────── words ───────────────────────────

def _rows_from_words(words: List[Word], tol: float) -> List[Row]:
    rows: List[Row] = []
    for w in sorted(words, key=lambda w: ((w.top + w.bottom) / 2, w.x0)):
        mid = (w.top + w.bottom) / 2
        if rows and abs(mid - (rows[-1].top + rows[-1].bottom) / 2) <= tol:
            r = rows[-1]
            r.words.append(w)
            r.top, r.bottom = min(r.top, w.top), max(r.bottom, w.bottom)
        else:
            rows.append(Row(w.top, w.bottom, [w]))
    return rows


def flatten(img):
    """Transparent PNG (a pasted screenshot) -> RGB on WHITE; a plain convert turns
    the transparent background black and OCR reads nothing (10/08/2026)."""
    from PIL import Image
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        return bg
    return img.convert("RGB")


# OCR readings tried on every image page; mark() keeps the one that places the
# most statement lines on THAT page. No single setting reads every page (10/08/2026
# training on 69 old workbooks): mode 6 (one text block) reads most statements,
# mode 4 (a column of lines) is what reads a gridded table (Cowtown letter: 2 vs 27
# invoice #s); enlarging helps a screenshot and a 150-dpi Burnco page, hurts others.
OCR_TARGET = 1700
OCR_READINGS = ((6, True), (4, False), (4, True), (6, False))     # (mode, enlarge)


def _ocr_reading(img, psm: int, enlarge: bool) -> List[Row]:
    import pytesseract
    from PIL import Image
    f = OCR_TARGET / img.size[0] if enlarge and img.size[0] < OCR_TARGET else 1.0
    work = img.convert("L")
    if f > 1.0:
        work = work.resize((int(img.size[0] * f), int(img.size[1] * f)), Image.LANCZOS)
    d = pytesseract.image_to_data(work, output_type=pytesseract.Output.DICT, config=f"--psm {psm}")
    words = [Word(t, d["left"][i] / f, (d["left"][i] + d["width"][i]) / f, d["top"][i] / f,
                  (d["top"][i] + d["height"][i]) / f)
             for i, t in enumerate(d["text"]) if t.strip() and float(d["conf"][i]) > 20]
    return _rows_from_words(words, tol=8 * img.size[0] / 1275)


def _ocr_page(img) -> "Page":
    alts = [_ocr_reading(img, psm, big) for psm, big in OCR_READINGS]
    return Page(img, alts[0], "ocr", alts)


def _ocr_rows(img) -> List[Row]:
    return _ocr_reading(img, *OCR_READINGS[0])


def load_pages(src: Path, max_pages: int = 30) -> List[Page]:
    """The statement's pages as images + their words. Text PDFs use the PDF's own
    words; a scanned page (no text) and an image file go through OCR."""
    from PIL import Image
    ext = src.suffix.lower()
    pages: List[Page] = []
    if ext == ".pdf":
        import pdfplumber
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(str(src))
        try:
            with pdfplumber.open(str(src)) as plumb:
                for i in range(min(len(pdf), max_pages)):
                    img = pdf[i].render(scale=SCALE).to_pil().convert("RGB")
                    ws = plumb.pages[i].extract_words(keep_blank_chars=False, use_text_flow=False)
                    words = [Word(w["text"], w["x0"] * SCALE, w["x1"] * SCALE,
                                  w["top"] * SCALE, w["bottom"] * SCALE) for w in ws]
                    if len(words) >= 5:
                        pages.append(Page(img, _rows_from_words(words, tol=2.5 * SCALE), "text"))
                    else:
                        pages.append(_ocr_page(img))
        finally:
            pdf.close()
    else:
        with Image.open(src) as im:
            img = flatten(im)
        pages.append(_ocr_page(img))
    return pages


def pages_from_images(images: List[object]) -> List[Page]:
    """Pages already rendered (the images attached to an old reconciliation Excel)."""
    return [_ocr_page(flatten(im)) for im in images]


# ─────────────────────────── matching ───────────────────────────

def _norm(t: str) -> str:
    return re.sub(r"[^A-Z0-9-]", "", (t or "").upper()).strip("-")


def _amount_forms(x: float) -> List[str]:
    a = abs(x)
    return [f"{a:,.2f}", f"{a:.2f}"]


_OCR_SWAP = str.maketrans("OQDILSBZG", "001115826")


def _loose(t: str) -> str:
    """OCR look-alikes folded: INV37S01 == INV37501, SUNO7863 == SUN07863."""
    return t.translate(_OCR_SWAP)


def _one_off(a: str, b: str) -> bool:
    """Same length, at most one character different."""
    return len(a) == len(b) and sum(x != y for x, y in zip(a, b)) <= 1


def _word_hits(word: str, ref: str, ocr: bool = False) -> bool:
    w = _norm(word)
    if not w or not ref:
        return False
    if w == ref:
        return True
    # glued tokens: 'INV#16541.' -> 'INV16541', '1080594MC6749' (Croell), 'RI34216970'
    if len(ref) >= 4 and (w.startswith(ref) or w.endswith(ref) or (len(ref) >= 6 and ref in w)):
        return True
    if ocr and len(ref) >= 6:                  # a read page: look-alikes / one bad, lost or extra char
        lw, lr = _loose(w), _loose(ref)
        if lw == lr or lr in lw or _one_off(lw, lr) or _one_off(lw[-len(lr):], lr):
            return True
        if abs(len(lw) - len(lr)) == 1 and min(len(lw), len(lr)) >= 6 and (
                lr.endswith(lw) or lw.endswith(lr) or lr.startswith(lw) or lw.startswith(lr)):
            return True                       # '§0033829287' for 50033829287
    return False


def _row_has_amount_loose(row: Row, amount: float) -> bool:
    """The amount, allowing ONE misread digit (OCR: 764.30 for 754.30)."""
    want = f"{abs(amount):,.2f}"
    return any(_one_off(re.sub(r"[^0-9.,]", "", w.text), want) for w in row.words)


def _row_has_amount(row: Row, amount: float) -> bool:
    t = row.text.replace("$", "")
    return any(f in t for f in _amount_forms(amount))


_EXTRA: Dict[int, List[Tuple[int, int]]] = {}     # line index -> other rows showing the same bill


def _pick_readings(pages: List[Page], lines: list) -> None:
    """For each OCR page keep the reading that shows the most statement lines (its
    ref - exact or look-alike - with the amount on the same row)."""
    refs = [(_norm(getattr(ln, "ref", "")), getattr(ln, "amount", 0.0)) for ln in lines]
    for pg in pages:
        if len(pg.alts) < 2:
            continue
        def score(rows):
            return sum(1 for ref, amt in refs if ref and any(
                _row_has_amount(r, amt) and any(_word_hits(w.text, ref, True) for w in r.words)
                for r in rows))
        pg.rows = max(pg.alts, key=score)


def locate(pages: List[Page], lines: list) -> List[Optional[Tuple[int, int]]]:
    """(page index, row index) for every statement line, or None. A row is used
    once. Passes, so a weaker match never takes a row a stronger one owns:
      1. the exact ref (among rows showing it, the one with the amount wins)
      2. a look-alike ref on a READ (OCR) page - only on a row with the amount
      3. the date + amount on one row, else an amount on exactly one free row."""
    _pick_readings(pages, lines)
    used = set()
    extra = _EXTRA
    extra.clear()
    out: List[Optional[Tuple[int, int]]] = [None] * len(lines)
    once: Dict[str, int] = {}
    for ln in lines:
        k = _norm(getattr(ln, "ref", ""))
        once[k] = once.get(k, 0) + 1
    rows = [(pi, ri, row, pg.source == "ocr") for pi, pg in enumerate(pages)
            for ri, row in enumerate(pg.rows)]

    for i, ln in enumerate(lines):                       # pass 1 - exact ref
        ref = _norm(getattr(ln, "ref", ""))
        cands = [(pi, ri) for pi, ri, row, _o in rows if (pi, ri) not in used and (
            any(_word_hits(w.text, ref) for w in row.words) if ref
            else "FORWARD" in row.text.upper())]
        best = next((c for c in cands if _row_has_amount(pages[c[0]].rows[c[1]], ln.amount)), None)
        if best is None and cands and ref and ref not in ("PMT", "DISC"):
            best = cands[0]
        if best is None:
            continue
        used.add(best)
        out[i] = best
        # The same bill printed again (a tear-off stub, Cintas): band every copy -
        # only for a real invoice # read once (never PMT / DISC / a repeated #).
        if ref not in ("PMT", "DISC") and len(ref) >= 4 and once[ref] == 1:
            for c in cands:
                if c != best and c not in used and _row_has_amount(pages[c[0]].rows[c[1]], ln.amount):
                    used.add(c)
                    extra.setdefault(i, []).append(c)

    # pass 2 - look-alike ref on a READ (OCR) page. Every (line, row) candidate is
    # ranked and the strongest are placed first, so a sequential neighbour
    # (CM_SUN00232 vs 00233) never takes the row of the true one.
    cand: List[Tuple[Tuple[int, int], int, Tuple[int, int]]] = []
    for i, ln in enumerate(lines):
        ref = _norm(getattr(ln, "ref", ""))
        if out[i] is not None or len(ref) < 6:
            continue
        lr = _loose(ref)
        hits = []
        for pi, ri, row, ocr in rows:
            if not ocr or (pi, ri) in used:
                continue
            words = [_loose(_norm(w.text)) for w in row.words]
            same = any(w == lr or (len(w) > len(lr) and lr in w) for w in words)
            near = same or any(_word_hits(w.text, ref, True) for w in row.words)
            if not near:
                continue
            amt = 0 if _row_has_amount(row, ln.amount) else 1 if _row_has_amount_loose(row, ln.amount) else 2
            hits.append(((0 if same else 1, amt), (pi, ri)))
        for rank, pos in hits:
            if rank[1] < 2 or (rank[0] == 0 and len(hits) == 1 and len(ref) >= 7):
                cand.append((rank, i, pos))
    for rank, i, pos in sorted(cand):
        if out[i] is None and pos not in used:
            used.add(pos)
            out[i] = pos

    for i, ln in enumerate(lines):                       # pass 3 - date / amount
        if out[i] is not None or not getattr(ln, "amount", 0):
            continue
        free = [(pi, ri) for pi, ri, row, _o in rows
                if (pi, ri) not in used and _row_has_amount(row, ln.amount)]
        dated = [c for c in free if _has_line_date(pages[c[0]].rows[c[1]], getattr(ln, "date", ""))]
        pick = dated[0] if dated else (free[0] if len(free) == 1 and abs(ln.amount) >= 1 else None)
        if pick is not None:
            used.add(pick)
            out[i] = pick
    return out


def _has_line_date(row: Row, iso: str) -> bool:
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    if not m:
        return False
    y, mo, d = m.groups()
    forms = {f"{int(mo)}/{int(d)}/{y}", f"{mo}/{d}/{y}", f"{mo}/{d}/{y[2:]}", f"{int(mo)}/{int(d)}/{y[2:]}"}
    forms |= {f.replace("/", "-") for f in forms}          # OCR / some layouts: 06-04-2026
    return any(f in row.text for f in forms)


def _bill_like(row: Row) -> bool:
    """A row that looks like a bill line: a date AND ends in an amount, and is
    not a total / aging / header line."""
    t = row.text
    up = t.upper()
    if not DATE_RE.search(t) or not row.words:
        return False
    last = sorted(row.words, key=lambda w: w.x0)[-1].text
    if not MONEY_RE.match(last.replace("$", "")):
        return False
    money = [w.text for w in sorted(row.words, key=lambda w: w.x0)
             if MONEY_RE.match(w.text.replace("$", "")) and re.search(r"\d", w.text)]
    amt = money[-2] if len(money) >= 2 else money[-1] if money else ""
    if amt and float(re.sub(r"[^0-9.]", "", amt) or 0) == 0:
        return False      # a $0 row owes nothing (2+ amounts: the last is the running balance)
    return not any(k in up for k in ("TOTAL", "AMOUNT DUE", "BALANCE DUE", "PAGE ", "STATEMENT DATE",
                                     "AS OF", "FINANCE CHARGES", "BALANCE BEFORE", "INVOICES & ON"))


# ─────────────────────────── drawing ───────────────────────────

def mark(pages: List[Page], lines: list, bucket_of: Callable[[object], str],
         skipped: Optional[list] = None) -> Result:
    """Band every located line in its bucket colour; outline bill-like rows that
    got no band. `skipped` = rows the parser left out on purpose (another
    customer's) - greyed, not flagged. Returns the marked images and the misses."""
    from PIL import Image, ImageDraw
    res = Result(total=len(lines), source=",".join(sorted({p.source for p in pages})))
    skipped = skipped or []
    where = locate(pages, list(lines) + list(skipped))
    copies = dict(_EXTRA)
    bands: Dict[int, List[Tuple[Row, str]]] = {}
    for i, w in enumerate(where[len(lines):]):
        for pos in ([w] if w else []) + copies.get(len(lines) + i, []):
            bands.setdefault(pos[0], []).append((pages[pos[0]].rows[pos[1]], "skipped"))
    for i, (ln, w) in enumerate(zip(lines, where[:len(lines)])):
        if w is None:
            amt = getattr(ln, "amount", 0.0)
            res.not_found.append(f"{getattr(ln, 'ref', '') or 'balance forward'} · "
                                 f"{getattr(ln, 'date', '')} · {amt:,.2f}")
            continue
        key = bucket_of(ln)
        for pos in [w] + copies.get(i, []):
            bands.setdefault(pos[0], []).append((pages[pos[0]].rows[pos[1]], key))
        res.counts[key] = res.counts.get(key, 0) + 1
        res.located += 1
    for pi, pg in enumerate(pages):
        img = pg.image.copy()
        over = Image.new("RGBA", img.size, (0, 0, 0, 0))
        dr = ImageDraw.Draw(over)
        W = img.size[0]
        page_bands = bands.get(pi, [])
        for row, key in page_bands:
            rgb = BUCKETS.get(key, BUCKETS["other"])[1]
            y0, y1 = row.top - 3, row.bottom + 3
            dr.rectangle([0, y0, W, y1], fill=rgb + (70,))
            dr.rectangle([0, y0, 10, y1], fill=rgb + (255,))
        covered = [(r.top - 3, r.bottom + 3) for r, _k in page_bands]
        start = _table_start(pg, page_bands)
        for row in pg.rows:
            if row.top < start or not _bill_like(row):
                continue
            mid = (row.top + row.bottom) / 2
            if any(a <= mid <= b for a, b in covered):
                continue
            dr.rectangle([4, row.top - 4, W - 4, row.bottom + 4], outline=MISSED_RGB + (255,), width=4)
            res.unread_rows.append(f"page {pi + 1}: {row.text[:90]}")
        res.pages.append(Image.alpha_composite(img.convert("RGBA"), over).convert("RGB"))
    return res


def _table_start(pg: Page, page_bands: list) -> float:
    """Where the bill table starts: its column header row (a date/invoice word with
    an amount/balance word), else the first banded row, else the top."""
    for row in pg.rows:
        up = row.text.upper()
        if ("DATE" in up or "INVOICE" in up or "NUM" in up) and \
                any(k in up for k in ("AMOUNT", "BALANCE", "DUE", "CHARGE")):
            return row.bottom
    return min((r.top for r, _k in page_bands), default=0.0)


def bucket_for(category: str, approval: str = "approved") -> str:
    """ReconRow.category (+ shared/bill_approval state for a matched bill) -> bucket."""
    matched = {"not approved": "approve", "check QBO": "checkqbo"}.get(approval, "matched")
    return {"MATCHED": matched, "CLERK_AMOUNT_MISMATCH": "mismatch",
            "VENDOR_TAX_VIOLATION": "tax", "LIKELY_VENDOR_LAG": "lag",
            "MISSING_IN_QBO": "enter"}.get(category, "other")
