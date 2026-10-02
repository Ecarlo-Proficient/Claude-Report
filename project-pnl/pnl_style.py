"""ONE look for every sheet of the project P&L workbook (the owner 2026-10-02:
"something about all of this that is just wrong, amateur, hard to read,
disconnected and untrustworthy"). Every sheet was built one request at a time
with its own fills, fonts and grid; new and rebuilt sheets take their bars,
bands, grid lines, number formats and tie-out marks from here instead.

Colour encodes meaning only, the same way on every sheet:
  navy bar     = a sheet or block title          green  = income / ties
  orange band  = costs                           grey   = sub-headers, totals
  amber        = a figure to look at (info, not an error)
  red          = does not tie / over budget
"""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SIZE = 12                       # body; titles SIZE + 1, notes SIZE - 1
NAVY = "1F3A5F"
LINK = "0563C1"
RED = "C00000"

BAR = PatternFill("solid", fgColor=NAVY)
INCOME = PatternFill("solid", fgColor="E2EFDA")
COSTS = PatternFill("solid", fgColor="FCE4D6")
SUB = PatternFill("solid", fgColor="EDEDED")
TOTAL = PatternFill("solid", fgColor="D9D9D9")
INFO = PatternFill("solid", fgColor="FFF2CC")
OK = PatternFill("solid", fgColor="C6EFCE")
BAD = PatternFill("solid", fgColor="FFC7CE")

MONEY = '#,##0.00;[Red]-#,##0.00;"-"'      # zero reads as a dash
PCT = '0%'
DATE = 'mm/dd/yyyy'

GRID = Side(style="thin", color="D9D9D9")
EDGE = Side(style="medium", color="000000")
RULE = Side(style="thin", color="000000")


def font(bold=False, color="000000", size=SIZE, underline=None, italic=False):
    return Font(bold=bold, color=color, size=size, underline=underline, italic=italic)


def bar(ws, row, c0, c1, text, fill=BAR, merge=True):
    """A title bar: white bold text on `fill`, merged and centred across c0..c1."""
    c = ws.cell(row=row, column=c0, value=text)
    c.font = font(bold=True, color="FFFFFF", size=SIZE + 1)
    c.alignment = Alignment(horizontal="center", vertical="center")
    for cc in range(c0, c1 + 1):
        ws.cell(row=row, column=cc).fill = fill
    if merge and c1 > c0:
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)


def header(ws, row, labels, c0=1, fill=SUB, align_from=2):
    """Column headers: bold on the light grey band, numbers right-aligned."""
    for i, h in enumerate(labels):
        c = ws.cell(row=row, column=c0 + i, value=h)
        c.font = font(bold=True)
        c.fill = fill
        c.alignment = Alignment(horizontal="right" if i + 1 >= align_from and h else "left",
                                vertical="center", wrap_text=True)
        c.border = Border(bottom=RULE)


def grid(ws, r0, r1, c0, c1):
    """A light rule under every row so a table reads as rows, not floating text."""
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border
            if b.bottom is None or b.bottom.style is None:
                cell.border = Border(left=b.left, right=b.right, top=b.top, bottom=GRID)


def box(ws, r0, r1, c0, c1, side=EDGE):
    """Outline r0..r1 x c0..c1, keeping each cell's inner borders."""
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border
            cell.border = Border(left=side if c == c0 else b.left,
                                 right=side if c == c1 else b.right,
                                 top=side if r == r0 else b.top,
                                 bottom=side if r == r1 else b.bottom)


def money(cell, bold=False, fill=None):
    cell.number_format = MONEY
    cell.font = font(bold=bold)
    cell.alignment = Alignment(horizontal="right")
    if fill is not None:
        cell.fill = fill
    return cell


def tie_formula(diff_ref: str, ok="✓ ties", bad="✗ off") -> str:
    """The text of a tie-out cell: a cent of rounding is still a tie."""
    return f'=IF(ABS({diff_ref})<0.005,"{ok}","{bad}")'


def tie_colours(ws, rng: str):
    """Green when the cell says it ties, red when it does not - one rule pair,
    so every tie-out on every sheet reads the same."""
    first = rng.split(":")[0].replace("$", "")
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'LEFT({first},1)="✓"'], fill=OK, font=Font(bold=True, color="006100")))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'LEFT({first},1)="✗"'], fill=BAD, font=Font(bold=True, color="9C0006")))
