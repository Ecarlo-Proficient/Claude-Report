"""xlsx corruption verifier — the standing guard so an openpyxl file NEVER
reaches the user with the Excel "we found a problem with some content" repair
prompt again (the user 2026-08-08: "this constantly keeps happening").

Runs the whole hard-won checklist at the XML layer, on a SAVED .xlsx, so every
Excel-writing tool can self-check before handing the file over. Cheap, no Excel
needed. `verify_xlsx(path)` returns a list of human-readable issues (empty =
clean); `assert_clean(path)` raises ValueError if any.

Vectors (each has bitten us at least once):
  · Table displayName must be a valid Excel name (no space/comma/dash/em-dash;
    start with a letter or underscore) — the 'WIP as of …' bug, 2026-08-08.
  · Table columns: no blank, no duplicate names.
  · Inline rich text (multi-run `<is>`) — Mac Excel 'String properties', 2026-07-21.
  · Frozen-pane sheet view: exactly one selection, no stray `topRight` — the
    'Repaired Records: View' bug, 2026-08-07.
  · Overlapping merged cells.
  · Style / dxf indices in range.
  · A table column NAME that disagrees with the header cell under it — an
    empty header cell becomes a tableColumn literally named "None", which is
    non-blank so the blank-name check never fires. The table is internally
    valid and still disagrees with its sheet, and Excel repairs it. CP800,
    2026-09-03: a column-insert pass blanked B1 while the ref still began at B1.
  · Cells out of ascending column order within a <row>, duplicate cell refs,
    or <row> elements out of order — ECMA-376 requires ascending order and Excel
    silently drops/repairs the row. Bit us 2026-08-26 when a hand-XML edit
    inserted new <c> elements at the START of a row instead of at their
    column position (the mileage template repair).
"""
import re
import zipfile

try:
    from openpyxl.utils import range_boundaries
except Exception:                       # openpyxl always present in this repo
    range_boundaries = None

_VALID_TABLE_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_SHEET_RE = re.compile(r"xl/worksheets/sheet\d+\.xml$")
_TABLE_RE = re.compile(r"xl/tables/table\d+\.xml$")


def _col_index(col: str) -> int:
    """A1-style column letters -> 1-based index, without importing openpyxl."""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n

def _row_cells(xml: str, row: int, shared: list) -> dict:
    """{column index -> text} for one <row>, inline strings and shared alike."""
    out = {}
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % row, xml, re.S)
    if not m:
        return out
    for c in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*)>(.*?)</c>', m.group(1), re.S):
        ref, attrs, body = c.group(1), c.group(2), c.group(3)
        tm = re.search(r't="(\w+)"', attrs)
        t = tm.group(1) if tm else "n"
        if t == "inlineStr":
            v = re.search(r"<t[^>]*>(.*?)</t>", body, re.S)
            out[_col_index(ref)] = v.group(1) if v else ""
        elif t == "s":
            v = re.search(r"<v>(\d+)</v>", body)
            i = int(v.group(1)) if v else -1
            out[_col_index(ref)] = shared[i] if 0 <= i < len(shared) else ""
        else:
            v = re.search(r"<v>(.*?)</v>", body, re.S)
            out[_col_index(ref)] = v.group(1) if v else ""
    return out


def _col_name(i: int) -> str:
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _overlap(a, b):
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def verify_xlsx(path) -> list:
    """Return a list of corruption issues found in the .xlsx at `path`.
    Empty list == clean. Never raises on a well-formed zip."""
    issues = []
    with zipfile.ZipFile(str(path)) as z:
        names = z.namelist()
        shared_strings = []
        if "xl/sharedStrings.xml" in names:
            _sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
            shared_strings = [re.sub(r"<.*?>", "", si)
                              for si in re.findall(r"<si>(.*?)</si>", _sx, re.S)]

        # ── tables: valid displayName, clean columns, UNIQUE names workbook-wide ──
        all_table_names = []
        for n in names:
            if not _TABLE_RE.match(n):
                continue
            x = z.read(n).decode("utf-8", "replace")
            m = re.search(r'displayName="([^"]*)"', x)
            if m:
                all_table_names.append(m.group(1))
                if not _VALID_TABLE_NAME.match(m.group(1)):
                    issues.append(f"{n}: invalid table name {m.group(1)!r} "
                                  f"(no spaces/commas/dashes; must start with a letter)")
            cols = re.findall(r'<tableColumn\b[^>]*\bname="([^"]*)"', x)
            if any(not c.strip() for c in cols):
                issues.append(f"{n}: a table column has a blank name")
            dups = {c for c in cols if cols.count(c) > 1}
            if dups:
                issues.append(f"{n}: duplicate table column name(s) {sorted(dups)}")
        # A table displayName must be unique across the WHOLE workbook, or Excel
        # (and openpyxl on re-read) rejects it — the combined-file bug, 2026-08-08.
        dup_tables = {t for t in all_table_names if all_table_names.count(t) > 1}
        if dup_tables:
            issues.append(f"duplicate table name(s) across the workbook: {sorted(dup_tables)}")

        # ── styles.xml internal consistency. A <dxf> may hold AT MOST ONE
        #    <numFmt>; a declared count that disagrees with the actual child
        #    count, or a custom numFmtId (>=164) used by a cellXf but never
        #    defined in <numFmts>, is the same class of defect. All three trip
        #    Excel's repair prompt while LibreOffice opens the file happily —
        #    2026-08-26, a str.replace() without a count injected a second
        #    <numFmt> into a <dxf> during a hand-XML edit. ──
        if "xl/styles.xml" in names:
            sx = z.read("xl/styles.xml").decode("utf-8", "replace")
            dblk = re.search(r"<dxfs\b.*?</dxfs>", sx, re.S)
            if dblk:
                for i, dx in enumerate(re.findall(r"<dxf>.*?</dxf>", dblk.group(0), re.S)):
                    if dx.count("<numFmt ") > 1:
                        issues.append(f"xl/styles.xml: dxf #{i} has "
                                      f"{dx.count('<numFmt ')} <numFmt> children (max 1)")
            for tag, child in (("numFmts", "numFmt"), ("fonts", "font"),
                               ("fills", "fill"), ("borders", "border"),
                               ("cellStyleXfs", "xf"), ("cellXfs", "xf"),
                               ("cellStyles", "cellStyle"), ("dxfs", "dxf")):
                cm = re.search(r'<%s count="(\d+)"' % tag, sx)
                blk = re.search(r"<%s\b.*?</%s>" % (tag, tag), sx, re.S)
                if not cm or not blk:
                    continue
                actual = len(re.findall(r"<%s[ />]" % child, blk.group(0)))
                if int(cm.group(1)) != actual:
                    issues.append(f"xl/styles.xml: <{tag}> declares "
                                  f"count={cm.group(1)} but holds {actual}")
            cxb = re.search(r"<cellXfs\b.*?</cellXfs>", sx, re.S)
            if cxb:
                defined = {int(v) for v in re.findall(r'<numFmt numFmtId="(\d+)"', sx)}
                used = {int(v) for v in re.findall(r'<xf numFmtId="(\d+)"', cxb.group(0))}
                missing = sorted(u for u in used if u >= 164 and u not in defined)
                if missing:
                    issues.append(f"xl/styles.xml: cellXfs use custom numFmtId "
                                  f"{missing} that <numFmts> never defines")

        # ── styles: index ceilings ──
        ncx, ndxf = 10 ** 9, 0
        if "xl/styles.xml" in names:
            st = z.read("xl/styles.xml").decode("utf-8", "replace")
            mm = re.search(r'<cellXfs count="(\d+)"', st)
            if mm:
                ncx = int(mm.group(1))
            ndxf = len(re.findall(r"<dxf>", st))

        # ── worksheets: rich text, sheet view, merges, style refs ──
        for n in names:
            if not _SHEET_RE.match(n):
                continue
            x = z.read(n).decode("utf-8", "replace")
            if any(si.count("<r>") > 1 for si in re.findall(r"<is>(.*?)</is>", x, re.S)):
                issues.append(f"{n}: inline rich-text runs (Mac Excel rejects these)")
            if "<pane " in x:
                sels = re.findall(r"<selection[^>]*/>", x)
                if len(sels) > 1 or "topRight" in x:
                    issues.append(f"{n}: invalid frozen-pane sheet view "
                                  f"({len(sels)} selections"
                                  f"{', has topRight' if 'topRight' in x else ''})")
            merges = re.findall(r'<mergeCell ref="([^"]+)"', x)
            if range_boundaries and merges:
                boxes = [range_boundaries(mrf) for mrf in merges]
                found = False
                for i in range(len(boxes)):
                    for j in range(i + 1, len(boxes)):
                        if _overlap(boxes[i], boxes[j]):
                            issues.append(f"{n}: overlapping merged cells "
                                          f"{merges[i]} / {merges[j]}")
                            found = True
                            break
                    if found:
                        break
            for s in re.findall(r'\bs="(\d+)"', x):
                if int(s) >= ncx:
                    issues.append(f"{n}: style index {s} ≥ cellXfs count {ncx}")
                    break
            for d in re.findall(r'dxfId="(\d+)"', x):
                if int(d) >= ndxf:
                    issues.append(f"{n}: dxfId {d} ≥ dxf count {ndxf}")
                    break

            # ── cell / row ordering: ECMA-376 requires <c> ascending by column
            #    within a <row>, and <row> ascending by r. Excel does NOT warn —
            #    it "repairs" the file and drops the offending row. 2026-08-26. ──
            for rm2 in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
                cols = re.findall(r'<c r="([A-Z]+)\d+"', rm2.group(2))
                idx = [_col_index(c) for c in cols]
                if idx != sorted(idx):
                    issues.append(f"{n}: row {rm2.group(1)} has cells out of column "
                                  f"order ({'/'.join(cols)}) — Excel will drop the row")
                    break
                if len(set(idx)) != len(idx):
                    issues.append(f"{n}: row {rm2.group(1)} has duplicate cell refs")
                    break
            rows_seen = [int(v) for v in re.findall(r'<row r="(\d+)"', x)]
            if rows_seen != sorted(rows_seen):
                issues.append(f"{n}: <row> elements are not in ascending order")

        # ── table range must stay WITHIN its sheet's used rows. A stale table
        #    ref after a row insert/delete (ref points past the last row) is the
        #    top "we found a problem with some content" cause — the 12-31 build
        #    bug, 2026-08-17 (deleted 60 rows, table still said A3:R118). ──
        sheet_maxrow = {}
        for n in names:
            if not _SHEET_RE.match(n):
                continue
            x = z.read(n).decode("utf-8", "replace")
            dm = re.search(r'<dimension ref="[A-Z]+\d+:[A-Z]+(\d+)"', x)
            if dm:
                sheet_maxrow[n.split("/")[-1]] = int(dm.group(1))
        for n in names:
            sm = re.match(r"xl/worksheets/(sheet\d+\.xml)$", n)
            rel = f"xl/worksheets/_rels/{sm.group(1)}.rels" if sm else None
            if not sm or rel not in names:
                continue
            rx = z.read(rel).decode("utf-8", "replace")
            smax = sheet_maxrow.get(sm.group(1))
            for tgt in re.findall(r'Target="([^"]*tables/table\d+\.xml)"', rx):
                tfn = "xl/tables/" + tgt.split("/")[-1]
                if tfn not in names or smax is None:
                    continue
                tx = z.read(tfn).decode("utf-8", "replace")
                rm = re.search(r'\bref="[A-Z]+\d+:[A-Z]+(\d+)"', tx)
                nm = re.search(r'displayName="([^"]*)"', tx)
                # ── a table's column NAMES must match the header cells under
                #    its ref. openpyxl writes an EMPTY header cell out as the
                #    literal string "None", which is non-blank, so the
                #    blank-name check above never sees it (CP800, 2026-09-03).
                full = re.search(r'\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', tx)
                cols_named = re.findall(r'<tableColumn[^>]*\bname="([^"]*)"', tx)
                if full and cols_named:
                    c0 = _col_index(full.group(1))
                    c1 = _col_index(full.group(3))
                    hrow = int(full.group(2))
                    span = c1 - c0 + 1
                    tname = nm.group(1) if nm else "?"
                    if span != len(cols_named):
                        issues.append(f"{tfn}: table {tname!r} ref spans {span} "
                                      f"column(s) but declares {len(cols_named)} "
                                      f"tableColumn(s)")
                    cells = _row_cells(z.read(n).decode("utf-8", "replace"),
                                       hrow, shared_strings)
                    for k, cname in enumerate(cols_named):
                        actual = (cells.get(c0 + k) or "").strip()
                        where = f"{_col_name(c0 + k)}{hrow}"
                        if not actual:
                            issues.append(
                                f"{tfn}: table {tname!r} column {k + 1} ({where}) is "
                                f"EMPTY on the sheet but the table names it "
                                f"{cname!r} — Excel will offer to repair")
                        elif actual != cname:
                            issues.append(
                                f"{tfn}: table {tname!r} column {k + 1} ({where}) "
                                f"sheet says {actual!r}, table says {cname!r}")
                if rm and int(rm.group(1)) > smax:
                    issues.append(
                        f"{tfn}: table {nm.group(1) if nm else '?'!r} ref ends at "
                        f"row {rm.group(1)} but sheet {sm.group(1)} has only {smax} "
                        f"rows — stale table range after a row edit (fix the table "
                        f"ref, or drop the table)")
    return issues


def assert_clean(path) -> None:
    """Raise ValueError if the .xlsx would trip Excel's repair prompt."""
    issues = verify_xlsx(path)
    if issues:
        raise ValueError("xlsx corruption check FAILED for "
                         f"{path}:\n  - " + "\n  - ".join(issues))


def safe_table_name(raw: str, seen=None) -> str:
    """A guaranteed-valid, unique Excel table displayName from any string."""
    name = re.sub(r"[^A-Za-z0-9]", "", str(raw)) or "T"
    if not name[0].isalpha():
        name = "T" + name
    name = name[:255]
    if seen is not None:
        base, k = name, 2
        while name in seen:
            name = f"{base}{k}"
            k += 1
        seen.add(name)
    return name


if __name__ == "__main__":
    import sys
    for p in sys.argv[1:]:
        found = verify_xlsx(p)
        print(f"{p}: {'CLEAN' if not found else 'ISSUES'}")
        for it in found:
            print("  -", it)
