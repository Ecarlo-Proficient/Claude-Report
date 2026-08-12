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


def _overlap(a, b):
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def verify_xlsx(path) -> list:
    """Return a list of corruption issues found in the .xlsx at `path`.
    Empty list == clean. Never raises on a well-formed zip."""
    issues = []
    with zipfile.ZipFile(str(path)) as z:
        names = z.namelist()

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
