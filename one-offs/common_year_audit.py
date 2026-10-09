#!/usr/bin/env python3
"""
common_year_audit.py - which parts of the Common share were worked on this year?

The question (the user 2026-10-05): "how do we audit the files so we only bring over files
from 2026" - for the Synology Common -> SharePoint move.

Read-only. One walk of the share, one row per JOB (not per file), so a live job is never split:
if anything inside a job folder changed in the year, the whole job moves; if nothing did, it
stays behind as archive. Moving single files would carry a 2026 invoice over without the 2024
plans and stress letter of the same job.

What counts as one job / one unit:
  RP jobs           CURRENT PROJECTS/Residential/<builder>/<address>
  CP jobs           CURRENT PROJECTS/Awarded Projects Commercial projects/<job>
                    (and .../2022 & PREVIOUS/<job>)
  CP bids           CURRENT PROJECTS/Commercial/BIDS <year>/<bid>
  Schedule / bonds  OPERATIONS/SCHEDULE/<year>, CURRENT PROJECTS/CITY BONDS/<year>
  Staff folders     the person's whole folder at the top of the share
  everything else   the first folder under each top folder (one vendor, one sub, ...)
  Loose files above those levels are their own unit.

A file counts for the year when it was CREATED in it or EDITED AND SAVED in it (last modified).
Opening a file does not change either date. Files that existed before the share was copied onto
the Synology all show a created date of 05/30/2025 (the copy day), so the created date can only
ever say "new this year" for files really made this year - it never wrongly marks an old file.

Speed: the share is read over the network, about 4 ms per file the first time (265,000 files
= roughly 20-60 minutes, longer while the office is busy). A second run is much faster.

Output in <CompanyHealth>/Analysis/Common <year> audit (mm-dd-yyyy)/:
  Common <year> audit - jobs.csv    one row per job: Move / Archive, counts, last changed
  Common <year> audit - files.csv   every file created or edited in the year, and which
plus a summary by area printed at the end.

    ~/.venvs/proficient/bin/python one-offs/common_year_audit.py
    ~/.venvs/proficient/bin/python one-offs/common_year_audit.py --year 2026 --root /Volumes/Common
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import paths  # noqa: E402
from shared.xlsx_guard import csv_cell  # noqa: E402

JUNK = re.compile(r"^(~\$.*|Thumbs\.db|\.DS_Store|.*\.tmp|\._.*|desktop\.ini)$", re.I)
COMPANY_TOP = {"CURRENT PROJECTS", "IT", "Inventory Folder", "OPERATIONS", "SUBCONTRACTORS",
               "VEHICLE AND EQUIPMENT INFO", "VENDORS"}
SKIP_TOP = {"#recycle"}


def area_and_depth(parts):
    """(area label, how many path parts make one unit) for a path relative to the share root."""
    a = parts[0]
    b = parts[1] if len(parts) > 1 else ""
    c = parts[2] if len(parts) > 2 else ""
    if a not in COMPANY_TOP:
        return "Staff folders", 1
    if a == "CURRENT PROJECTS":
        if b == "Residential":
            return "RP jobs", 4
        if b == "Awarded Projects Commercial projects":
            return "CP jobs", 4 if c == "2022 & PREVIOUS" else 3
        if b == "Commercial":
            return "CP bids / estimating", 4 if c.startswith("BIDS") else 3
        if b == "CITY BONDS":
            return "City bonds", 3
        return "Shared project files", 2
    if a == "OPERATIONS" and b == "SCHEDULE":
        return "Schedule", 3
    return a.title(), 2


def fmt_date(ts):
    return dt.datetime.fromtimestamp(ts).strftime("%m/%d/%Y") if ts else ""


def main():
    ap = argparse.ArgumentParser(description="Common share: which jobs changed in a year (read-only).")
    ap.add_argument("--root", default="/Volumes/Common", help="the mounted Common share")
    ap.add_argument("--year", type=int, default=dt.date.today().year)
    ap.add_argument("--out", help="folder for the two CSVs (default: CompanyHealth/Analysis/Common <year> audit (date)/)")
    args = ap.parse_args()

    root = Path(args.root)
    if not (root / "CURRENT PROJECTS").is_dir():
        sys.exit(f"{root} is not the Common share (no CURRENT PROJECTS folder). Connect it in Finder first.")
    start_ts = dt.datetime(args.year, 1, 1).timestamp()

    outdir = Path(args.out) if args.out else paths.analysis_dir(f"Common {args.year} audit")
    outdir.mkdir(parents=True, exist_ok=True)
    jobs_csv = outdir / f"Common {args.year} audit - jobs.csv"
    files_csv = outdir / f"Common {args.year} audit - files.csv"
    ffh = open(files_csv, "w", newline="", encoding="utf-8-sig")
    fw = csv.writer(ffh)
    fw.writerow(["Area", "Job / folder on Common", "File", "Why it counts", "Created", "Last saved", "MB"])

    # unit key -> [files, bytes, newest save, files in year, bytes in year, created in year, edited in year]
    units = defaultdict(lambda: [0, 0, 0, 0, 0, 0, 0])
    area_of = {}
    t0 = time.time()
    seen = 0
    for top in sorted(os.listdir(root)):
        if top.startswith(".") or top in SKIP_TOP:
            continue
        for dirpath, dirnames, filenames in os.walk(root / top):
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]
            rel_dir = Path(dirpath).relative_to(root).parts
            for f in filenames:
                if JUNK.match(f):
                    continue
                try:
                    st = os.lstat(os.path.join(dirpath, f))
                except OSError:
                    continue
                parts = rel_dir + (f,)
                area, depth = area_and_depth(parts)
                key = "/".join(parts[:depth])
                area_of[key] = area
                u = units[key]
                u[0] += 1
                u[1] += st.st_size
                u[2] = max(u[2], st.st_mtime)
                born = getattr(st, "st_birthtime", 0) or 0
                created = born >= start_ts
                edited = not created and st.st_mtime >= start_ts
                if created or edited:
                    u[3] += 1
                    u[4] += st.st_size
                    u[5 if created else 6] += 1
                    fw.writerow([csv_cell(area), csv_cell(key), csv_cell("/".join(parts[depth:]) or f),
                                 f"Created in {args.year}" if created else f"Edited in {args.year}",
                                 fmt_date(born), fmt_date(st.st_mtime), round(st.st_size / 1024 ** 2, 2)])
                seen += 1
                if seen % 10000 == 0:
                    print(f"  {seen:,} files read ({int(time.time() - t0)} s)", flush=True)

    ffh.close()
    gb = lambda b: round(b / 1024 ** 3, 2)  # noqa: E731
    rows = sorted(units.items(), key=lambda kv: (area_of[kv[0]], kv[0].lower()))
    with open(jobs_csv, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["Area", "Job / folder on Common", "Decision", "Files", "GB", "Last saved",
                    f"Created in {args.year}", f"Edited in {args.year}", f"GB from {args.year}"])
        for key, (n, size, newest, n_y, s_y, n_c, n_e) in rows:
            w.writerow([csv_cell(area_of[key]), csv_cell(key), "Move" if n_y else "Archive", n, gb(size),
                        fmt_date(newest), n_c, n_e, gb(s_y)])

    summ = defaultdict(lambda: [0, 0, 0, 0, 0, 0, 0])  # units, units moving, files, GB, files moving, GB moving, files in year
    for key, (n, size, newest, n_y, s_y, _c, _e) in units.items():
        s = summ[area_of[key]]
        s[0] += 1
        s[2] += n
        s[3] += size
        s[6] += n_y
        if n_y:
            s[1] += 1
            s[4] += n
            s[5] += size
    print(f"\nRead {seen:,} files in {int(time.time() - t0)} s.\n")
    hdr = f"{'Area':24} {'Jobs':>7} {'Move':>7} | {'All files':>10} {'GB':>7} | {'Files moving':>12} {'GB':>7} | {'From ' + str(args.year):>13}"
    print(hdr)
    print("-" * len(hdr))
    tot = [0] * 7
    for area, s in sorted(summ.items(), key=lambda kv: -kv[1][2]):
        print(f"{area[:24]:24} {s[0]:>7,} {s[1]:>7,} | {s[2]:>10,} {gb(s[3]):>7} | {s[4]:>12,} {gb(s[5]):>7} | {s[6]:>13,}")
        tot = [a + b for a, b in zip(tot, s)]
    print("-" * len(hdr))
    print(f"{'TOTAL':24} {tot[0]:>7,} {tot[1]:>7,} | {tot[2]:>10,} {gb(tot[3]):>7} | {tot[4]:>12,} {gb(tot[5]):>7} | {tot[6]:>13,}")
    print(f"\nOne row per job: {jobs_csv}\nEvery {args.year} file:  {files_csv}")


if __name__ == "__main__":
    main()
