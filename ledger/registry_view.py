"""
registry_view.py - read the systems & process registry for the ledger's Systems tab.

WHAT THIS IS
The registry lives in the AI Brain_Vault as markdown (`02_processes/*.md`), one
file per domain, each holding a pipe table of process rows. This module parses
those tables on demand so the ledger can render them as a LIVE view: edit the
markdown, reload the tab, see the change. Nothing is cached to the ledger DB and
nothing is written back - the vault stays the single owner of registry truth.

It replaced the daily markdown digest (retired 2026-08-19, the owner: "we just
need to have this in the Project Ledger, my systems and processes live view").

ROW SCHEMA (do not change it - it is the vault's, not ours)
    ID | Process | Owner | Operators & touchers | Record | Automation | Cadence | H | State | Life

  H     (health) - how it is running:  green / yellow / red / none
  State (belief) - how sure the description is: confirmed | inferred | proposed | corrected
  Life  (reality) - whether it runs:   idea | agreed | building | live | retired
  State and Life are DIFFERENT AXES and are never merged. The Life column is
  absent from older domain files; a row without one is live.

PARSING NOTES (learned from the eight files as they actually are)
  - Only tables whose header's first cell is exactly "ID" are registry tables.
    The files carry other pipe tables (glossaries, routing) that must be ignored.
  - Cells carry markdown: **bold**, ~~strike~~ (a retired row), [[wiki|links]],
    `code`. Display wants plain text, so it is stripped.
  - A retired row is marked by ~~strike~~ on the ID, by State/Life "retired", or
    both. All three spellings resolve to the same retired flag.
  - Names never appear here: owners are role handles. This module does not read
    the roster and must never resolve a handle to a person.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import paths  # noqa: E402

# Domain files in the order the work happens, matching the registry README.
DOMAIN_FILES = [
    "estimating-and-bidding.md",
    "project-setup.md",
    "field-operations.md",
    "accounts-payable.md",
    "accounts-receivable.md",
    "accounting-and-close.md",
    "payroll.md",
    "platform-and-infra.md",
]

COLUMNS = ["id", "process", "owner", "touchers", "record",
           "automation", "cadence", "health", "state", "life"]

HEALTH = {"\U0001F7E2": ("green", "running"),
          "\U0001F7E1": ("yellow", "fragile"),
          "\U0001F534": ("red", "broken"),
          "⚪": ("none", "nothing to fail yet")}

_STATES = ("confirmed", "corrected", "inferred", "proposed", "retired")
_LIVES = ("idea", "agreed", "building", "live", "retired")
_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
# `# 4 - AP: Bills, Subs, Vendors (AP)` -> num 4, title, code AP.
# The class below MUST keep the em dash: the vault headings are written with one, so
# stripping it here (house style is hyphen/en dash in AUTHORED text) breaks every file.
_TITLE = re.compile(r"^#\s*(\d+)\s*[-–—]\s*(.+?)\s*\(([A-Z]{2,4})\)\s*$")


def _plain(cell: str) -> str:
    """Markdown cell -> display text. Keeps the words, drops the decoration."""
    t = cell.strip()
    t = re.sub(r"\[\[[^\]|]*\|([^\]]*)\]\]", r"\1", t)   # [[target|label]] -> label
    t = re.sub(r"\[\[([^\]]*)\]\]", r"\1", t)            # [[target]]       -> target
    t = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", t)       # [label](url)     -> label
    t = t.replace("~~", "").replace("**", "").replace("`", "")
    t = re.sub(r"(?<!\*)\*(?!\*)", "", t)
    return re.sub(r"\s+", " ", t).strip()


def _split_row(line: str) -> list[str]:
    """Split a markdown table row into cells, tolerating escaped pipes."""
    body = line.strip()
    if body.startswith("|"):
        body = body[1:]
    if body.endswith("|"):
        body = body[:-1]
    return [c.replace("\x00", "|") for c in
            body.replace("\\|", "\x00").split("|")]


def _is_separator(line: str) -> bool:
    return bool(re.fullmatch(r"\|[\s:|-]+\|?", line.strip()))


def _classify_state(raw: str) -> tuple[str, str]:
    """State cell -> (kind, confirmed-on date or '')."""
    low = raw.lower()
    kind = next((s for s in _STATES if s in low), "unknown")
    m = _DATE.search(raw)
    return kind, m.group(0) if m else ""


def _classify_life(raw: str, state_kind: str, struck: bool) -> str:
    """Life cell -> one of _LIVES. Absent means live; a struck row is retired."""
    low = raw.lower()
    for life in _LIVES:
        if life in low:
            return life
    if struck or state_kind == "retired":
        return "retired"
    return "live"


def _parse_file(path: Path) -> dict | None:
    """One domain file -> {code, num, title, file, rows[]}. None if unreadable."""
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return None

    lines = text.splitlines()
    num, title, code = 0, path.stem.replace("-", " ").title(), ""
    for line in lines[:40]:
        m = _TITLE.match(line.strip())
        if m:
            num, title, code = int(m.group(1)), m.group(2).strip(), m.group(3)
            break

    rows: list[dict] = []
    header: list[str] | None = None
    for line in lines:
        stripped = line.strip()
        if not stripped.startswith("|"):
            header = None            # any non-table line ends the block
            continue
        if _is_separator(stripped):
            continue

        cells = _split_row(stripped)
        first = _plain(cells[0]).lower() if cells else ""

        if first == "id":            # a registry table starts here
            header = [_plain(c).lower() for c in cells]
            continue
        if header is None:           # some other table in the file - skip it
            continue

        # A cell containing an unescaped pipe would over-split; fold the excess
        # back into the last column rather than dropping the row.
        if len(cells) > len(header):
            cells = cells[:len(header) - 1] + ["|".join(cells[len(header) - 1:])]

        struck = cells[0].strip().startswith("~~")
        vals = [_plain(c) for c in cells] + [""] * (len(COLUMNS) - len(cells))
        row = dict(zip(COLUMNS, vals[:len(COLUMNS)]))
        if not row["id"]:
            continue

        state_kind, confirmed_on = _classify_state(row["state"])
        hkey, hlabel = HEALTH.get(row["health"].strip()[:1], ("none", ""))
        row.update(
            domain=title, domain_code=code, domain_num=num, file=path.name,
            health_key=hkey, health_label=hlabel,
            state_kind=state_kind, confirmed_on=confirmed_on,
            life_key=_classify_life(row["life"], state_kind, struck),
            retired=struck or state_kind == "retired",
        )
        rows.append(row)

    return {"code": code, "num": num, "title": title,
            "file": path.name, "rows": rows}


def load_registry(root: Path | None = None) -> dict:
    """Parse every domain file. Always returns a dict; missing vault -> ok=False."""
    base = Path(root) if root else paths.process_registry_dir()
    if not base.is_dir():
        return {"ok": False, "source": str(base), "domains": [], "rows": [],
                "error": f"process registry not found at {base} "
                         f"(set ACB_VAULT_DIR in machine.env)"}

    domains, rows, missing = [], [], []
    for name in DOMAIN_FILES:
        parsed = _parse_file(base / name)
        if parsed is None:
            missing.append(name)
            continue
        domains.append(parsed)
        rows.extend(parsed["rows"])

    domains.sort(key=lambda d: (d["num"] or 99, d["title"]))
    return {"ok": True, "source": str(base), "domains": domains, "rows": rows,
            "missing": missing,
            "counts": _counts(rows), "owners": _owners(rows)}


def _counts(rows: list[dict]) -> dict:
    live = [r for r in rows if not r["retired"]]
    def tally(key):
        out: dict[str, int] = {}
        for r in live:
            out[r[key]] = out.get(r[key], 0) + 1
        return out
    return {"total": len(rows), "active": len(live),
            "retired": len(rows) - len(live),
            "health": tally("health_key"), "state": tally("state_kind"),
            "life": tally("life_key")}


def _owners(rows: list[dict]) -> list[dict]:
    """Owner handles with how much each carries. Handles only - never names."""
    agg: dict[str, dict] = {}
    for r in rows:
        if r["retired"] or not r["owner"]:
            continue
        o = agg.setdefault(r["owner"], {"owner": r["owner"], "count": 0, "red": 0})
        o["count"] += 1
        if r["health_key"] == "red":
            o["red"] += 1
    return sorted(agg.values(), key=lambda o: (-o["count"], o["owner"]))


if __name__ == "__main__":
    data = load_registry()
    if not data["ok"]:
        print(data["error"])
        raise SystemExit(1)
    c = data["counts"]
    print(f"source : {data['source']}")
    print(f"rows   : {c['active']} active, {c['retired']} retired "
          f"across {len(data['domains'])} domains")
    print(f"health : {c['health']}")
    print(f"state  : {c['state']}")
    print(f"life   : {c['life']}")
    if data["missing"]:
        print(f"MISSING: {', '.join(data['missing'])}")
    for d in data["domains"]:
        print(f"  {d['code'] or '??'}  {len(d['rows']):3d}  {d['title']}")
