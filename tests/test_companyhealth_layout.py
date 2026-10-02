"""The CompanyHealth folder is organized, not a dump (the owner 2026-09-28).

Every file a tool reads or writes there goes through a NAMED folder in
shared/paths.py - register_file() / reports_dir() / analysis_dir() /
companyhealth_sources_dir(). Joining a file straight onto companyhealth_dir()
is how the root filled up, so no module outside shared/paths.py may call it."""
import datetime as dt
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from shared import paths

SKIP_DIRS = {".git", "node_modules", "__pycache__", ".venv", "venv"}


def _py_files():
    for p in ROOT.rglob("*.py"):
        if SKIP_DIRS & set(p.relative_to(ROOT).parts):
            continue
        yield p


def test_no_tool_writes_at_the_companyhealth_root():
    offenders = []
    for p in _py_files():
        rel = p.relative_to(ROOT).as_posix()
        if rel in ("shared/paths.py", "tests/test_companyhealth_layout.py"):
            continue
        for n, line in enumerate(p.read_text(errors="ignore").splitlines(), 1):
            if re.search(r"\bcompanyhealth_dir\(\)", line):
                offenders.append(f"{rel}:{n}")
    assert not offenders, ("use paths.register_file / reports_dir / analysis_dir, "
                           "never companyhealth_dir() directly: " + ", ".join(offenders))


def test_named_folders(tmp_path, monkeypatch):
    monkeypatch.setenv("ACB_COMPANYHEALTH_DIR", str(tmp_path))
    assert paths.register_file("job_rulings.json") == tmp_path / "Registers" / "job_rulings.json"
    assert paths.reports_dir() == tmp_path / "Reports"
    d = paths.analysis_dir("MFD192 Sunrise / joint check", dt.date(2026, 9, 28))
    assert d == tmp_path / "Analysis" / "MFD192 Sunrise - joint check (09-28-2026)"
    assert d.is_dir()


def test_organize_plan_sorts_known_files(tmp_path, monkeypatch):
    monkeypatch.setenv("ACB_COMPANYHEALTH_DIR", str(tmp_path))
    for n in ("job_rulings.json", "Weekly Schedule.xlsx", "MFD PnL - Internal - Director Cut.xlsx",
              "some one-off.xlsx", "Open Project Ledger.command"):
        (tmp_path / n).write_text("x")
    plan = {s.name: (t.parent.name if t else None) for s, t in paths.organize_plan()}
    assert plan == {"job_rulings.json": "Registers", "Weekly Schedule.xlsx": "Reports",
                    "MFD PnL - Internal - Director Cut.xlsx": "Reports", "some one-off.xlsx": None}
