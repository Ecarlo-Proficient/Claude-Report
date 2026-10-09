"""The office server image carries READ code only (2026-09-30, the Synology move plan).

The server holds its own QuickBooks login, and Intuit has no read-only connection - so the guard is what goes INTO
the image: the Dockerfile's allow-list. This fails when
  * the Dockerfile copies something off the allow-list, or .dockerignore lets in something the Dockerfile doesn't list,
  * a copied path is missing from the repo,
  * a QuickBooks writer (pay bills, re-apply, recode, bulk close, reclass, anything under wip/ or one-offs/) is copied,
  * any copied .py can send a write to QuickBooks - a POST / PUT / DELETE in a file that talks to /v3/company. The one
    allowed POST is the login renewal in shared/qbo_api.py (Intuit's OAuth endpoint).
"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCKERFILE = ROOT / "docker" / "Dockerfile"
IGNORE = ROOT / "docker" / ".dockerignore"

ALLOWED = {"python-env/requirements.lock", "shared/", "bill-tracker/", "invoice-sync/", "ledger/refresh_mirror.py",
           "docker/scheduler.py"}
WRITE_CALL = re.compile(r"requests\.(post|put|patch|delete)\(|\.request\(\s*[\"'](POST|PUT|DELETE)|method\s*=\s*[\"'](POST|PUT|DELETE)")


def _copied() -> set:
    out = set()
    for line in DOCKERFILE.read_text().splitlines():
        m = re.match(r"\s*COPY\s+(\S+)\s+\S+", line)
        if m:
            out.add(m.group(1))
    return out


def _let_in() -> set:
    return {ln.strip()[1:] for ln in IGNORE.read_text().splitlines() if ln.strip().startswith("!")}


def _files(paths: set) -> list:
    got = []
    for p in paths:
        f = ROOT / p
        got += sorted(f.rglob("*.py")) if f.is_dir() else ([f] if f.suffix == ".py" else [])
    return got


def test_dockerfile_copies_only_the_allow_list():
    assert _copied() == ALLOWED, f"Dockerfile COPY list drifted from the allow-list: {_copied() ^ ALLOWED}"


def test_dockerignore_matches_the_allow_list():
    assert _let_in() == ALLOWED, f".dockerignore lets in a different set: {_let_in() ^ ALLOWED}"


def test_every_copied_path_exists():
    missing = [p for p in ALLOWED if not (ROOT / p).exists()]
    assert not missing, f"copied but not in the repo: {missing}"


def test_no_writer_is_copied():
    for p in _copied():
        assert p not in ("ledger/pay_bills.py", "ledger/reapply_check.py"), f"{p} is a QuickBooks writer"
        assert p != "ledger/", "the whole ledger folder carries the QuickBooks writers - copy files by name"
        assert not p.startswith(("wip/", "one-offs/", "debt-schedule/")), f"{p} is a folder with QuickBooks writers"


def test_no_copied_file_can_write_to_quickbooks():
    bad = []
    for f in _files(_copied()):
        src = f.read_text(errors="replace")
        if "/v3/company" not in src and "quickbooks.api.intuit.com" not in src:
            continue
        lines = src.splitlines()
        for i, ln in enumerate(lines):
            if not WRITE_CALL.search(ln):
                continue
            near = "\n".join(lines[max(0, i - 8):i + 8])
            if f.relative_to(ROOT).as_posix() == "shared/qbo_api.py" and "oauth.platform.intuit.com" in near:
                continue                               # the login renewal - the only allowed POST
            bad.append(f"{f.relative_to(ROOT)}:{i + 1}: {ln.strip()[:90]}")
    assert not bad, "a file in the office server image can write to QuickBooks:\n" + "\n".join(bad)
