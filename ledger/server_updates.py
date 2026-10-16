"""ledger/server_updates.py - every commit still waiting to reach the office server, and where it stands
(owner 2026-10-16: "i need to see at a glance all the pushes pending to be in synology and their status when i
click refresh in the project ledger. that way i know what is up and what is not").

How a change reaches the Synology: a commit lands on `dev` -> the release PR merges dev into `main` (one approving
review + green CI) -> the server's nightly update (docker/update_server.sh, 03:00, or run on demand in DSM) pulls
`main`, rebuilds and restarts. The update stamps the image with the commit it was built from; the scheduler reports
it in server-status.json (`build`), which the ledger already reads. This module compares that commit with GitHub:

  waiting for release   on dev only - needs the dev -> main release PR
  waiting for update    on main, not yet on the server - goes on at the next update
  not for the server    changes nothing the server runs (ledger/, wip/, docs ...) - it never goes there
plus, on any of them, `repaste` when it changes docker/update_server.sh: that script runs from text PASTED in DSM,
so it takes effect only once an admin re-pastes it (docker/README 7b).

"Changes the server" = touches a path the image copies (docker/Dockerfile's COPY lines, read from the file - the
same allow-list tests/test_office_server_image.py guards) or the Dockerfile / compose.yml themselves.

`classify()` is the pure rule (tests/test_server_updates.py); `collect()` runs git (fetch + log) and the GitHub CLI
(CI runs on dev, the open release PR). Read-only: nothing here pushes, merges or touches the server.
"""
from __future__ import annotations

import json
import re
import shutil
import subprocess
from pathlib import Path
from typing import Dict, Iterable, List, Optional

REPO = Path(__file__).resolve().parent.parent
DOCKERFILE = REPO / "docker" / "Dockerfile"
BUILD_FILES = {"docker/Dockerfile", "docker/compose.yml"}
PASTED_SCRIPT = "docker/update_server.sh"
UPDATE_TIME = "03:00"
_SEP = "\x1f"


def image_paths(dockerfile: Path = DOCKERFILE) -> List[str]:
    """The repo paths the server image copies, from the Dockerfile's COPY lines (a folder keeps its trailing /)."""
    out = []
    for line in dockerfile.read_text().splitlines():
        m = re.match(r"\s*COPY\s+(\S+)\s+\S+", line)
        if m:
            out.append(m.group(1))
    return out


def changes_server(files: Iterable[str], paths: List[str]) -> bool:
    for f in files:
        if f in BUILD_FILES:
            return True
        for p in paths:
            if (p.endswith("/") and f.startswith(p)) or f == p:
                return True
    return False


def classify(commits: List[dict], *, server_branch: Optional[str], on_main: set,
             paths: List[str], ci: Dict[str, dict]) -> List[dict]:
    """commits = newest first, each {sha, short, subject, date, files}. Returns the rows the ledger shows."""
    rows = []
    for c in commits:
        files = c.get("files") or []
        server = changes_server(files, paths)
        if not server:
            stage, word = "none", "Never goes"
        elif server_branch == "dev" or c["sha"] in on_main:
            stage, word = "update", f"Next update ({UPDATE_TIME})"
        else:
            stage, word = "release", "Needs release"
        run = ci.get(c["sha"]) or {}
        rows.append({"sha": c["sha"], "short": c["short"], "subject": c["subject"], "date": c["date"],
                     "stage": stage, "word": word, "changes_server": server, "repaste": PASTED_SCRIPT in files,
                     "ci": _ci_word(run), "ci_url": run.get("url")})
    return rows


def _ci_word(run: dict) -> Optional[str]:
    if not run:
        return None
    if run.get("status") != "completed":
        return "running"
    return {"success": "passed", "failure": "failed", "cancelled": "cancelled"}.get(run.get("conclusion"), run.get("conclusion"))


def _git(*args: str, timeout: int = 30) -> str:
    p = subprocess.run(["git", *args], cwd=str(REPO), capture_output=True, text=True, timeout=timeout)
    if p.returncode != 0:
        raise RuntimeError((p.stderr or p.stdout or "git failed").strip().splitlines()[-1])
    return p.stdout


def _log(rng: str) -> List[dict]:
    """Commits in `rng` (newest first, merges left out - a release merge brings no change of its own), with files."""
    out, cur = [], None
    for line in _git("log", "--no-merges", "--name-only", f"--format=@@{_SEP}%H{_SEP}%h{_SEP}%cI{_SEP}%s", rng).splitlines():
        if line.startswith("@@" + _SEP):
            _, sha, short, date, subject = line.split(_SEP, 4)
            cur = {"sha": sha, "short": short, "date": date, "subject": subject, "files": []}
            out.append(cur)
        elif line.strip() and cur is not None:
            cur["files"].append(line.strip())
    return out


def _gh() -> Optional[str]:
    return shutil.which("gh") or next((p for p in ("/opt/homebrew/bin/gh", "/usr/local/bin/gh") if Path(p).exists()), None)


def _gh_json(gh: str, *args: str):
    p = subprocess.run([gh, *args], cwd=str(REPO), capture_output=True, text=True, timeout=30)
    if p.returncode != 0:
        raise RuntimeError((p.stderr or "gh failed").strip().splitlines()[-1])
    return json.loads(p.stdout or "null")


def repo_url() -> Optional[str]:
    """origin as a browsable https URL (commit links), or None."""
    try:
        url = _git("remote", "get-url", "origin").strip()
    except (RuntimeError, subprocess.TimeoutExpired):
        return None
    m = re.match(r"(?:https://|git@)github\.com[/:](.+?)(?:\.git)?$", url)
    return f"https://github.com/{m.group(1)}" if m else None


def collect(build: Optional[dict], fetch: bool = True) -> dict:
    """Everything the ledger's 'Waiting for the server' block shows. `build` = server-status.json's build stamp."""
    build = build or {}
    server_sha, server_branch = build.get("commit"), build.get("branch")
    notes: List[str] = []
    fetched = False
    if fetch:
        try:
            _git("fetch", "--quiet", "origin", "main", "dev", timeout=45)
            fetched = True
        except (RuntimeError, subprocess.TimeoutExpired) as e:
            notes.append(f"could not reach GitHub ({e}) - showing what this Mac last fetched")

    server_known = False
    if server_sha:
        try:
            _git("cat-file", "-e", f"{server_sha}^{{commit}}")
            server_known = True
        except RuntimeError:
            notes.append(f"the server runs {server_sha[:7]}, which this Mac has not fetched")
    else:
        notes.append("Server version not reported yet - it starts once the DSM update script is re-pasted and the "
                     "next update rebuilds. Until then only commits not released to main are listed.")

    base = server_sha if server_known else "origin/main"
    try:
        commits = _log(f"{base}..origin/dev")
        on_main = {c["sha"] for c in _log(f"{base}..origin/main")} if server_known else set()
        if server_known:              # a commit on main that dev does not have (never expected) still shows
            have = {c["sha"] for c in commits}
            commits += [c for c in _log(f"{base}..origin/main") if c["sha"] not in have]
    except (RuntimeError, subprocess.TimeoutExpired) as e:
        return {"ok": False, "error": f"git: {e}", "server": build, "notes": notes}

    ci: Dict[str, dict] = {}
    release = None
    gh = _gh()
    if not gh:
        notes.append("GitHub CLI (gh) not found on this Mac - no CI or release PR status")
    else:
        try:
            for r in _gh_json(gh, "run", "list", "--branch", "dev", "--workflow", "CI", "--limit", "100",
                              "--json", "headSha,status,conclusion,url") or []:
                ci.setdefault(r["headSha"], r)
            prs = _gh_json(gh, "pr", "list", "--base", "main", "--head", "dev", "--state", "open",
                           "--json", "number,url,reviewDecision,isDraft") or []
            release = prs[0] if prs else None
        except (RuntimeError, subprocess.TimeoutExpired, ValueError) as e:
            notes.append(f"GitHub status unavailable ({e})")

    rows = classify(commits, server_branch=server_branch, on_main=on_main,
                    paths=image_paths(), ci=ci)
    counts = {k: sum(1 for r in rows if r["stage"] == k) for k in ("release", "update", "none")}
    return {"ok": True, "fetched": fetched, "repo_url": repo_url(), "server": build, "server_known": server_known, "rows": rows,
            "counts": counts, "release_pr": release, "repaste": any(r["repaste"] for r in rows), "notes": notes}
