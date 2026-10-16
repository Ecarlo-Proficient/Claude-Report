"""ledger/server_updates - what is still waiting to reach the office server, sorted by where it stands.

The rule: a commit that changes nothing the image copies never goes to the server; one on dev only waits for the
release PR; one on main waits for the next update. The server reports the commit it runs (docker/scheduler.py
build_info, stamped by update_server.sh), and the ledger's monitor passes it through.
"""
import datetime as dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "ledger"))
sys.path.insert(0, str(ROOT / "docker"))

import server_updates as U  # noqa: E402
import server_status as S  # noqa: E402

PATHS = ["shared/", "bill-tracker/", "ledger/refresh_mirror.py", "docker/scheduler.py"]


def _c(sha, files):
    return {"sha": sha, "short": sha[:7], "subject": f"commit {sha}", "date": "2026-10-16T10:00:00-05:00", "files": files}


def test_image_paths_are_the_dockerfile_allow_list():
    from test_office_server_image import ALLOWED
    assert set(U.image_paths()) == ALLOWED


def test_what_changes_the_server():
    assert U.changes_server(["shared/qbo_mirror.py"], PATHS)
    assert U.changes_server(["ledger/refresh_mirror.py"], PATHS)
    assert U.changes_server(["docker/Dockerfile"], PATHS)
    assert not U.changes_server(["ledger/dashboard.py", "ledger/static/app.js", "docs/ARCHITECTURE.md"], PATHS)
    assert not U.changes_server(["shared-notes.md"], PATHS)               # a prefix of a folder name is not the folder


def test_stages_on_a_live_server_running_main():
    commits = [_c("a" * 40, ["shared/xlsx_save.py"]), _c("b" * 40, ["bill-tracker/x.py"]),
               _c("c" * 40, ["ledger/static/app.js"]), _c("d" * 40, ["docker/update_server.sh"])]
    rows = U.classify(commits, server_branch="main", on_main={"b" * 40},
                      paths=PATHS, ci={"a" * 40: {"status": "completed", "conclusion": "success", "url": "u"},
                                       "b" * 40: {"status": "in_progress"}})
    by = {r["sha"][0]: r for r in rows}
    assert by["a"]["stage"] == "release" and by["a"]["ci"] == "passed"
    assert by["b"]["stage"] == "update" and by["b"]["ci"] == "running"
    assert by["c"]["stage"] == "none" and by["c"]["ci"] is None
    assert by["d"]["stage"] == "none" and by["d"]["repaste"] is True          # pasted in DSM, not part of the image


def test_a_server_on_dev_takes_dev_at_the_next_update():
    rows = U.classify([_c("a" * 40, ["shared/x.py"])], server_branch="dev", on_main=set(), paths=PATHS, ci={})
    assert rows[0]["stage"] == "update"


def test_the_server_reports_the_commit_it_was_built_from(monkeypatch):
    import scheduler
    monkeypatch.setenv("ACB_BUILD_COMMIT", "0b96146096")
    monkeypatch.setenv("ACB_BUILD_BRANCH", "main")
    monkeypatch.setenv("ACB_BUILT_AT", "2026-10-17T03:00:12-0500")
    assert scheduler.build_info() == {"commit": "0b96146096", "branch": "main", "built_at": "2026-10-17T03:00:12-0500"}
    monkeypatch.setenv("ACB_BUILD_COMMIT", "")
    assert scheduler.build_info()["commit"] is None                           # unstamped image: unknown, never guessed


def test_the_monitor_passes_the_build_through():
    now = dt.datetime(2026, 10, 16, 12, 0, tzinfo=dt.timezone.utc)
    st = {"mode": "live", "host": "office-server", "written": now.isoformat(),
          "build": {"commit": "abc", "branch": "main", "built_at": None}, "jobs": {}}
    assert S.assess(st, "server", now)["build"]["commit"] == "abc"
    assert S.assess({**st, "build": None}, "server", now)["build"] is None
