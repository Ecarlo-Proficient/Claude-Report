"""The office server's daily CP P&L job (docker/scheduler.py, 2026-10-13).

The owner: "set up synology pnl for cp only, to update daily once at 6am every day; the overview as the last piece".
What this pins:
  * the batch is `project-pnl/project_pnl_export.py active cp --from-folders --no-prompt` - the Active list is the
    job folders at the top of the awarded-projects folder (the WIP master is not read; owner: "the wip report has
    nothing to do with this"), every workbook first, the CP Overview LAST from those workbooks - the tool's own order;
  * test mode writes practice copies (--out under TEST_DIR) and never the Overview (--no-overview);
  * live mode runs only while the writer file says server (the rollback switch covers it);
  * a missing mount (the awarded-projects folder) = skipped + alert, never a run;
  * once a day survives a restart: the day is read back from status.json.
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _load(tmp: Path, mode: str, writer: str | None):
    env = {"ACB_SERVER_MODE": mode, "STATUS_DIR": str(tmp / "status"), "LOG_DIR": str(tmp / "logs"),
           "TEST_DIR": str(tmp / "test"), "WRITER_FILE": str(tmp / "writer.json"), "TEAMS_WEBHOOK_ALERTS": "",
           "ACB_CP_AWARDED_BASE": str(tmp / "cp")}
    os.environ.update(env)
    if writer is not None:
        (tmp / "writer.json").write_text(json.dumps({"writer": writer}))
    spec = importlib.util.spec_from_file_location(f"sched_{mode}_{writer}", ROOT / "docker" / "scheduler.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    started: list = []
    mod._start_bg = lambda name, argv, env_extra=None: started.append((name, argv)) or True
    return mod, started


def _mounts(tmp: Path):
    (tmp / "cp").mkdir()


def test_live_runs_active_cp_when_the_writer_file_says_server(tmp_path):
    _mounts(tmp_path)
    mod, started = _load(tmp_path, "live", "server")
    mod._pnl_cp()
    assert len(started) == 1
    name, argv = started[0]
    assert name == "pnl-cp"
    assert argv[1:] == ["project-pnl/project_pnl_export.py", "active", "cp", "--from-folders", "--no-prompt"]
    assert "--no-overview" not in argv            # the Overview is the last piece of the run


def test_live_stands_down_when_the_writer_file_says_mac(tmp_path):
    _mounts(tmp_path)
    mod, started = _load(tmp_path, "live", "mac")
    mod._pnl_cp()
    assert started == []
    assert "standing down" in mod._state["jobs"]["pnl-cp"]["skipped"]


def test_test_mode_writes_practice_copies_and_no_overview(tmp_path):
    _mounts(tmp_path)
    mod, started = _load(tmp_path, "test", None)
    mod._pnl_cp()
    _, argv = started[0]
    assert argv[-3:] == ["--out", str(tmp_path / "test" / "pnl"), "--no-overview"]


def test_missing_mount_skips_with_a_note(tmp_path):
    mod, started = _load(tmp_path, "live", "server")
    mod._pnl_cp()
    assert started == []
    assert "awarded-projects folder" in mod._state["jobs"]["pnl-cp"]["skipped"]


def test_once_a_day_survives_a_restart(tmp_path):
    _mounts(tmp_path)
    mod, _ = _load(tmp_path, "live", "server")
    (tmp_path / "status").mkdir()
    (tmp_path / "status" / "status.json").write_text(json.dumps(
        {"jobs": {"pnl-cp": {"last_ok": "2026-10-13T06:19:03-05:00"}}}))
    assert str(mod._last_ok_day("pnl-cp")) == "2026-10-13"
    assert mod._last_ok_day("mirror") is None


def test_the_six_oclock_default():
    spec = importlib.util.spec_from_file_location("sched_default", ROOT / "docker" / "scheduler.py")
    os.environ.pop("PNL_CP_AT", None)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert mod.PNL_CP_AT == "06:00"
    assert mod.TIMEOUTS["pnl-cp"] >= 3600


def _exporter():
    spec = importlib.util.spec_from_file_location("pnl_export", ROOT / "project-pnl" / "project_pnl_export.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_active_list_is_the_top_level_job_folders(tmp_path):
    for name in ("CP672 - 77 PAUL WILSON", "CP 785 - WAREHOUSE", "CP790-RETAIL", "Completed Projects",
                 "2022 & PREVIOUS", "Estimating templates", "~$CP999 lock"):
        (tmp_path / name).mkdir()
    (tmp_path / "Completed Projects" / "CP610 - DONE").mkdir()
    (tmp_path / "CP111 not a folder.txt").write_text("x")
    m = _exporter()
    assert m.active_cp_from_folders(tmp_path) == ["CP672", "CP785", "CP790"]


def test_from_folders_is_cp_only():
    m = _exporter()
    jobs, expanded, err = m.expand_from_folders(["active", "rp"])
    assert expanded and not jobs and "CP only" in err
    jobs, expanded, err = m.expand_from_folders(["CP672"])    # a named job passes straight through
    assert (jobs, expanded, err) == (["CP672"], False, "")
