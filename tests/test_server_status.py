"""The office server monitor (ledger/server_status.py, 10/09/2026): status.json + writer.json ->
ok · late · down · failing · standing down · off. Pure; no share, no network."""
import datetime as dt
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import server_status as ss  # noqa: E402

TZ = dt.timezone(dt.timedelta(hours=-5))
NOW = dt.datetime(2026, 10, 9, 10, 0, tzinfo=TZ)


def _iso(minutes_ago: float) -> str:
    return (NOW - dt.timedelta(minutes=minutes_ago)).isoformat(timespec="seconds")


def _status(written_ago=1, mirror_exit=0, ap_exit=0, mode="live", skipped=None):
    jobs = {"mirror": {"last_run": _iso(2), "last_ok": _iso(2 if mirror_exit == 0 else 5), "exit": mirror_exit, "seconds": 40, "tail": ["ok"]},
            "ap": {"last_run": _iso(7), "last_ok": _iso(7 if ap_exit == 0 else 22), "exit": ap_exit, "seconds": 90, "tail": ["Traceback", "KeyError: x"]},
            "ar": {"last_run": _iso(6), "last_ok": _iso(6), "exit": 0, "seconds": 30, "tail": []}}
    if skipped:
        jobs["apar"] = {"skipped": skipped}
    return {"mode": mode, "host": "office-server", "written": _iso(written_ago), "jobs": jobs}


def test_no_file_is_off():
    a = ss.assess(None, "server", NOW)
    assert a["present"] is False and a["health"] == "off" and a["jobs"] == []


def test_everything_clean_is_ok_with_jobs_in_order():
    a = ss.assess(_status(), "server", NOW)
    assert a["health"] == "ok" and a["mode"] == "live" and a["writer"] == "server"
    assert [j["job"] for j in a["jobs"]] == ["mirror", "ap", "ar"]
    assert a["jobs"][0]["label"] == "QuickBooks copy" and a["jobs"][0]["ok"] is True and a["jobs"][0]["ran_min_ago"] == 2.0


def test_a_failed_job_turns_failing_and_keeps_its_tail():
    a = ss.assess(_status(ap_exit=1), "server", NOW)
    assert a["health"] == "failing" and "Bill Tracker run" in a["reason"]
    ap = next(j for j in a["jobs"] if j["job"] == "ap")
    assert ap["ok"] is False and ap["exit"] == 1 and ap["tail"][-1] == "KeyError: x"


def test_quiet_file_is_late_then_down():
    assert ss.assess(_status(written_ago=25), "server", NOW)["health"] == "late"
    a = ss.assess(_status(written_ago=90), "server", NOW)
    assert a["health"] == "down" and "90 minutes" in a["reason"]


def test_down_outranks_a_failed_job_and_failing_outranks_late():
    assert ss.assess(_status(written_ago=90, ap_exit=1), "server", NOW)["health"] == "down"
    assert ss.assess(_status(written_ago=25, ap_exit=1), "server", NOW)["health"] == "failing"


def test_live_mode_without_the_server_as_writer_is_standing_down():
    a = ss.assess(_status(), "mac", NOW)
    assert a["health"] == "standing down" and '"mac"' in a["reason"]
    a = ss.assess(_status(), None, NOW)
    assert a["health"] == "standing down" and "fails closed" in a["reason"]
    assert ss.assess(_status(mode="test"), "mac", NOW)["health"] == "ok"      # test mode never writes - the writer file does not apply


def test_a_recent_apar_skip_is_standing_down_an_old_one_is_not():
    a = ss.assess(_status(skipped=f"{_iso(3)} mirror older than 30 min"), "server", NOW)
    assert a["health"] == "standing down" and a["reason"] == "AP / AR skipped: mirror older than 30 min"
    a = ss.assess(_status(skipped=f"{_iso(300)} mirror older than 30 min"), "server", NOW)
    assert a["health"] == "ok"                                                # five hours old, AP ran since


def test_read_takes_both_files_and_survives_garbage(tmp_path):
    (tmp_path / ss.STATUS_FILE).write_text(json.dumps(_status()))
    (tmp_path / ss.WRITER_FILE).write_text('{"writer": "Server"}')
    a = ss.read(tmp_path, NOW)
    assert a["present"] and a["writer"] == "server" and a["health"] == "ok" and a["dir"] == str(tmp_path)
    (tmp_path / ss.WRITER_FILE).write_text("not json")
    assert ss.read(tmp_path, NOW)["writer"] is None
    (tmp_path / ss.STATUS_FILE).write_text("[]")
    assert ss.read(tmp_path, NOW)["present"] is False
