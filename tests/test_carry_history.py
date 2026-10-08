"""docker/carry_history.py: the server's miscode history continues the Mac's at switch-over."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "docker"))

import carry_history as ch  # noqa: E402


def _e(first, last, status="open", runs=1):
    return {"first_seen": first, "last_seen": last, "status": status, "runs_seen": runs}


def test_mac_record_wins_and_server_newest_state_kept():
    mac = {"schema": 1, "runs": [{"date": "2026-09-01"}],
           "entries": {"a": _e("2026-09-01", "2026-10-07", runs=40), "b": _e("2026-09-03", "2026-10-07")}}
    server = {"schema": 1, "runs": [{"date": "2026-10-05"}] * 120,
              "entries": {"a": _e("2026-10-05", "2026-10-08", status="fixed", runs=300),
                          "c": _e("2026-10-06", "2026-10-08")}}
    out = ch.merge(mac, server)
    assert out["runs"] == mac["runs"]                         # test-week runs dropped
    assert out["entries"]["a"]["first_seen"] == "2026-09-01"  # the long record's start
    assert out["entries"]["a"]["status"] == "fixed"           # the newest state
    assert out["entries"]["a"]["last_seen"] == "2026-10-08"
    assert out["entries"]["a"]["runs_seen"] == 40             # the Mac's count continues
    assert set(out["entries"]) == {"a", "b", "c"}             # nothing dropped


def test_older_server_state_does_not_overwrite():
    mac = {"entries": {"a": _e("2026-09-01", "2026-10-08", status="fixed")}, "runs": []}
    server = {"entries": {"a": _e("2026-10-05", "2026-10-07", status="open")}, "runs": []}
    assert ch.merge(mac, server)["entries"]["a"]["status"] == "fixed"
