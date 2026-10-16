"""shared/server_release - what the Synology must have before a release reaches it, read from compose.yml and
server.env.example. The Mac-side check (ledger, release_check.py) and update_server.sh on the Synology must agree."""
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from shared import server_release as R  # noqa: E402

COMPOSE = '''services:
  s:
    env_file:
      - /volume1/docker/automation/server.env    # paths
    volumes:
      - /volume1/Accounting:/mnt/accounting
      - "/volume1/Common/CURRENT PROJECTS/Awarded:/mnt/cp"  # read + write
      - /volume1/docker/automation/bill-approval:/mnt/bill-approval:ro
'''


def test_compose_paths_and_env_keys():
    assert R.compose_paths(COMPOSE) == ["/volume1/docker/automation/server.env", "/volume1/Accounting",
                                        "/volume1/Common/CURRENT PROJECTS/Awarded",
                                        "/volume1/docker/automation/bill-approval"]
    assert R.env_keys("# c\nA=1\n# B=2\nC_D=\n") == ["A", "C_D"]


def test_mac_path_maps_a_share():
    assert R.mac_path("/volume1/docker/automation/x", Path("/Volumes")) == Path("/Volumes/docker/automation/x")
    assert R.mac_path("/data/x") is None


def test_readiness_marks_missing_new_and_unknown():
    old = COMPOSE.replace("      - /volume1/docker/automation/bill-approval:/mnt/bill-approval:ro\n", "")
    seen = {"/volume1/docker/automation/server.env": True, "/volume1/Accounting": True,
            "/volume1/Common/CURRENT PROJECTS/Awarded": None, "/volume1/docker/automation/bill-approval": False}
    rows = R.readiness(COMPOSE, "A=1\nACB_BILL_APPROVAL_XLSX=x\n", "A=1\n", exists=seen.get,
                       old_compose_text=old, old_example_text="A=1\n")
    s = R.summary(rows)
    assert not s["ready"]
    assert {r["what"] for r in s["missing"]} == {"/volume1/docker/automation/bill-approval", "ACB_BILL_APPROVAL_XLSX"}
    assert {r["what"] for r in s["new"]} == {"/volume1/docker/automation/bill-approval", "ACB_BILL_APPROVAL_XLSX"}
    assert [r["what"] for r in s["unknown"]] == ["/volume1/Common/CURRENT PROJECTS/Awarded"]
    assert R.summary(R.readiness(COMPOSE, "A=1\n", None, exists=lambda p: True))["unknown"][0]["what"] == "A"


def test_the_mac_check_and_the_synology_script_read_the_same_folders():
    """update_server.sh's sed and server_release's regex must find the same paths in the real compose.yml."""
    script = (ROOT / "docker" / "update_server.sh").read_text()
    sed_expr = re.search(r"sed -nE '(s/\^\[\[:space:\]\]\*-.*?/p)'", script).group(1)
    out = subprocess.run(f"sed -nE '{sed_expr}' docker/compose.yml | sed -E 's/[[:space:]]+$//'", shell=True,
                         cwd=ROOT, capture_output=True, text=True, check=True).stdout.splitlines()
    assert out == R.compose_paths((ROOT / "docker" / "compose.yml").read_text())


def test_developer_instructions_carry_exact_values_in_order():
    example = "A=1\n# QuickBooks' approval report (README 2c)\nACB_BILL_APPROVAL_XLSX=/mnt/bill-approval/x.xlsx\n"
    seen = {"/volume1/docker/automation/server.env": True, "/volume1/Accounting": True,
            "/volume1/Common/CURRENT PROJECTS/Awarded": True, "/volume1/docker/automation/bill-approval": False}
    rows = R.readiness(COMPOSE, example, "A=1\n", exists=seen.get)
    msg = R.developer_instructions(rows, True, COMPOSE, example, "dev -> main abc1234")
    assert msg.startswith("This is from Claude")                                  # forwarded as-is
    lines = msg.splitlines()
    i_folder = next(i for i, l in enumerate(lines) if "Create the folder: /volume1/docker/automation/bill-approval" in l)
    i_setting = next(i for i, l in enumerate(lines) if l.strip() == "ACB_BILL_APPROVAL_XLSX=/mnt/bill-approval/x.xlsx")
    i_paste = next(i for i, l in enumerate(lines) if "Re-paste" in l)
    i_run = next(i for i, l in enumerate(lines) if "Run the update task" in l)
    assert i_folder < i_setting < i_paste < i_run                                  # set up, re-paste, then run
    assert "approval report (README 2c)" in msg


def test_nothing_to_set_up_is_just_the_update():
    rows = R.readiness("", "A=1\n", "A=1\n", exists=lambda p: True)
    msg = R.developer_instructions(rows, False, "", "A=1\n", "main abc1234")
    assert "1. Run the update task" in msg and "Add this line" not in msg and "Re-paste" not in msg
