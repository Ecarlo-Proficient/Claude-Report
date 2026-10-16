"""docker/update_server.sh - the preflight that stops an update before it touches the server (10/16/2026: a new
compose mount that did not exist yet took the server down, and the log said the old container kept running).

The test runs the script's OWN `preflight` function (cut out of the file, never a copy) in bash against a throwaway
git repo whose compose.yml mounts temp folders: a missing mount always stops; a server.env setting the example lists
but the Synology lacks stops a release and is only reported on a restart of the same code.
"""
import re
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "docker" / "update_server.sh"

pytestmark = pytest.mark.skipif(not shutil.which("bash") or not shutil.which("git"), reason="needs bash + git")


def _function(name: str) -> str:
    text = SCRIPT.read_text()
    m = re.search(rf"^{name}\(\) \{{\n.*?^\}}\n", text, re.S | re.M)
    assert m, f"{name}() not found in update_server.sh"
    return m.group(0)


def _repo(tmp: Path, mounts, example_keys) -> Path:
    repo = tmp / "src"
    (repo / "docker").mkdir(parents=True)
    lines = ["services:", "  office-server:", "    env_file:", f"      - {tmp / 'server.env'}    # paths",
             "    volumes:"] + [f'      - "{m}:/mnt/x{i}"   # note' for i, m in enumerate(mounts)]
    (repo / "docker" / "compose.yml").write_text("\n".join(lines) + "\n")
    (repo / "docker" / "server.env.example").write_text(
        "# comment\n" + "".join(f"{k}=value\n" for k in example_keys) + "# OPTIONAL_COMMENTED=1\n")
    for args in (["init", "-q"], ["add", "-A"], ["-c", "user.name=t", "-c", "user.email=t@t", "commit", "-qm", "c"]):
        subprocess.run(["git", *args], cwd=repo, check=True)
    return repo


def _run(tmp: Path, repo: Path, kind: str):
    harness = f'set -uo pipefail\nBASE="{tmp}"\nG() {{ git -C "{repo}" "$@"; }}\n{_function("preflight")}preflight HEAD {kind}\n'
    p = subprocess.run(["bash", "-c", harness], capture_output=True, text=True)
    return p.returncode, p.stdout + p.stderr


def test_everything_in_place_passes(tmp_path):
    (tmp_path / "with space folder").mkdir()
    (tmp_path / "server.env").write_text("A=1\nB=2\n")
    repo = _repo(tmp_path, [tmp_path / "with space folder"], ["A", "B"])
    rc, out = _run(tmp_path, repo, "release")
    assert rc == 0 and "preflight ok" in out


def test_a_missing_mount_stops_before_any_change(tmp_path):
    (tmp_path / "server.env").write_text("A=1\n")
    missing = tmp_path / "bill-approval"
    repo = _repo(tmp_path, [missing], ["A"])
    for kind in ("release", "restart"):
        rc, out = _run(tmp_path, repo, kind)
        assert rc == 1, out
        assert f"MISSING on the Synology: {missing}" in out and "STOPPED before any change" in out


def test_a_missing_setting_stops_a_release_and_warns_on_a_restart(tmp_path):
    (tmp_path / "server.env").write_text("A=1\n")
    repo = _repo(tmp_path, [], ["A", "ACB_BILL_APPROVAL_XLSX"])
    rc, out = _run(tmp_path, repo, "release")
    assert rc == 1 and "MISSING in server.env: ACB_BILL_APPROVAL_XLSX=" in out
    rc, out = _run(tmp_path, repo, "restart")
    assert rc == 0 and "WARNING - the server runs without" in out and "ACB_BILL_APPROVAL_XLSX" in out


def test_the_real_compose_file_names_every_mount():
    """The pattern the script uses finds each host path in the real compose.yml, spaces kept, nothing extra."""
    pat = re.compile(r'^\s*-\s*"?(/[^:"#]*)')
    found = [m.group(1).rstrip() for m in map(pat.match, (ROOT / "docker" / "compose.yml").read_text().splitlines()) if m]
    assert "/volume1/docker/automation/bill-approval" in found
    assert "/volume1/Common/CURRENT PROJECTS/Awarded Projects Commercial projects" in found
    assert all(not p.endswith(" ") for p in found) and len(found) == len(set(found))


# ── the whole script, end to end: this morning's incident replayed ──────────────────────────────────────────────
# A copy of update_server.sh with only its two location lines pointed at a temp folder and a stub `docker` / `sleep`
# first on PATH; a local bare repo plays GitHub. The stub records every docker call.

STUB = r'''#!/bin/bash
echo "$*" >> "$STUB_LOG"
case "$*" in
  "compose version") exit 0 ;;
  inspect*) echo "${STUB_RUNNING:-true}"; exit 0 ;;
  *" up "*|*" up") exit "${STUB_UP_RC:-0}" ;;
esac
exit 0
'''


def _world(tmp: Path, mounts_exist: bool):
    base = tmp / "automation"
    base.mkdir()
    stub_dir = tmp / "bin"
    stub_dir.mkdir()
    for name, body in (("docker", STUB), ("sleep", "#!/bin/bash\nexit 0\n")):
        (stub_dir / name).write_text(body)
        (stub_dir / name).chmod(0o755)
    (base / "server.env").write_text("ACB_SERVER_MODE=live\nA=1\n")
    mount = tmp / "bill-approval"
    if mounts_exist:
        mount.mkdir()
    origin, work = tmp / "origin.git", tmp / "seed"
    subprocess.run(["git", "init", "-q", "--bare", "-b", "main", str(origin)], check=True)
    subprocess.run(["git", "clone", "-q", str(origin), str(work)], check=True)
    g = lambda *a: subprocess.run(["git", "-C", str(work), "-c", "user.name=t", "-c", "user.email=t@t", *a], check=True)
    (work / "docker").mkdir()
    (work / "docker" / "server.env.example").write_text("ACB_SERVER_MODE=test\nA=1\n")
    (work / "docker" / "compose.yml").write_text(f"services:\n  s:\n    env_file:\n      - {base / 'server.env'}\n")
    g("add", "-A"); g("commit", "-qm", "first"); g("push", "-q", "origin", "HEAD:main")
    subprocess.run(["git", "clone", "-q", "-b", "main", str(origin), str(base / "src")], check=True)
    # the release: compose gains the approval-report mount
    (work / "docker" / "compose.yml").write_text(
        f"services:\n  s:\n    env_file:\n      - {base / 'server.env'}\n    volumes:\n      - {mount}:/mnt/bill-approval:ro\n")
    g("commit", "-qam", "add the bill-approval mount"); g("push", "-q", "origin", "HEAD:main")
    script = SCRIPT.read_text()
    script = re.sub(r"^export PATH=.*$", f'export PATH="{stub_dir}:$PATH"', script, flags=re.M)
    script = re.sub(r"^BASE=.*$", f'BASE="{base}"', script, flags=re.M)
    copy = tmp / "update_server.sh"
    copy.write_text(script)
    return base, copy


def _update(tmp: Path, base: Path, copy: Path, **env):
    import os
    e = {**os.environ, "STUB_LOG": str(tmp / "docker.calls"), "GIT_SSH_COMMAND": "", **env}
    rc = subprocess.run(["bash", str(copy)], env=e).returncode
    calls = (tmp / "docker.calls").read_text() if (tmp / "docker.calls").exists() else ""
    return rc, (base / "update_server.log").read_text(), calls


def _head(base: Path) -> str:
    return subprocess.run(["git", "-C", str(base / "src"), "log", "-1", "--format=%s"],
                          capture_output=True, text=True).stdout.strip()


def test_the_10_16_incident_now_stops_and_leaves_the_server_alone(tmp_path):
    base, copy = _world(tmp_path, mounts_exist=False)
    rc, log, calls = _update(tmp_path, base, copy)
    assert rc == 1
    assert "STOPPED before any change" in log and "bill-approval" in log
    assert _head(base) == "first"                                   # nothing pulled
    assert " up" not in calls                                        # the container never touched


def test_with_the_folder_in_place_the_update_goes_through(tmp_path):
    base, copy = _world(tmp_path, mounts_exist=True)
    rc, log, calls = _update(tmp_path, base, copy)
    assert rc == 0, log
    assert "preflight ok" in log and "office-server is running" in log and log.rstrip().endswith("done")
    assert _head(base) == "add the bill-approval mount"
    assert "up -d --build" in calls


def test_a_failed_start_says_the_server_is_down(tmp_path):
    base, copy = _world(tmp_path, mounts_exist=True)
    rc, log, _ = _update(tmp_path, base, copy, STUB_UP_RC="1", STUB_RUNNING="false")
    assert rc == 1
    assert "THE SERVER IS DOWN" in log and "keeps running" not in log.split("BUILD / START FAILED", 1)[-1]
