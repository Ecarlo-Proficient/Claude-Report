"""shared/server_release.py - is the Synology ready for the code about to reach it? (10/16/2026)

The update that took the office server down: a release added a compose mount (the SharePoint approval report's
Cloud Sync folder) that nobody had made on the Synology yet. Code can be checked on a laptop; the Synology's
folders and settings cannot, unless something looks. This is THE rule for what the server needs, read from the
code itself - never a hand-kept list:

  mounts    every host path docker/compose.yml mounts or loads (volumes + env_file) must exist on the Synology
  settings  every KEY= line in docker/server.env.example must be in the Synology's server.env
  re-paste  a change to docker/update_server.sh only takes effect once its text is re-pasted in DSM

Three places use it, so they can never disagree:
  docker/update_server.sh   the same two checks in bash, run ON the Synology before any change (the last guard)
  docker/release_check.py   run from a Mac before merging the release PR (docker/RELEASE.md)
  the ledger                gear -> Office server -> Waiting for the server
From a Mac the Synology's shares are read where they are mounted (/Volumes/<share>, ACB_SYNOLOGY_VOLUMES);
what cannot be seen from here (the encrypted keys folder) is reported as "can't check from this Mac", never as ok.
"""
from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Callable, Dict, List, Optional

COMPOSE = "docker/compose.yml"
ENV_EXAMPLE = "docker/server.env.example"
PASTED_SCRIPT = "docker/update_server.sh"
SERVER_ENV_HOST = "/volume1/docker/automation/server.env"
_MOUNT_RE = re.compile(r'^\s*-\s*"?(/[^:"#]*)')        # the same pattern update_server.sh uses
_KEY_RE = re.compile(r"^([A-Z][A-Z0-9_]*)=", re.M)


def compose_paths(compose_text: str) -> List[str]:
    """Host paths compose.yml mounts or loads, in file order (a quoted path keeps its spaces)."""
    out = []
    for line in compose_text.splitlines():
        m = _MOUNT_RE.match(line)
        if m and m.group(1).rstrip() not in out:
            out.append(m.group(1).rstrip())
    return out


def env_keys(env_text: str) -> List[str]:
    """The settings an env file sets (KEY= lines; a commented line does not count)."""
    return _KEY_RE.findall(env_text or "")


def volumes_root() -> Path:
    return Path(os.getenv("ACB_SYNOLOGY_VOLUMES", "/Volumes"))


def mac_path(host_path: str, root: Optional[Path] = None) -> Optional[Path]:
    """/volume1/<share>/rest -> <root>/<share>/rest, the share as this Mac mounts it; None = not a share path."""
    parts = Path(host_path).parts
    if len(parts) < 3 or parts[1] != "volume1":
        return None
    return (root or volumes_root()).joinpath(*parts[2:])


def mac_exists(host_path: str) -> Optional[bool]:
    """True / False when this Mac can see the share; None when the share is not mounted here."""
    parts = Path(host_path).parts
    if len(parts) < 3 or parts[1] != "volume1":
        return None
    share = volumes_root() / parts[2]
    if not share.exists():
        return None
    return share.joinpath(*parts[3:]).exists()


def readiness(compose_text: str, example_text: str, server_env_text: Optional[str],
              exists: Callable[[str], Optional[bool]] = mac_exists,
              old_compose_text: str = "", old_example_text: str = "") -> List[Dict]:
    """One row per thing the code needs: {kind, what, ok (True / False / None = can't check), new}.
    `new` = it was not needed by the code the server runs (old_*), so it is this release's setup step."""
    rows = []
    old_paths, old_keys = set(compose_paths(old_compose_text)), set(env_keys(old_example_text))
    for p in compose_paths(compose_text):
        rows.append({"kind": "mount", "what": p, "ok": exists(p), "new": bool(old_compose_text) and p not in old_paths})
    have = set(env_keys(server_env_text)) if server_env_text is not None else None
    for k in env_keys(example_text):
        rows.append({"kind": "setting", "what": k, "ok": (k in have) if have is not None else None,
                     "new": bool(old_example_text) and k not in old_keys})
    return rows


def server_env_text() -> Optional[str]:
    """The Synology's server.env as this Mac sees it (paths and ids - no secrets), or None."""
    p = mac_path(SERVER_ENV_HOST)
    try:
        return p.read_text() if p else None
    except OSError:
        return None


def summary(rows: List[Dict]) -> Dict:
    missing = [r for r in rows if r["ok"] is False]
    unknown = [r for r in rows if r["ok"] is None]
    return {"ready": not missing, "missing": missing, "unknown": unknown,
            "new": [r for r in rows if r["new"]]}


# ── the developer's instructions ────────────────────────────────────────────────────────────────────────────────
# The checks are ours (the owner + Claude) before a release; the developer only carries out the Synology side. So the
# check writes his instructions: every step with its exact value, from the same files, ready to forward as-is
# (a message for someone else says it is from Claude - the owner forwards it unchanged).

def compose_notes(compose_text: str) -> Dict[str, str]:
    """{host path: the comment on its compose line} - what each folder is for."""
    out = {}
    for line in compose_text.splitlines():
        m = _MOUNT_RE.match(line)
        if m:
            out[m.group(1).rstrip()] = line.split("#", 1)[1].strip() if "#" in line else ""
    return out


def example_settings(example_text: str) -> Dict[str, Dict[str, str]]:
    """{KEY: {line, note}} from server.env.example - the exact line to add and the comment above it."""
    out, note = {}, []
    for line in (example_text or "").splitlines():
        s = line.strip()
        if s.startswith("#"):
            note.append(s.lstrip("# ").strip())
        elif _KEY_RE.match(s):
            out[_KEY_RE.match(s).group(1)] = {"line": s, "note": " ".join(note)}
            note = []
        else:
            note = []
    return out


def developer_instructions(rows: List[Dict], repaste: bool, compose_text: str, example_text: str,
                           release: str) -> str:
    """Plain steps for the developer, in the order they must happen. Empty setup = just the update."""
    notes, settings = compose_notes(compose_text), example_settings(example_text)
    steps = []
    for r in rows:
        if r["kind"] == "mount" and (r["ok"] is False or (r["ok"] is None and r["new"])):
            why = notes.get(r["what"], "")
            steps.append(f"{'Make sure this exists' if r['ok'] is None else 'Create the folder'}: {r['what']}"
                         + (f"\n   (it is: {why}; docker/README.md says how it is filled)" if why else ""))
    for r in rows:
        if r["kind"] == "setting" and r["ok"] is False:
            s = settings.get(r["what"], {"line": f"{r['what']}=", "note": ""})
            steps.append(f"Add this line to /volume1/docker/automation/server.env:\n   {s['line']}"
                         + (f"\n   ({s['note']})" if s["note"] else ""))
    if repaste:
        steps.append("Re-paste the whole text of docker/update_server.sh (from the main branch) into the DSM update task: "
                     "Control Panel > Task Scheduler > the update task > Edit > Task Settings > Run command.")
    steps.append("Run the update task (Task Scheduler > the update task > Run).")
    steps.append("Check the end of /volume1/docker/automation/update_server.log: it ends with \"office-server is "
                 "running\" and \"done\". If it says STOPPED, it lists what is still missing - set that up and Run again.")
    body = "\n".join(f"{i}. {s}" for i, s in enumerate(steps, 1))
    return ("This is from Claude (the owner's AI assistant), not from the owner directly.\n\n"
            f"Office server release: {release}. Steps on the Synology, in this order:\n\n{body}\n")
