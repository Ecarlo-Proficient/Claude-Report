#!/usr/bin/env python3
"""docker/release_check.py - is the Synology ready for the release? Run it before merging the release PR
(docker/RELEASE.md step 1). Read-only: it fetches from GitHub and looks at the Synology's shares from this Mac.

    "$ACB_PY" docker/release_check.py            # what is on dev, against what the server runs
    "$ACB_PY" docker/release_check.py --to main  # what is on main (the next nightly update)

Prints every folder and setting the code needs (shared/server_release - the same rule update_server.sh applies on
the Synology), marks what this release adds, and exits 1 when something is missing - so a release that would stop
on the Synology is caught here first. "can't check from this Mac" = a share not mounted here (the encrypted keys
folder never is). It ends with the developer's instructions for this release, ready to forward as-is: the checks
are ours (the owner + Claude), the developer carries out the Synology steps (docker/RELEASE.md).
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
from shared import paths, server_release as R  # noqa: E402

WORD = {True: "ok", False: "MISSING", None: "can't check from this Mac"}


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=str(REPO), capture_output=True, text=True, check=True).stdout


def _server_commit() -> str | None:
    try:
        s = json.loads((paths.accounting_base() / "_automation" / "server-status.json").read_text())
        return ((s.get("build") or {}).get("commit")) or None
    except (OSError, ValueError):
        return None


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--to", default="dev", choices=("dev", "main"), help="the branch about to reach the server")
    a = ap.parse_args()
    _git("fetch", "--quiet", "origin", "main", "dev")
    target = f"origin/{a.to}"
    base = _server_commit()
    base_word = f"the server's {base[:7]}" if base else "origin/main (the server has not reported its version)"
    base = base or "origin/main"
    show = lambda ref, p: subprocess.run(["git", "show", f"{ref}:{p}"], cwd=str(REPO), capture_output=True, text=True).stdout
    compose, example = show(target, R.COMPOSE), show(target, R.ENV_EXAMPLE)
    rows = R.readiness(compose, example, R.server_env_text(),
                       old_compose_text=show(base, R.COMPOSE), old_example_text=show(base, R.ENV_EXAMPLE))
    changed = _git("diff", "--name-only", f"{base}..{target}").split()
    message = R.developer_instructions(rows, R.PASTED_SCRIPT in changed, compose, example,
                                       f"{a.to} {_git('rev-parse', '--short', target).strip()}")
    s = R.summary(rows)
    print(f"Release check: {target} against {base_word}\n")
    if s["new"]:
        print("New in this release - set up on the Synology BEFORE merging:")
        for r in s["new"]:
            print(f"  [{WORD[r['ok']]}] {r['kind']}: {r['what']}")
        print()
    if R.PASTED_SCRIPT in changed:
        print(f"  [to do] {R.PASTED_SCRIPT} changed - re-paste its text into the DSM update task\n")
    others = [r for r in s["missing"] + s["unknown"] if not r["new"]]
    if others:
        print("Also needed and not confirmed:")
        for r in others:
            print(f"  [{WORD[r['ok']]}] {r['kind']}: {r['what']}")
        print()
    if s["missing"]:
        print(f"NOT READY - {len(s['missing'])} missing. The update would stop on the Synology (nothing changed) "
              f"until they are there - the steps below set them up.")
    else:
        print("READY - every folder and setting this Mac can see is on the Synology.")
    print("\n---- for the developer (forward as-is) ----\n")
    print(message)
    return 1 if s["missing"] else 0


if __name__ == "__main__":
    sys.exit(main())
