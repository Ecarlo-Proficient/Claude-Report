#!/bin/bash
# update_server.sh - bring the office server up to date from DSM, no terminal (owner 2026-10-07).
#
# SECURITY: this runs as root. Keep its text PASTED in the DSM task (Control Panel > Task Scheduler > User-defined
# script, user root) - only DSM admins can edit that. NEVER point the task at a file inside the checkout: anyone who
# can push to the repo would then run code as root on the Synology. When this file changes, re-paste it.
#
# What it does: pulls the checkout's branch (read-only deploy key, GitHub's host key PINNED), rebuilds + restarts the
# container (the mode in server.env is untouched), puts the share's inherited permissions back on the test trackers,
# and appends everything - including every commit pulled - to /volume1/docker/automation/update_server.log.
# A live server runs main only: in live mode it moves a checkout on another branch onto main itself.
# It refuses: a checkout with local edits, a second run at the same time.
# The DSM task runs it every night (03:00) and on demand: with no new commits it rebuilds nothing (it only applies a
# changed server.env), so once live, whatever is merged to main reaches the server by the next morning, hands-free.
#
# PREFLIGHT (10/16/2026 - the update that took the server down: a new compose mount, /volume1/docker/automation/
# bill-approval, did not exist yet, Docker removed the old container and could not start the new one, and the log
# said "the old container keeps running"). Before anything changes it reads the NEW commit's compose.yml and
# server.env.example: every folder / file compose mounts or loads must exist, and server.env must carry every setting
# the example lists. Anything missing = STOPPED: nothing pulled, nothing rebuilt, the running server untouched, and
# the log names each missing item. After every start it checks the container really runs and says DOWN if not; a failed
# `up` logs `ps -a` and exits 1 even when the old container still runs (the update did not go in).
# The checklist the admin follows is docker/RELEASE.md.
set -uo pipefail
export PATH=/usr/local/bin:/usr/bin:/bin:/usr/local/sbin:/usr/sbin:/sbin:/usr/syno/bin:/usr/syno/sbin
BASE=/volume1/docker/automation
SRC="$BASE/src"
exec >>"$BASE/update_server.log" 2>&1
echo
echo "===== $(date '+%m/%d/%Y %H:%M:%S') update_server"
if command -v flock >/dev/null 2>&1; then
  exec 9>/tmp/update_server.lock
  flock -n 9 || { echo "another update is running - nothing changed"; exit 1; }
fi
command -v git >/dev/null 2>&1 || { echo "git not found - nothing changed"; exit 1; }
cd "$SRC" || { echo "no checkout at $SRC - nothing changed"; exit 1; }
G() { git -c safe.directory="$SRC" "$@"; }
owner="$(stat -c %u:%g "$SRC")"
KH="$(mktemp)"; trap 'rm -f "$KH"' EXIT
# GitHub's published ed25519 host key (api.github.com/meta; SHA256:+DiY3wvvV6TuJJhbpZisF/zLDA0zPMSvHdkr4UvCOqU)
# HostKeyAlgorithms=ssh-ed25519: DSM's ssh asks for ECDSA first, which then fails against this pin (10/08 first run).
echo "github.com ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAIOMqqnkVzrm0SdG6UOoqKLsabgH5C9okWi0dh2l9GKJl" > "$KH"
export GIT_SSH_COMMAND="ssh -i /volume1/automation-keys/deploy_key -o IdentitiesOnly=yes -o HostKeyAlgorithms=ssh-ed25519 -o StrictHostKeyChecking=yes -o UserKnownHostsFile=$KH"
branch="$(G rev-parse --abbrev-ref HEAD)"
mode="$(sed -n 's/^ACB_SERVER_MODE=//p' "$BASE/server.env" | tr -d '[:space:]')"
echo "mode ${mode:-?}, branch $branch, before $(G log --oneline -1)"
if [ -n "$(G status --porcelain --untracked-files=no)" ]; then echo "the checkout has local edits - nothing changed"; G status --short; exit 1; fi
before="$(G rev-parse HEAD)"
if docker compose version >/dev/null 2>&1; then DC="docker compose"; else DC="docker-compose"; fi

# what the Synology must already have for the code at $1: compose's mounts + env files exist, server.env carries every
# setting server.env.example lists. $2 = "release" (new code: a missing setting stops too) or "restart" (same code:
# only a missing mount stops - the container cannot start without it; a missing setting is reported).
preflight() {
  local ref="$1" kind="$2" hard="" soft="" p k
  while IFS= read -r p; do
    [ -n "$p" ] && [ ! -e "$p" ] && hard="$hard
    MISSING on the Synology: $p"
  done <<PATHS
$(G show "$ref:docker/compose.yml" | sed -nE 's/^[[:space:]]*-[[:space:]]*"?(\/[^:"#]*).*/\1/p' | sed -E 's/[[:space:]]+$//')
PATHS
  for k in $(G show "$ref:docker/server.env.example" | sed -nE 's/^([A-Z][A-Z0-9_]*)=.*/\1/p'); do
    grep -qE "^$k=" "$BASE/server.env" || soft="$soft
    MISSING in server.env: $k= (the line and its note are in docker/server.env.example)"
  done
  [ "$kind" = "release" ] && { hard="$hard$soft"; soft=""; }
  [ -n "$soft" ] && echo "WARNING - the server runs without:$soft"
  if [ -n "$hard" ]; then
    echo "STOPPED before any change - the Synology is not ready for $(G log --oneline -1 "$ref" | cut -c1-80):$hard"
    echo "Set these up (docker/RELEASE.md), then Run the task again. The server keeps running what it runs now."
    return 1
  fi
  echo "preflight ok - every mount and setting this code needs is on the Synology"
}

# after a start: is the container really running? (a failed `up` can leave it removed - say so, never guess)
running_or_down() {
  sleep 5
  if [ "$(docker inspect -f '{{.State.Running}}' office-server 2>/dev/null)" = "true" ]; then
    echo "office-server is running"; return 0
  fi
  echo "THE SERVER IS DOWN - office-server is not running. Fix the error above and Run the task again."; return 1
}

# `up` itself returned an error: log what Docker really has and never end on "done" - even when a container still runs
# (a build error leaves the old one up), the update did not go in and the task must show it failed.
up_failed() {
  echo "$1 FAILED - what Docker has now:"
  $DC -f docker/compose.yml ps -a
  running_or_down && echo "the update did NOT go in - office-server still runs what it ran before. Fix the error above and Run the task again."
  exit 1
}

target="$branch"
[ "$mode" = "live" ] && target=main          # a LIVE server runs main only
if ! G fetch origin "$target"; then
  chown -R "$owner" "$SRC"; echo "FETCH FAILED - nothing changed, the server keeps running"; exit 1
fi
if [ "$(G rev-parse "origin/$target")" != "$before" ]; then
  preflight "origin/$target" release || { chown -R "$owner" "$SRC"; exit 1; }
fi
if [ "$branch" != "$target" ]; then
  # going live: move the checkout onto the reviewed main (same key, same pin)
  if ! G checkout -B "$target" "origin/$target"; then
    chown -R "$owner" "$SRC"; echo "could not switch to $target - nothing rebuilt, the server keeps running"; exit 1
  fi
  echo "switched $branch -> $target (live runs main only)"
  branch="$target"
elif ! G merge --ff-only "origin/$branch"; then
  chown -R "$owner" "$SRC"; echo "PULL FAILED - nothing rebuilt, the server keeps running"; exit 1
fi
chown -R "$owner" "$SRC"
if [ "$(G rev-parse HEAD)" = "$before" ]; then
  # the nightly run, most nights: no new code - no rebuild. `up -d` still applies a changed server.env (going live,
  # a new setting) and starts a container that is down; otherwise a no-op, the running step is never interrupted.
  echo "no new commits - nothing rebuilt"
  preflight HEAD restart || exit 1
  $DC -f docker/compose.yml up -d || up_failed RESTART
  running_or_down || exit 1
  $DC -f docker/compose.yml ps
  echo "done"; exit 0
fi
echo "commits pulled:"; G log --oneline "$before..HEAD"
# stamp the image with the code it is built from: the scheduler reports it, the ledger lists what is still waiting
export ACB_BUILD_COMMIT="$(G rev-parse HEAD)" ACB_BUILD_BRANCH="$branch" ACB_BUILT_AT="$(date '+%Y-%m-%dT%H:%M:%S%z')"
$DC -f docker/compose.yml up -d --build || up_failed "BUILD / START"
running_or_down || exit 1
for f in "/volume1/Accounting/_server-test/Bill Tracker.xlsx" "/volume1/Accounting/_server-test/Bill Tracker - compare copy.xlsx"; do
  [ -f "$f" ] || continue
  if command -v synoacltool >/dev/null 2>&1; then synoacltool -enforce-inherit "$f"; else chmod 664 "$f"; fi
  echo "permissions reset: $f"
done
$DC -f docker/compose.yml ps
echo "done"
