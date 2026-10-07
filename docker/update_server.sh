#!/bin/bash
# update_server.sh - bring the office server up to date WITHOUT a terminal (owner 2026-10-07: "why the dev? we can't do it?").
#
# Run as root from DSM: Control Panel > Task Scheduler > Create > Scheduled Task > User-defined script, user root,
# script:  bash /volume1/docker/automation/update_server.sh   -> save, select it, Run.
# It: pulls the branch the checkout is on (read-only deploy key), rebuilds + restarts the container (the mode in
# server.env is unchanged - test stays test), and puts the share's inherited permissions back on the test trackers
# (files an older build locked to its own account). Everything it does is appended to
# /volume1/docker/automation/update_server.log, readable from the Mac over the docker share.
set -u
LOG=/volume1/docker/automation/update_server.log
exec >>"$LOG" 2>&1
echo
echo "===== $(date '+%m/%d/%Y %H:%M:%S') update_server"
SRC=/volume1/docker/automation/src
cd "$SRC" || { echo "no checkout at $SRC"; exit 1; }
G() { git -c safe.directory="$SRC" "$@"; }
export GIT_SSH_COMMAND="ssh -i /volume1/automation-keys/deploy_key -o IdentitiesOnly=yes -o StrictHostKeyChecking=accept-new"
branch="$(G rev-parse --abbrev-ref HEAD)"
echo "before: $(G log --oneline -1) [$branch]"
G fetch origin "$branch" && G merge --ff-only "origin/$branch" || { echo "PULL FAILED - nothing rebuilt"; exit 1; }
echo "after:  $(G log --oneline -1)"
if docker compose version >/dev/null 2>&1; then DC="docker compose"; else DC="docker-compose"; fi
$DC -f docker/compose.yml up -d --build || { echo "BUILD FAILED - the old container keeps running"; exit 1; }
for f in "/volume1/Accounting/_server-test/Bill Tracker.xlsx" "/volume1/Accounting/_server-test/Bill Tracker - compare copy.xlsx"; do
  [ -f "$f" ] || continue
  if command -v synoacltool >/dev/null 2>&1; then synoacltool -enforce-inherit "$f"; else chmod 664 "$f"; fi
  echo "permissions reset: $f"
done
$DC -f docker/compose.yml ps
echo "done"
