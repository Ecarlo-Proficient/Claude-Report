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
if [ "$mode" = "live" ] && [ "$branch" != "main" ]; then
  # going live: a LIVE server runs main only, so move the checkout onto the reviewed main (same key, same pin)
  if ! { G fetch origin main && G checkout -B main origin/main; }; then
    chown -R "$owner" "$SRC"; echo "could not switch to main - nothing rebuilt, the server keeps running"; exit 1
  fi
  echo "switched $branch -> main (live runs main only)"
  branch=main
fi
if ! { G fetch origin "$branch" && G merge --ff-only "origin/$branch"; }; then
  chown -R "$owner" "$SRC"; echo "PULL FAILED - nothing rebuilt, the server keeps running"; exit 1
fi
chown -R "$owner" "$SRC"
if docker compose version >/dev/null 2>&1; then DC="docker compose"; else DC="docker-compose"; fi
if [ "$(G rev-parse HEAD)" = "$before" ]; then
  # the nightly run, most nights: no new code - no rebuild. `up -d` still applies a changed server.env (going live,
  # a new setting) and is a no-op otherwise, so the running step is never interrupted for nothing.
  echo "no new commits - nothing rebuilt"
  $DC -f docker/compose.yml up -d || { echo "RESTART FAILED - the old container keeps running"; exit 1; }
  $DC -f docker/compose.yml ps
  echo "done"; exit 0
fi
echo "commits pulled:"; G log --oneline "$before..HEAD"
$DC -f docker/compose.yml up -d --build || { echo "BUILD FAILED - the old container keeps running"; exit 1; }
for f in "/volume1/Accounting/_server-test/Bill Tracker.xlsx" "/volume1/Accounting/_server-test/Bill Tracker - compare copy.xlsx"; do
  [ -f "$f" ] || continue
  if command -v synoacltool >/dev/null 2>&1; then synoacltool -enforce-inherit "$f"; else chmod 664 "$f"; fi
  echo "permissions reset: $f"
done
$DC -f docker/compose.yml ps
echo "done"
