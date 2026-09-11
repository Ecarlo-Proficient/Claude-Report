#!/usr/bin/env bash
# worktree.sh - one session = one worktree (the multi-session protocol,
# 2026-08-27). Parallel sessions each get their own checkout + branch against
# the SAME local repo: no shared-tree contention, shared hooks, and git itself
# refuses one branch in two worktrees. machine.env is inherited from the main
# clone automatically (shared/paths.py resolves it via the git common dir).
#
#   new <topic>   fresh worktree ../<repo>-wt-<topic> on branch wt/<topic>,
#                 cut from up-to-date origin/dev. Prints the cd line.
#   done          run INSIDE a worktree when the work is committed: rebases
#                 onto fresh origin/dev, pushes dev (the pre-push hook gates
#                 the pushed commit), retries the push race, then removes the
#                 worktree and its branch. The flawless-landing command.
#   list          worktrees + wt/* branches.
#
# Acceptance-tested end to end 2026-08-27: worktree created while the main
# tree sat dirty with another session's work - paths.py resolved the main
# machine.env, preflight ran green in isolation, and this very commit landed
# on dev from the worktree through `done` (rebase, gated push, cleanup).
set -uo pipefail

die() { echo "worktree.sh: $*" >&2; exit 1; }

# Resolve the MAIN clone root from anywhere (main or a linked worktree).
common="$(git rev-parse --path-format=absolute --git-common-dir 2>/dev/null)" \
  || die "not inside a git repository"
main_root="$(dirname "$common")"

cmd="${1:-}"
case "$cmd" in

new)
  topic="${2:-}"
  [ -n "$topic" ] || die "usage: worktree.sh new <topic>   (topic: a-z 0-9 -)"
  case "$topic" in
    *[!a-z0-9-]*) die "topic must be lowercase a-z, 0-9, and - only" ;;
  esac
  wt_path="$(dirname "$main_root")/$(basename "$main_root")-wt-$topic"
  [ ! -e "$wt_path" ] || die "already exists: $wt_path"
  branch="wt/$topic"
  git -C "$main_root" fetch origin --quiet || die "git fetch failed"
  if git -C "$main_root" show-ref --verify --quiet "refs/heads/$branch"; then
    die "branch $branch already exists - pick another topic or finish that one"
  fi
  git -C "$main_root" worktree add --quiet -b "$branch" "$wt_path" origin/dev \
    || die "worktree add failed"
  echo "worktree ready on $branch (cut from origin/dev):"
  echo
  echo "cd \"$wt_path\""
  echo
  echo "Start the session there. Land it when committed with:"
  echo "  bash .github/worktree.sh done"
  ;;

done)
  # Must be inside a linked worktree on a wt/* branch - never the main clone.
  here="$(git rev-parse --show-toplevel 2>/dev/null)" || die "not in a repo"
  [ "$here" != "$main_root" ] || die "run this INSIDE a worktree, not the main clone"
  branch="$(git rev-parse --abbrev-ref HEAD)"
  case "$branch" in
    wt/*) ;;
    *) die "HEAD is '$branch', expected a wt/* branch - refusing" ;;
  esac
  if [ -n "$(git status --porcelain)" ]; then
    git status --short
    die "uncommitted changes above - commit (or discard) them first"
  fi

  # Land: rebase onto fresh dev, push, retry the race a bounded number of
  # times. A rebase CONFLICT stops here with state preserved - resolving it
  # is the session's job, nothing is auto-aborted.
  landed=0
  for attempt in 1 2 3; do
    git fetch origin --quiet || die "git fetch failed"
    if ! git rebase --quiet origin/dev; then
      echo "worktree.sh: rebase hit a conflict (attempt $attempt)." >&2
      echo "Resolve it, 'git rebase --continue', then rerun: bash .github/worktree.sh done" >&2
      exit 1
    fi
    if git push origin HEAD:dev; then
      landed=1
      break
    fi
    echo "worktree.sh: push raced another session's landing - retrying ($attempt/3)"
  done
  [ "$landed" -eq 1 ] || die "push did not land after 3 attempts - rerun when quieter"

  cd "$main_root" || die "cannot cd to main clone"
  git worktree remove --force "$here" || die "landed OK but could not remove $here - remove it by hand"
  git branch -D "$branch" >/dev/null
  echo "landed on dev and cleaned up ($branch gone). If your shell still sits in"
  echo "the removed folder, cd out:"
  echo
  echo "cd \"$main_root\""
  echo
  echo "The main clone's local dev is behind until its next git pull - that is normal."
  ;;

list)
  git -C "$main_root" worktree list
  echo
  git -C "$main_root" branch --list 'wt/*'
  ;;

*)
  die "usage: worktree.sh new <topic> | done | list"
  ;;
esac
