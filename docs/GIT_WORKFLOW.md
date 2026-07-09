# Git Workflow — step by step

Plain-language reference for the commands you actually need: pulling, pushing to `dev`, submitting a release to `main`, and what branch protection blocks.

## Branch roles

- `main` — the released, stable branch. Every change must go through a pull request with 1 approving review and a passing `test` CI check. No direct pushes, no force pushes, no branch deletion — enforced even for repo admins.
- `dev` — the working branch. Direct pushes are allowed (no PR/review required), but force pushes and deletion are still blocked, and CI still runs on every push.

## One-time setup (new machine)

```bash
git clone https://github.com/Ecarlo-Proficient/Claude-Report.git
```

```bash
cd Claude-Report
```

```bash
gh auth login
```

## Start of every session

Always pull before making changes — a merged PR from someone else may have moved things.

```bash
git checkout dev
```

```bash
git pull origin dev
```

## Making changes on dev

Direct pushes to `dev` are allowed.

```bash
git add -A
```

```bash
git commit -m "<what changed>"
```

```bash
git push origin dev
```

Pushing triggers CI automatically (`.github/workflows/ci.yml`). Check status at:
https://github.com/Ecarlo-Proficient/Claude-Report/actions

## Submitting dev to main (release)

Once `dev` is tested and CI is green, open a pull request into `main`:

```bash
gh pr create --base main --head dev --title "Release: sync dev to main"
```

Before it can merge, it needs:
- The `test` CI check passing
- 1 approving review (a fresh push after approval resets it — stale reviews are dismissed)

Merge it once approved and green:

```bash
gh pr merge --merge
```

`--squash` or `--rebase` work too, in place of `--merge`, if you'd rather not create a merge commit. Merging from the PR page in the browser is equally fine.

## After merging to main

`main` auto-tags on every merge (`.github/workflows/tag-release.yml`), so rollback is always one command:

```bash
git tag --list
```

```bash
git checkout <tag>
```

Bring `dev` back in sync with the merged `main` so the branches don't drift apart:

```bash
git checkout dev
```

```bash
git merge main -m "Sync dev with main"
```

```bash
git push origin dev
```

## Feature branches (optional, for bigger or riskier changes)

You don't have to commit straight to `dev`. Branch off it, then PR back into `dev` — CI must pass, no review required:

```bash
git checkout dev
```

```bash
git pull origin dev
```

```bash
git checkout -b feature/<short-name>
```

```bash
git push -u origin feature/<short-name>
```

```bash
gh pr create --base dev --head feature/<short-name> --title "<what this does>"
```

## Just pulling the latest (read-only / running scripts against QBO)

If you're not making changes, this is all you need:

```bash
git checkout main
```

```bash
git pull origin main
```

## What's blocked

- Force-pushing to `main` or `dev` (`git push --force` is rejected).
- Deleting `main` or `dev`.
- Pushing directly to `main` — even for admins; it must go through a PR.
- Merging a PR into `main` without 1 approval and a green `test` check.

## Cheat sheet

| I want to... | Command |
|---|---|
| Get latest dev | `git pull origin dev` |
| Get latest main | `git pull origin main` |
| Push my changes to dev | `git push origin dev` |
| Open dev → main PR | `gh pr create --base main --head dev --title "Release: sync dev to main"` |
| Merge an approved PR | `gh pr merge --merge` |
| See CI status on a PR | `gh pr checks` |
| List release tags | `git tag --list` |
| Roll back to a tag | `git checkout <tag>` |
