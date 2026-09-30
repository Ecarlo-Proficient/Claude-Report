# docker/ - the office server (Synology)

The package that runs the **QBO mirror** and the **AP/AR sync** on the office Synology, so they no longer depend on the
owner's Mac. One container, one process (`scheduler.py`):

| Job | When | What it writes |
|---|---|---|
| Mirror refresh | every 3 min | the server's own QBO mirror (its own QuickBooks login) |
| Count check | nightly 02:30 | the mirror (sweeps anything the change feed missed) |
| AP then AR (one job) | every 15 min | Bill Tracker.xlsx, Invoice Tracker.xlsx (Accounting share), Notion, Teams |

Everything else - the ledger, WIP, P&Ls, every QuickBooks WRITER - stays on the owner's Mac. The image carries read
code only; `tests/test_office_server_image.py` fails the build if a writer is ever copied in. (This folder replaced
the retired invoice-sync-only container on 09/30/2026.)

The plan this follows: the owner's one-page "Office server move" (09/30/2026) and the tools map `docs/architecture.html`.

---

## Runbook - for the developer (and the developer's Claude session)

Do the steps **in order**. Each **STOP** means: report to the owner and wait for a yes before going on. Never paste a
key, token or password into chat, email, a commit or this repo - keys are typed in on the Synology only.

### 0. Preconditions - STOP if any is not true

- [ ] The Synology **check-up** (`CHECKUP.md`, look only) is done and the owner has seen the results.
- [ ] The **lock-down** built from it is done: from outside, the only way in is the Fortinet VPN (no QuickConnect, no
      port forwards); two-step login on admin accounts; the default `admin` account off; the clock synced (NTP);
      locked snapshots + one backup off the Synology.
- [ ] Container Manager is installed; a plain, non-admin **automation user** exists.

### 1. Folders on the Synology

| Folder | What | Owner / rights |
|---|---|---|
| `/volume1/docker/automation/data` | the mirror, status, logs, state | automation user, 700 |
| `/volume1/automation-keys` | an **encrypted shared folder** (Control Panel > Shared Folder > Encryption) | automation user only |
| `/volume1/docker/automation/po-tracker` | the Cloud Sync copy (step 2) | automation user, read |
| `/volume1/Accounting/_automation/registers` | the owner's rule files, pushed from the Mac | automation user read/write |
| `/volume1/docker/automation/src` | this repo, cloned (step 6) | automation user |

Put the automation user's uid/gid (`id <user>` over SSH, or Control Panel) into `compose.yml` > `build.args`.

### 2. The PO tracker - Synology Cloud Sync, one way

The owner keeps using the PO tracker on OneDrive (`Proficient Office - Purchase Orders`). Cloud Sync > add the
Microsoft 365 / OneDrive for Business connection > that ONE folder > local path `/volume1/docker/automation/po-tracker`
> **Sync direction: download remote changes only**. The server reads the copy; it never writes back.

### 3. QuickBooks - the server's OWN app. STOP before connecting.

1. In the Intuit developer workspace (the same team workspace as the developer's own app), create a **new app**, e.g.
   "Office Server", scope Accounting, production keys.
2. **Never** use the owner's app or the developer's "EC-Data Export" app - a second connection on an app+company cuts
   off the first (the owner's Mac or the developer's clone stops working).
3. Connecting the app to the company needs a **QuickBooks admin** of the company to approve. If the developer is not
   one, send the owner the Connect link (the Intuit OAuth Playground with the new app's keys works) - the owner clicks
   Approve. Take the **refresh token** from that flow.
4. Type the four QBO values into `/volume1/automation-keys/secrets.env` (from `secrets.env.example`), mode 600.

### 4. The other keys

In the same `secrets.env`:
- `MIRROR_KEY` - the server's own mirror key. Generate ON the Synology and paste straight into the file:
  `python3 -c "import os,base64;print(base64.b64encode(os.urandom(32)).decode())"`
- `NOTION_SECRET`, `TEAMS_WEBHOOK_MFD_PAID`, `TEAMS_WEBHOOK_ALERTS` - the owner provides them on the Synology.

### 5. Paths and ids

Copy `server.env.example` to `/volume1/docker/automation/server.env`. Fill the Notion ids from the owner's Mac
(`invoice-sync/.env` - ids, not secrets). Keep `ACB_SERVER_MODE=test`. Check the General List file name for the
current year.

The owner runs, on the Mac, once (and after any rule change): `bash docker/push_registers.sh`

### 6. Code, build, start - in TEST mode

Clone the private repo into `/volume1/docker/automation/src` with a **read-only deploy key** (GitHub > the repo >
Settings > Deploy keys; never a personal token). Test week: the `dev` branch is fine. **Live runs `main` only.**

    cd /volume1/docker/automation/src
    git pull
    docker compose -f docker/compose.yml up -d --build
    docker compose -f docker/compose.yml logs -f office-server

The first mirror refresh seeds the whole mirror (~20 min, ~300k records) - expected.

### 7. The test week - STOP each day with the results

In test mode the server writes ONLY `Accounting/_server-test/` (its own Bill Tracker + Invoice Tracker, seeded from
copies of the live ones), runs AR as a dry run (Notion and Teams untouched) and never touches the live files.

Each day, report to the owner:
- `/volume1/docker/automation/data/status/status.json` - every job ok? any alerts in Teams?
- the test Bill Tracker / Invoice Tracker vs the Mac's live ones, sheet by sheet (the same rows and totals);
- **open-file test**: open the test trackers in Excel on a Mac over the share and leave them open; the server's next
  runs must still land and a reopen shows the new data.

### 8. Going live - NOT YET

Live needs these first (tracked in `docker/STATUS.md`): the Notion "only update when changed" fix, the Mac honoring the
writer file (sync-all stands down), the Bill Tracker Inputs split, the nightly mirror cross-check. Then, with the
owner's go: `ACB_SERVER_MODE=live`, create `Accounting/_automation/writer.json` = `{"writer": "server"}`, restart.

**Rollback** at any time: set the writer file to `{"writer": "mac"}`. The server stands down at its next run and the
owner's `sync-all` runs AP/AR as before.

---

## Files

| File | What |
|---|---|
| `Dockerfile` | the image: pinned base (digest) + pinned packages (sha256), the read-only allow-list |
| `.dockerignore` | lets in ONLY the allow-list |
| `compose.yml` | the Synology mounts, the key folder, no ports |
| `scheduler.py` | the clock: mirror / nightly check / AP-then-AR, test vs live, writer file, status, alerts |
| `server.env.example` | paths + ids (no secrets) |
| `secrets.env.example` | the keys' names (values typed on the Synology only) |
| `push_registers.sh` | Mac side: sends the owner's rule files to the server share (never cost_code_history.json) |
| `CHECKUP.md` | step 0: the look-only Synology check-up |
