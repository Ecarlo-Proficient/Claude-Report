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

## Setup - for the developer (and the developer's Claude session)

The steps in order. Two rules that are not optional: the server gets its **own** Intuit app (step 3), and keys are
typed in on the Synology only - never in chat, email, a commit or this repo.

### 0. The box

Before anything goes on it: from outside, the only way in is the Fortinet VPN (no QuickConnect, no port forwards);
two-step login on admin accounts; the default `admin` account off; the clock synced (NTP); locked snapshots + one
backup off the Synology. `CHECKUP.md` is a handy list of what to look at. Install Container Manager and create a plain,
non-admin **automation user**.

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

### 3. QuickBooks auth - the server's OWN app

1. In the Intuit developer workspace (the same team workspace as the developer's own app), create a **new app**, e.g.
   "Office Server", scope Accounting, production keys.
2. **Never** use the owner's app or the developer's "EC-Data Export" app - a second connection on an app+company cuts
   off the first (the owner's Mac or the developer's clone stops working).
3. Connect it to the company (the Intuit OAuth Playground with the new app's production keys, scope Accounting) and
   take the **refresh token**. The approval must come from a **QuickBooks admin** of the company - if that is not the
   developer, send the owner the connect link to approve.
4. Type the four QBO values into `/volume1/automation-keys/secrets.env` (from `secrets.env.example`), mode 600.

### 4. The other keys

In the same `secrets.env`:
- `MIRROR_KEY` - the server's own mirror key. Generate ON the Synology and paste straight into the file:
  `python3 -c "import os,base64;print(base64.b64encode(os.urandom(32)).decode())"`
- `NOTION_SECRET`, `TEAMS_WEBHOOK_MFD_PAID`, `TEAMS_WEBHOOK_ALERTS` - the owner provides them on the Synology.
- `TEAMS_WEBHOOK_PAYMENTS_TEST` - the Payments channel's Workflows webhook, for the QuickBooks payment card test (step 7).
  Leave `TEAMS_WEBHOOK_PAYMENTS` blank until go-live.

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

### 7. The test week

In test mode the server writes ONLY `Accounting/_server-test/` (its own Bill Tracker + Invoice Tracker, seeded from
copies of the live ones), runs AR as a dry run (Notion and Teams untouched) and never touches the live files.

What to watch:
- `/volume1/docker/automation/data/status/status.json` - every job ok; alerts land in Teams;
- the test Bill Tracker / Invoice Tracker vs the Mac's live ones, sheet by sheet (the same rows and totals);
- open the test trackers in Excel on a Mac over the share and leave them open - the server's next runs must still land
  and a reopen shows the new data.
- **QuickBooks payment cards - only after the AP/AR checks above are clean; set `PAYMENTS_TEST=1` in server.env** (`payments-test` in status.json, log `payments-test.log`): every payment a client makes
  through the QuickBooks invoice link posts ONE "TEST - QuickBooks payment received" card to the Payments channel within
  ~15 minutes (the first run only records what is already there). Check each card against QuickBooks (client, project,
  amount, invoice #, paid in full / still open) and that a typed-in check never posts. No test webhook yet = the log
  lists what would post. Go-live: move the URL to `TEAMS_WEBHOOK_PAYMENTS`, blank the TEST line, restart.

### 8. Going live - NOT YET

Live needs these first (tracked in `docker/STATUS.md`): the Notion "only update when changed" fix, the Mac honoring the
writer file (sync-all stands down), the Bill Tracker Inputs split, the nightly mirror cross-check. Then:
`ACB_SERVER_MODE=live`, create `Accounting/_automation/writer.json` = `{"writer": "server"}`, restart.

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
| `CHECKUP.md` | a list of what to look at on the box (step 0) |
