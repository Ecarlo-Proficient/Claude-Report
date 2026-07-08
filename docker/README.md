# Invoice Sync — Docker package · **v1.0.0**

This directory holds the Docker package for the AR invoice sync — the **v1.0.0 "true release."** It moves the sync off the user's Mac onto an always-on Synology container. The container runs `run_invoice_sync.py` on a 15-minute loop: pulls open invoices from QuickBooks Online, upserts them into two Notion databases, sweeps for paid invoices, archives invoices deleted in QBO (CDC), and posts MFD payment events to a Teams channel.

## Versioning

| Line | Version | What it is |
|---|---|---|
| **Docker** | `v1.0.0` | The true release — this package. Set via `APP_VERSION=1.0.0` baked into the image. Production target. |
| **Mac** | `mvN` | The Mac-only lineage (manual `sync-ar` + visual viewer). Currently `mv1`; bump on Mac-side changes. Stays live as dev/fallback while Docker is tested. |

Every run logs its identity (`Starting … [v1.0.0 (docker)]` or `[mv1 (mac)]`) and every Teams alert names the runtime — so while both run during testing, you always know which instance is talking. The label comes from `automation-worker/version.py`.

> **Coexistence note:** during testing, Excel export is **disabled in the container** (`SKIP_EXCEL_EXPORT=1`) so the Mac keeps owning the Excel mirror and the two don't fight over the OneDrive file. The container does the core QBO→Notion work; the Mac `sync-ar` continues as today.

---

## What this container does

Runs `automation-worker/run_invoice_sync.py` on a 15-minute loop. Each pass:

1. Authenticates to QuickBooks Online via OAuth refresh token
2. Pulls open invoices (`Balance > 0`) from QBO
3. Upserts them into two Notion databases (Res/Com and MFD), routed by Project # prefix
4. Sweeps for invoices QBO no longer reports as open → flips to Paid (or archives the Notion row if the invoice was deleted while still open)
5. **CDC deletion pass** — asks QBO what invoices were deleted since the last run and archives the matching Notion rows, including ones already marked Paid (catches void→delete). Watermark persists in `/data/state`.
6. Archives Notion rows whose Paid Date is older than 12 months
7. Posts MFD paid / short-pay events (with the billed line items) to a Teams channel; on a failed or errored run, posts a **warning** to the ops-alert Teams webhook
8. Excel mirror: **skipped in v1** (`SKIP_EXCEL_EXPORT=1`). The OneDrive bridge below is pre-staged for when Excel moves into the container
9. Sleeps until the next run

No inbound network. Outbound HTTPS to `api.notion.com` and `*.intuit.com`/`oauth.platform.intuit.com`, plus `*.logic.azure.com` for Teams webhooks.

---

## Files in this directory

| File | Purpose |
|---|---|
| `Dockerfile` | Image definition. Python 3.11-slim base. |
| `docker-compose.yml` | Service definition with volumes, env file, restart policy, resource limits. |
| `entrypoint.sh` | Run loop with `SYNC_INTERVAL_SECONDS` (default 900) between syncs. Handles SIGTERM for clean shutdown. |
| `.env.docker.example` | Template for the runtime config + secrets file. Copy to `.env.docker` and fill in. |
| `.dockerignore` | Excludes secrets, caches, and unrelated files from build context. |
| `README.md` | This file. |

---

## How to build and run

From this directory:

```bash
# 1. Copy and fill in the env file (chmod 600 on host)
cp .env.docker.example .env.docker
chmod 600 .env.docker
# edit .env.docker with real NOTION_SECRET, QBO_* values

# 2. Build and start
docker compose up -d --build

# 3. Watch logs
docker compose logs -f invoice-sync

# Expected output every 15 min:
#   Starting qbo_invoices_to_notion [v1.0.0 (docker)] dry_run=False
#   ... summary lines per division ...
#   qbo_invoices_to_notion clean.
```

To verify QBO → Notion is working without affecting production data:

```bash
# Run once in dry-run mode (no Notion writes)
docker compose run --rm -e SYNC_RUN_ONCE=1 invoice-sync python3 run_invoice_sync.py --dry-run
```

To stop:

```bash
docker compose down
```

---

## Secrets handling

The test setup uses a plain `.env.docker` file, chmod 600, owned by the user running Docker on the Synology. The file contains:

- `NOTION_SECRET` — Notion internal integration token
- `QBO_CLIENT_ID`, `QBO_CLIENT_SECRET`, `QBO_COMPANY_ID`, `QBO_REFRESH_TOKEN` — QuickBooks Online OAuth
- `TEAMS_WEBHOOK_MFD_PAID` — Teams Workflows webhook for MFD payment cards (optional; posting credential)
- `TEAMS_WEBHOOK_ALERTS` — Teams Workflows webhook for sync failure/error warnings (optional; posting credential). This is your unattended-container alarm — set it.

On Mac these webhooks live in the Keychain; in the container they're env vars in `.env.docker`. Treat the populated `.env.docker` as a secrets file (chmod 600, never commit).

The QBO refresh token rotates periodically. The container persists rotations to a named Docker volume (`proficient-invoice-sync-state` mounted at `/data`) so restarts pick up the latest token. On first start the env-file values seed the file; subsequent rotations write back to the volume.

**Open question for IT:** Synology-native preference — Docker secrets, Container Manager's built-in env handling, or an external vault (HashiCorp / Bitwarden CLI)? Env-file works for the test; IT picks the production pattern.

---

## Volumes

| Mount type | Mount point | Purpose |
|---|---|---|
| Named volume `proficient-invoice-sync-state` | `/data` | Rotated QBO refresh tokens (`qbo_secrets.json`) **and** `/data/state` (sync + CDC-deletion watermarks). Must survive container restarts — without persisted state, CDC re-scans the last 30 days every recreate. |
| Bind mount | `/data/onedrive` | The Excel mirror (`Open_Invoices.xlsx`) lands here every sync. IT bind-mounts this to a Synology folder that **Cloud Sync** mirrors to OneDrive. See "OneDrive bridge" below. |
| Named volume `proficient-invoice-sync-logs` | `/app/automation-worker/logs` | Human-readable file logs alongside `docker logs`. Optional — `docker logs` alone is sufficient. |

Named volumes are used for state and logs so IT doesn't need to manage UID/GID on the host filesystem. The OneDrive bridge is a bind mount because it has to be in the Cloud-Sync-watched folder. The container runs as UID 1000 by default; override via the `APP_UID` / `APP_GID` build args in `docker-compose.yml` if needed.

---

## OneDrive bridge

The AR clerk's workflow includes a read-only Excel mirror of open invoices that lives in OneDrive (`Collections/Open_Invoices.xlsx`). The sync writes the file every run; OneDrive's coauthoring handles distribution to the clerk's machine. The container writes Excel to `/data/onedrive/Open_Invoices.xlsx` (a Synology folder bind-mounted into the container). **Synology Cloud Sync** is the bridge — it mirrors that folder to OneDrive within seconds.

**Setup on the Synology** (IT):

1. Install / open the **Cloud Sync** package on DSM.
2. Add a new connection → **Microsoft OneDrive for Business**.
3. Authenticate with the OneDrive account that owns the `Collections/` folder.
4. Choose sync direction: **Bidirectional** (so when the clerk opens the file, the Office lock file syncs back to the Synology; the script's lock-guard check uses this to skip writes during clerk edits).
5. Set the local path to `/volume1/docker/invoice-sync/onedrive` (or whatever path IT prefers — must match the bind mount in `docker-compose.yml`).
6. Set the remote path to OneDrive's `Collections/` folder.

**Why Cloud Sync vs. rclone**: Cloud Sync is Synology-native, has the OAuth wired in via the GUI, survives reboots, and is what IT already knows. Rclone would work too but adds a moving part with no benefit.

**Lock guard interaction**: the export script checks for Excel's hidden `~$Open_Invoices.xlsx` file in the same folder before writing. With Cloud Sync running bidirectional, the clerk's open-file lock syncs back to the Synology folder, so the container correctly skips the write while the clerk is in the file. Same mechanism as the Mac version of the sync.

---

## Outbound network requirements

Container needs HTTPS (443/tcp) egress to:

- `api.notion.com` — Notion REST API
- `quickbooks.api.intuit.com` — QBO REST API
- `oauth.platform.intuit.com` — OAuth token refresh
- `*.logic.azure.com` — Teams Workflows webhooks (MFD payment cards + ops alerts), only if those webhooks are configured

No inbound traffic. No DNS or NTP requirements beyond standard.

---

## Resource sizing

Expected footprint:

- **RAM:** peaks around 200–300 MB during a sync, drops to ~80 MB while sleeping. Compose caps at 512 MB.
- **CPU:** ~5% of one core during a sync (mostly I/O wait on API responses), idle between runs.
- **Disk:** image ~150 MB. Volume usage is a few KB of JSON (rotated tokens) + log file growth (capped at 50 MB by Docker log rotation).
- **Network:** ~1–2 MB per sync run (Notion + QBO API traffic).

Light footprint overall — the script is I/O bound (waiting on QBO and Notion APIs), not compute-bound.

---

## Open questions for the IT call

| # | Question |
|---|---|
| 1 | Are you running Container Manager (Synology's GUI wrapper), the plain `docker` / `docker compose` CLI, Portainer, or something else for managing containers? |
| 2 | Preferred secrets approach for production: env file + chmod 600, Docker secrets, Synology's built-in secrets handling, or external vault? |
| 3 | Should the container run as a non-root UID/GID you specify (so bind-mounted host paths inherit the right owner), or are named volumes fine? |
| 4 | Outbound firewall — is unrestricted HTTPS egress fine, or do you need a specific allowlist? (See list above.) |
| 5 | Log shipping — is `docker logs` (json-file driver, 10MB × 5 rotation) acceptable, or do you have a centralized log collector (Loki, ELK, Synology Log Center) we should ship into? |
| 6 | Restart / monitoring — does Synology surface container health to your monitoring stack, or should the container ping an external healthcheck URL? |
| 7 | Updates — do you want a CI pipeline (GitHub Actions builds an image, Synology pulls), or a manual `git pull && docker compose up -d --build` from a checkout on the NAS? |

---

## Out of scope (not in this package)

The container does only the AR invoice sync (`run_invoice_sync.py`). No other automation scripts, no QBO webhook receiver. If real-time event-driven sync is needed later, a webhook receiver can be added as a second service in the same `docker-compose.yml`, exposed publicly via Cloudflare Tunnel — but that's a separate conversation.

---

## Rollback

If anything goes sideways during the test, the Mac side is untouched. Stop the container, sync resumes its current Mac launchd behavior. Nothing in Notion gets corrupted — the sync is idempotent (upsert by Invoice ID).

```bash
docker compose down
# Mac sync continues as normal — no separate action needed.
```
