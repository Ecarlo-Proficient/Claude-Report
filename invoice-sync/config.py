"""
Configuration loader for the automation worker.

Loads non-secret config from .env, and secrets from the OS keystore.

Platform handling:
  - macOS → secrets come from Keychain via the `keyring` library.
              First access triggers a system dialog. Click "Always Allow"
              once and the binary can read the key silently thereafter.
  - Linux (Pi) → secrets come from a separate .env.secrets file
              (chmod 600, Pi-only, never copied from Mac).
              A keyring-compatible libsecret backend also works if present.

Never put secrets in .env. .env is considered non-sensitive config only.
"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

# Repo root on sys.path so `shared/` (the key library) is importable.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_vault as kc  # noqa: E402


# Project root = folder this file lives in (invoice-sync/)
PROJECT_ROOT = Path(__file__).resolve().parent

# Pre-2026-07 layout: this folder was automation-worker/. A git pull renames
# the tracked files but leaves untracked .env / state/ behind in the old
# folder — fall back there so the sync keeps working until they're moved
# (then delete the old folder).
_LEGACY_DIR = PROJECT_ROOT.parent / "automation-worker"

# Load .env (non-secret config) — new location first, legacy fallback.
_ENV_FILE = PROJECT_ROOT / ".env"
if not _ENV_FILE.exists() and (_LEGACY_DIR / ".env").exists():
    _ENV_FILE = _LEGACY_DIR / ".env"
load_dotenv(_ENV_FILE)

# Pi-only secrets file. On Mac this file should NOT exist — Mac uses Keychain.
# Kept separate from .env so backup / rsync rules can exclude it explicitly.
_SECRETS_FILE = PROJECT_ROOT / ".env.secrets"
if not _SECRETS_FILE.exists() and (_LEGACY_DIR / ".env.secrets").exists():
    _SECRETS_FILE = _LEGACY_DIR / ".env.secrets"
if _SECRETS_FILE.exists():
    load_dotenv(_SECRETS_FILE, override=True)


@dataclass(frozen=True)
class Config:
    # Notion
    notion_secret: str

    # Invoice tracker (QBO → Notion)
    invoice_res_com_ds_id: str
    invoice_mfd_ds_id: str
    customer_list_ds_id: str       # for Res/Com customer relation lookup
    mfd_client_list_ds_id: str     # for MFD customer relation lookup
    invoice_paid_retention_months: int

    # WIP tracker — Notion path RETIRED 2026-06-25.
    # The Notion-based WIP DBs were deprecated when WIP source-of-truth pivoted
    # to Excel on SharePoint. The 4 WIP_*_DS_ID fields were removed from this
    # config. The new wip_sync.py (Excel-targeted, Test-sheet-only via
    # wip_excel_guard.py) will land its own config fields when written.
    wip_lookback_months: int        # kept — used by QBO query window (still relevant for Excel sync)
    wip_min_activity_usd: float     # kept — same reason

    # Behavior
    overlap_seconds: int
    initial_lookback: str  # ISO 8601 string

    # Paths
    state_dir: Path
    log_dir: Path

    # Teams MFD paid/short-pay Workflows webhook (POSTING CREDENTIAL — Keychain,
    # not .env). Optional: empty string disables Teams notifications.
    teams_webhook_mfd_paid: str = ""

    # Teams OPERATIONS-ALERT webhook — separate channel for sync failure/error
    # warnings (esp. for the unattended Docker container). Optional.
    teams_webhook_alerts: str = ""

    # Notion API
    notion_api_base: str = "https://api.notion.com/v1"
    # 2025-09-03+ required for /data_sources/{id}/query endpoint (multi-source DBs).
    # Earlier versions return 400 "invalid_request_url" for this path.
    notion_version: str = "2025-09-03"


def _require_env(name: str) -> str:
    val = os.getenv(name)
    if not val:
        raise RuntimeError(
            f"Missing required env var {name}. "
            f"Copy .env.example to .env and fill in real values."
        )
    return val


def _get_notion_secret() -> str:
    """
    The Notion integration token - env NOTION_SECRET (Linux / Docker), else the key library via
    Key Helper (shared/qbo_vault.get_secret; moved there from the older Keychain store 09/29/2026).
    Never logs or prints the secret itself.
    """
    try:
        val = kc.get_secret("NOTION_SECRET")
    except kc.SecretsError as e:
        raise RuntimeError(f"Notion token lookup failed: {e}")
    if val:
        return val
    raise RuntimeError(
        "Notion secret not found. On Mac: python3 shared/setup_qbo.py --rotate NOTION_SECRET. "
        "On Linux: put NOTION_SECRET=... in .env.secrets (chmod 600)."
    )


def _get_optional_webhook(env_var: str) -> str:
    """
    An OPTIONAL Teams Workflows webhook URL (a posting credential): env `env_var`, else the key library
    via Key Helper (shared/qbo_vault.get_secret, same name). "" = not configured, the feature is then
    silently disabled. Never raises, never logs the value.
    """
    try:
        return kc.get_secret(env_var)
    except kc.SecretsError:
        return ""   # optional secret - never block the sync on a lookup error


def _get_teams_webhook() -> str:
    """MFD paid/short-pay notification webhook (env TEAMS_WEBHOOK_MFD_PAID)."""
    return _get_optional_webhook("TEAMS_WEBHOOK_MFD_PAID")


def _get_teams_alert_webhook() -> str:
    """Operations-alert webhook for sync failures/errors (env TEAMS_WEBHOOK_ALERTS)."""
    return _get_optional_webhook("TEAMS_WEBHOOK_ALERTS")


def load_config() -> Config:
    # State dir holds the sync watermark + CDC deletion watermark. Override with
    # STATE_DIR (Docker points this at a persistent volume, e.g. /data/state, so
    # the CDC changedSince watermark survives a container recreate).
    # Legacy fallback: if state/ hasn't been moved out of automation-worker/
    # yet, keep using it there — losing the CDC watermark would re-scan
    # deletions from scratch.
    _default_state = PROJECT_ROOT / "state"
    if not _default_state.exists() and (_LEGACY_DIR / "state").exists():
        _default_state = _LEGACY_DIR / "state"
    state_dir = Path(os.getenv("STATE_DIR", str(_default_state)))
    # Logs live OUTSIDE the project folder (privacy: project folder is
    # AI-session-visible). Override with LOG_DIR env var (e.g. Docker).
    log_dir = Path(
        os.getenv("LOG_DIR", str(Path.home() / "Library/Logs/Proficient/automation-worker"))
    )
    state_dir.mkdir(exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)

    return Config(
        notion_secret=_get_notion_secret(),
        invoice_res_com_ds_id=_require_env("INVOICE_RES_COM_DS_ID"),
        invoice_mfd_ds_id=_require_env("INVOICE_MFD_DS_ID"),
        customer_list_ds_id=_require_env("CUSTOMER_LIST_DS_ID"),
        mfd_client_list_ds_id=_require_env("MFD_CLIENT_LIST_DS_ID"),
        invoice_paid_retention_months=int(os.getenv("INVOICE_PAID_RETENTION_MONTHS", "12")),
        # WIP_*_DS_ID removed 2026-06-25 — Notion WIP path retired.
        wip_lookback_months=int(os.getenv("WIP_LOOKBACK_MONTHS", "24")),
        wip_min_activity_usd=float(os.getenv("WIP_MIN_ACTIVITY_USD", "5000")),
        overlap_seconds=int(os.getenv("OVERLAP_SECONDS", "30")),
        initial_lookback=os.getenv("INITIAL_LOOKBACK", "2026-01-01T00:00:00Z"),
        state_dir=state_dir,
        log_dir=log_dir,
        teams_webhook_mfd_paid=_get_teams_webhook(),
        teams_webhook_alerts=_get_teams_alert_webhook(),
    )
