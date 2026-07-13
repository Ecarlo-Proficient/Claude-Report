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


# Project root = folder this file lives in
PROJECT_ROOT = Path(__file__).resolve().parent

# Load .env (non-secret config) from project root
load_dotenv(PROJECT_ROOT / ".env")

# Pi-only secrets file. On Mac this file should NOT exist — Mac uses Keychain.
# Kept separate from .env so backup / rsync rules can exclude it explicitly.
_SECRETS_FILE = PROJECT_ROOT / ".env.secrets"
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
    Fetch the Notion integration secret.

    Resolution order:
      1. macOS → Keychain via `keyring` (service = KEYSTORE_SERVICE,
         username = KEYSTORE_KEY_NOTION).
      2. Linux / anything else → environment variable NOTION_SECRET
         (expected to come from .env.secrets on the Pi, loaded above).
      3. Fallback: if keyring backend is available on Linux (libsecret),
         try that before giving up.

    Never logs or prints the secret itself. Only raises descriptive errors.
    """
    service = os.getenv("KEYSTORE_SERVICE", "proficient-automation-worker")
    key_name = os.getenv("KEYSTORE_KEY_NOTION", "notion")

    # Environment variable takes precedence on non-Mac platforms (Pi).
    if sys.platform != "darwin":
        env_val = os.getenv("NOTION_SECRET")
        if env_val:
            return env_val

    # Keychain / libsecret path.
    try:
        import keyring  # imported lazily so non-Mac installs without it still start
    except ImportError:
        keyring = None  # type: ignore

    if keyring is not None:
        try:
            stored = keyring.get_password(service, key_name)
        except Exception as e:
            raise RuntimeError(
                f"Keystore lookup failed for {service}/{key_name}: {e}. "
                f"On Mac, run `python setup_keychain.py` to store the secret."
            )
        if stored:
            return stored

    # Last-chance environment fallback (handy for CI / debugging).
    env_val = os.getenv("NOTION_SECRET")
    if env_val:
        return env_val

    raise RuntimeError(
        "Notion secret not found. "
        "On Mac: run `python setup_keychain.py` once. "
        "On Pi: put NOTION_SECRET=... in .env.secrets (chmod 600)."
    )


def _get_optional_webhook(env_var: str, keystore_key_default: str,
                          keystore_key_env: str) -> str:
    """
    Resolve an OPTIONAL Teams Workflows webhook URL (a posting credential).

    Resolution order:
      1. Non-Mac (Linux / Docker) → environment variable `env_var` (primary;
         no Keychain there).
      2. macOS → Keychain (service = KEYSTORE_SERVICE, username from
         `keystore_key_env` env var, default `keystore_key_default`).
      3. Fallback → `env_var` (e.g. value still in .env pre-migration).

    Returns "" if nothing is configured — the feature is then silently disabled.
    Never raises, never logs the value.
    """
    service = os.getenv("KEYSTORE_SERVICE", "proficient-automation-worker")
    key_name = os.getenv(keystore_key_env, keystore_key_default)

    if sys.platform != "darwin":
        env_val = os.getenv(env_var, "").strip()
        if env_val:
            return env_val

    try:
        import keyring  # lazy import so non-Mac installs without it still start
    except ImportError:
        keyring = None  # type: ignore

    if keyring is not None:
        try:
            stored = keyring.get_password(service, key_name)
        except Exception:
            stored = None  # optional secret — never block the sync on a lookup error
        if stored:
            return stored.strip()

    return os.getenv(env_var, "").strip()


def _get_teams_webhook() -> str:
    """MFD paid/short-pay notification webhook (env TEAMS_WEBHOOK_MFD_PAID)."""
    return _get_optional_webhook(
        "TEAMS_WEBHOOK_MFD_PAID", "teams_webhook_mfd_paid", "KEYSTORE_KEY_TEAMS_WEBHOOK"
    )


def _get_teams_alert_webhook() -> str:
    """Operations-alert webhook for sync failures/errors (env TEAMS_WEBHOOK_ALERTS)."""
    return _get_optional_webhook(
        "TEAMS_WEBHOOK_ALERTS", "teams_webhook_alerts", "KEYSTORE_KEY_TEAMS_ALERTS"
    )


def load_config() -> Config:
    # State dir holds the sync watermark + CDC deletion watermark. Override with
    # STATE_DIR (Docker points this at a persistent volume, e.g. /data/state, so
    # the CDC changedSince watermark survives a container recreate).
    state_dir = Path(os.getenv("STATE_DIR", str(PROJECT_ROOT / "state")))
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
