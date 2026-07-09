"""
paths.py — per-machine path resolution for the automation suite.

WHY THIS EXISTS
Output locations (OneDrive mirror, CompanyHealth folder, export files) differ
per machine. Code is identical on every clone, so paths must come from
configuration, not source. This module is the single lookup point.

HOW IT RESOLVES (first hit wins)
  1. process environment variable          (highest — ad-hoc overrides)
  2. machine.env at the repo root          (per-machine file, GITIGNORED)
  3. the default passed by the caller      (the owner's original paths)

On a machine with no machine.env and no env vars, every script behaves
EXACTLY as before this module existed. See machine.env.example for keys.

USAGE
    import paths
    OUT_DIR = paths.get_path("ACB_PNL_OUT_DIR",
                             paths.onedrive_base() / "Automations-/PROJECT P&Ls")
"""
from __future__ import annotations

import os
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
_MACHINE_ENV = _ROOT / "machine.env"


def _load_machine_env() -> dict:
    """Parse machine.env (KEY=VALUE lines, # comments). Missing file → {}."""
    data: dict = {}
    if _MACHINE_ENV.exists():
        for raw in _MACHINE_ENV.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            data[key.strip()] = val.strip().strip('"').strip("'")
    return data


_MACHINE = _load_machine_env()


def get(key: str, default: str = "") -> str:
    """String lookup: env var > machine.env > default."""
    return os.environ.get(key) or _MACHINE.get(key) or default


def get_path(key: str, default: "Path | str") -> Path:
    """Path lookup: env var > machine.env > default. Expands ~ in overrides."""
    val = os.environ.get(key) or _MACHINE.get(key)
    return Path(val).expanduser() if val else Path(default)


# ── Shared roots (most outputs derive from these two) ──────────────────────

_DEFAULT_ONEDRIVE_BASE = (
    Path.home() / "Library/CloudStorage/OneDrive-ProficientConcrete,LLC"
)
_DEFAULT_COMPANYHEALTH = Path.home() / "Documents" / "CompanyHealth"


def onedrive_base() -> Path:
    """Local mirror of the company OneDrive. Override: ACB_ONEDRIVE_BASE."""
    return get_path("ACB_ONEDRIVE_BASE", _DEFAULT_ONEDRIVE_BASE)


def companyhealth_dir() -> Path:
    """Local (non-synced) company-health folder. Override: ACB_COMPANYHEALTH_DIR."""
    return get_path("ACB_COMPANYHEALTH_DIR", _DEFAULT_COMPANYHEALTH)
