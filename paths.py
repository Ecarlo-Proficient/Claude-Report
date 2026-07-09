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


# ── Self-check (`python3 paths.py`) ────────────────────────────────────────
# Run this on a fresh machine to confirm your machine.env resolves correctly
# BEFORE running a real script. It reads/writes nothing but reports where each
# value came from and whether the target directory exists and is writable.

def _source(key: str) -> str:
    """Where would get()/get_path() pick this key up from?"""
    if key in os.environ and os.environ[key]:
        return "env var"
    if key in _MACHINE and _MACHINE[key]:
        return "machine.env"
    return "default"


def _dir_status(p: Path) -> str:
    """Human-readable existence + writability for a directory (or its parent)."""
    if p.exists():
        writable = os.access(p, os.W_OK)
        return "exists, writable" if writable else "exists, NOT writable"
    # Not there yet — a script would create it, so check the nearest parent.
    parent = p.parent
    while not parent.exists() and parent != parent.parent:
        parent = parent.parent
    if os.access(parent, os.W_OK):
        return f"missing (creatable — parent {parent} is writable)"
    return f"missing (parent {parent} NOT writable — WILL FAIL)"


def _print_root(label: str, key: str, value: Path) -> None:
    print(f"  {label}")
    print(f"    key    : {key}")
    print(f"    value  : {value}")
    print(f"    source : {_source(key)}")
    print(f"    status : {_dir_status(value)}")


def _self_check() -> None:
    print("paths.py — resolved configuration for this machine\n")
    if _MACHINE_ENV.exists():
        print(f"machine.env : FOUND — {_MACHINE_ENV}")
        print(f"              keys: {', '.join(sorted(_MACHINE)) or '(none parsed)'}")
    else:
        print(f"machine.env : not present — using owner defaults")
        print(f"              (copy machine.env.example -> machine.env to override)")

    print("\nShared roots (most outputs derive from these):")
    _print_root("OneDrive mirror", "ACB_ONEDRIVE_BASE", onedrive_base())
    _print_root("Company-health folder", "ACB_COMPANYHEALTH_DIR", companyhealth_dir())

    print("\nPer-file overrides (optional — unset means the script's own default):")
    for key in ("INVOICE_EXPORT_PATH", "WIP_EXCEL_PATH"):
        val = os.environ.get(key) or _MACHINE.get(key)
        if val:
            print(f"  {key} = {val}  (source: {_source(key)})")
        else:
            print(f"  {key} = <unset — derives from ACB_ONEDRIVE_BASE default>")

    print("\nAny line reading 'WILL FAIL' above means a script writing there will "
          "error.\nFix by pointing that root at a writable folder in machine.env.")


if __name__ == "__main__":
    _self_check()
