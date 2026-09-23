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
import subprocess
from pathlib import Path

# machine.env lives at the REPO ROOT (one level above shared/) — same place
# as machine.env.example. Machines with an existing machine.env keep working
# across the 2026-07 restructure without touching it.
_REPO_ROOT = Path(__file__).resolve().parent.parent
_MACHINE_ENV = _REPO_ROOT / "machine.env"


def _main_worktree_env(repo_root: Path) -> "Path | None":
    """machine.env is per-machine and gitignored, so a linked git WORKTREE
    (one session = one worktree, 2026-08-27) starts life without one. Resolve
    the MAIN clone through the shared git common dir and use its copy, so
    worktrees inherit this machine's paths with zero setup. Any failure
    (no git, not a repo, no file) falls back to the old behavior."""
    try:
        out = subprocess.run(
            ["git", "-C", str(repo_root), "rev-parse",
             "--path-format=absolute", "--git-common-dir"],
            capture_output=True, text=True, timeout=5,
        )
        if out.returncode != 0:
            return None
        cand = Path(out.stdout.strip()).parent / "machine.env"
        return cand if cand.is_file() else None
    except Exception:
        return None


if not _MACHINE_ENV.exists():
    _MACHINE_ENV = _main_worktree_env(_REPO_ROOT) or _MACHINE_ENV


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
_DEFAULT_VAULT = Path.home() / "Documents" / "Claude" / "AI Brain_Vault"
_DEFAULT_ACCOUNTING_BASE = Path("/Volumes/Accounting")


def onedrive_base() -> Path:
    """Local mirror of the company OneDrive. Override: ACB_ONEDRIVE_BASE."""
    return get_path("ACB_ONEDRIVE_BASE", _DEFAULT_ONEDRIVE_BASE)


def accounting_base() -> Path:
    """The Accounting file share (Synology, must be mounted). Override:
    ACB_ACCOUNTING_BASE. Home of Accounts Payable/ (Vendor Statements + the Bill
    Tracker) and Accounts Receivable/."""
    return get_path("ACB_ACCOUNTING_BASE", _DEFAULT_ACCOUNTING_BASE)


# The roots a run cannot do without. A sync that starts with one of these missing wastes
# the owner's time (Touch ID, a QBO pull, then a mkdir traceback) - so every sync alias
# checks first and STOPs in plain words (owner 2026-09-23: "change sync-all, sync-ap,
# sync-ar to STOP if something isn't mounted, that way i don't waste time").
MOUNTS = {
    "accounting": (accounting_base, "the Accounting share (Synology) - Bill Tracker, vendor statements, stubs"),
    "onedrive":   (onedrive_base,   "the OneDrive mirror - Open_Invoices.xlsx, the WIP master, P&Ls"),
}


def missing_mounts(names) -> list:
    """[(name, path, description)] of the named roots that are not on disk right now."""
    out = []
    for n in names:
        fn, desc = MOUNTS[n]
        path = fn()
        if not path.is_dir():
            out.append((n, path, desc))
    return out


def require_mounts(names, what: str = "this run") -> None:
    """STOP (exit 2) with one plain message naming every missing root, before any work."""
    gone = missing_mounts(names)
    if not gone:
        return
    lines = [f"STOP - {what} needs a drive that is not mounted right now. Nothing was run."]
    for n, path, desc in gone:
        lines.append(f"   missing: {path}   ({desc})")
    lines.append("   Reconnect it (Finder > Go > Connect to Server for the Synology share; open OneDrive for "
                 "the mirror), then run again.")
    raise SystemExit("\n".join(lines))


def require_accounting_share(what: str = "this run") -> Path:
    """STOP, in plain words, when the Accounting share is not mounted - never a traceback
    and never a fallback path (owner 2026-09-23, the bill sync died in mkdir on
    /Volumes/Accounting after the share dropped). Returns the base when it is there."""
    base = accounting_base()
    if base.is_dir():
        return base
    raise SystemExit(
        f"STOP - the Accounting share is not mounted, so {what} has nowhere to write.\n"
        f"   expected: {base}\n"
        f"   Reconnect it in Finder (Go > Connect to Server, the Synology Accounting share), "
        f"then run again. Nothing was changed.")


def bill_tracker_xlsx() -> Path:
    """The AP Bill Tracker workbook - the ONE resolver every tool that reads or
    writes it shares, so the path can never drift. Moved 2026-09-16 into the
    Accounting share's Accounts Payable/ folder (was OneDrive Automations-/).
    Override: ACB_BILL_TRACKER_XLSX (e.g. the Dockerized invoice-sync points it at
    its own in-container mount)."""
    return get_path("ACB_BILL_TRACKER_XLSX",
                    accounting_base() / "Accounts Payable" / "Bill Tracker.xlsx")


def bill_payment_stubs_dir() -> Path:
    """Where printed bill payment stubs live (owner 2026-09-22): the Accounting share's
    Accounts Payable/Bill Payment Stubs/<vendor>/. Override: ACB_BILL_PAYMENT_STUBS_DIR.
    The share must be mounted - a missing mount is a STOP, never a fallback."""
    return get_path("ACB_BILL_PAYMENT_STUBS_DIR",
                    accounting_base() / "Accounts Payable" / "Bill Payment Stubs")


def companyhealth_dir() -> Path:
    """Local (non-synced) company-health folder. Override: ACB_COMPANYHEALTH_DIR.
    Holds only the two things the owner opens — Company Tracker.xlsx and
    Company Dashboard.html; the workbooks that feed them live in _sources/."""
    return get_path("ACB_COMPANYHEALTH_DIR", _DEFAULT_COMPANYHEALTH)


def vault_dir() -> Path:
    """The AI Brain_Vault — the business brain and the systems/process registry.
    Override: ACB_VAULT_DIR. READ-ONLY from this repo: tools may render vault
    content (the ledger's Systems tab reads 02_processes/), never write it."""
    return get_path("ACB_VAULT_DIR", _DEFAULT_VAULT)


def process_registry_dir() -> Path:
    """The registry's domain files (02_processes/) inside the vault."""
    return vault_dir() / "02_processes"


def process_guides_dir() -> Path:
    """The one-page process guides (assets/processes/) inside the vault:
    `<PROCESS-ID>_<slug>.html` (source) + `.pdf` (the handout). READ-ONLY from
    this repo - the ledger's Systems tab links a registry row to its guide."""
    return vault_dir() / "assets" / "processes"


def companyhealth_sources_dir() -> Path:
    """Where the intermediate tracker workbooks live (the data layer the one
    workbook reads). Kept out of the top level so only the deliverables show
    (the user 2026-07-28). Created on demand."""
    d = companyhealth_dir() / "_sources"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return d


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
    _print_root("AI Brain_Vault (read-only)", "ACB_VAULT_DIR", vault_dir())

    print("\nPer-file overrides (optional — unset means the script's own default):")
    for key in ("INVOICE_EXPORT_PATH", "WIP_EXCEL_PATH", "ACB_GENERAL_LIST_XLSX"):
        val = os.environ.get(key) or _MACHINE.get(key)
        if val:
            print(f"  {key} = {val}  (source: {_source(key)})")
        else:
            print(f"  {key} = <unset — derives from ACB_ONEDRIVE_BASE default>")

    print("\nAny line reading 'WILL FAIL' above means a script writing there will "
          "error.\nFix by pointing that root at a writable folder in machine.env.")


if __name__ == "__main__":
    import sys as _sys
    if len(_sys.argv) > 1 and _sys.argv[1] == "--require":
        # python3 shared/paths.py --require accounting,onedrive --for "sync-all"   -> exit 2 + STOP when one is missing
        names = [n for n in _sys.argv[2].split(",") if n]
        what = _sys.argv[4] if len(_sys.argv) > 4 and _sys.argv[3] == "--for" else "this run"
        try:
            require_mounts(names, what)
        except SystemExit as e:
            print(e)
            _sys.exit(2)
        _sys.exit(0)
    _self_check()
