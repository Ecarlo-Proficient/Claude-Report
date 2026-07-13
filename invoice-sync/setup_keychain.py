"""
One-time Keychain setup for the automation worker (macOS).

Stores worker secrets in the macOS Keychain under:

    service  = $KEYSTORE_SERVICE          (default: proficient-automation-worker)
    key name = per-secret (see below)

Secrets it can store:

  (default)   Notion integration secret
                key name = $KEYSTORE_KEY_NOTION          (default: notion)

  --teams     Teams MFD paid/short-pay Workflows webhook URL
                key name = $KEYSTORE_KEY_TEAMS_WEBHOOK   (default: teams_webhook_mfd_paid)

Usage:

    python setup_keychain.py                  # paste Notion secret (hidden)
    python setup_keychain.py --teams          # paste Teams webhook URL (hidden)
    python setup_keychain.py --teams --from-env
                                              # migrate the existing
                                              # TEAMS_WEBHOOK_MFD_PAID value out
                                              # of .env into Keychain, then blank
                                              # the .env line — no re-paste needed

Why this file exists:
  - Interactive pastes go through getpass() so the secret never appears
    on-screen, in shell history, or in any log.
  - The --from-env migration reads the value from .env *inside this script*
    (it is never printed) and removes it from .env once Keychain confirms it.
  - After a secret is stored, the first worker run triggers a Keychain prompt
    ("python3 wants to access key '…'"). Click "Always Allow" once; silent
    access thereafter.

Re-run any time to rotate — it overwrites.

Mac-only. On the Pi/Docker, put secrets in the environment (.env.secrets /
.env.docker) instead — there is no Keychain there.
"""
from __future__ import annotations

import argparse
import os
import sys
from getpass import getpass
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent
ENV_PATH = PROJECT_ROOT / ".env"
load_dotenv(ENV_PATH)


def _store(keyring, service: str, key_name: str, secret: str) -> int:
    """Store one secret in Keychain and verify by length-only read-back."""
    try:
        keyring.set_password(service, key_name, secret)
    except Exception as e:
        print(f"Failed to store in Keychain: {e}", file=sys.stderr)
        return 2

    try:
        check = keyring.get_password(service, key_name) or ""
    except Exception:
        check = ""

    if check == secret:
        print(f"Stored OK under {service}/{key_name}. "
              f"({len(secret)} chars, not displayed.)")
        return 0
    print("Stored, but read-back failed. Check Keychain Access.app manually.",
          file=sys.stderr)
    return 2


def _blank_env_line(var_name: str) -> bool:
    """
    Replace `VAR=<value>` in .env with an empty assignment plus a pointer
    comment. Returns True if a non-empty value was cleared. The value is never
    printed. Leaves all other lines untouched.
    """
    if not ENV_PATH.exists():
        return False
    lines = ENV_PATH.read_text().splitlines(keepends=False)
    changed = False
    out = []
    for line in lines:
        stripped = line.lstrip()
        if stripped.startswith(f"{var_name}=") and not stripped.startswith("#"):
            had_value = len(line.split("=", 1)[1].strip()) > 0
            out.append(f"# {var_name} moved to Keychain (setup_keychain.py --teams). "
                       f"Env var is now a Linux/Docker-only fallback.")
            out.append(f"{var_name}=")
            changed = changed or had_value
        else:
            out.append(line)
    if changed:
        ENV_PATH.write_text("\n".join(out) + "\n")
    return changed


def _setup_notion(keyring, service: str) -> int:
    key_name = os.getenv("KEYSTORE_KEY_NOTION", "notion")
    print("Storing the NOTION integration secret.")
    print(f"  service  = {service}")
    print(f"  key name = {key_name}")
    print()
    print("Paste your Notion integration secret and press Enter.")
    print("The paste is HIDDEN — you will not see it appear.")
    print("(Cancel with Ctrl-C if you change your mind.)")
    print()
    try:
        secret = getpass("Notion integration secret: ").strip()
    except KeyboardInterrupt:
        print("\nCancelled.")
        return 1
    if not secret:
        print("Empty input — nothing stored.", file=sys.stderr)
        return 1
    if not (secret.startswith("ntn_") or secret.startswith("secret_")):
        print("(Warning: secret doesn't start with 'ntn_' or 'secret_'. "
              "Stored anyway, but double-check you pasted the right thing.)")
    rc = _store(keyring, service, key_name, secret)
    if rc == 0:
        print()
        print("Next: run `python run_invoice_sync.py --dry-run`. On first access a")
        print("Keychain dialog asks whether to allow Python to read the key. Click")
        print("'Always Allow' — you won't be asked again for this binary.")
    return rc


def _setup_teams(keyring, service: str, from_env: bool) -> int:
    key_name = os.getenv("KEYSTORE_KEY_TEAMS_WEBHOOK", "teams_webhook_mfd_paid")
    print("Storing the TEAMS MFD paid/short-pay Workflows webhook URL.")
    print(f"  service  = {service}")
    print(f"  key name = {key_name}")
    print()

    if from_env:
        # Read the existing value straight from .env (never printed).
        secret = os.getenv("TEAMS_WEBHOOK_MFD_PAID", "").strip()
        if not secret:
            print("No TEAMS_WEBHOOK_MFD_PAID value found in .env — nothing to "
                  "migrate. Run without --from-env to paste it instead.",
                  file=sys.stderr)
            return 1
        print(f"Migrating the value already in .env ({len(secret)} chars, not "
              f"displayed) into Keychain…")
    else:
        print("Paste the Teams webhook URL and press Enter.")
        print("The paste is HIDDEN — you will not see it appear.")
        print("(Cancel with Ctrl-C if you change your mind.)")
        print()
        try:
            secret = getpass("Teams webhook URL: ").strip()
        except KeyboardInterrupt:
            print("\nCancelled.")
            return 1
        if not secret:
            print("Empty input — nothing stored.", file=sys.stderr)
            return 1
        if not secret.lower().startswith("https://"):
            print("(Warning: that doesn't look like an https:// URL. Stored "
                  "anyway, but double-check you pasted the right thing.)")

    rc = _store(keyring, service, key_name, secret)
    if rc != 0:
        return rc

    if from_env:
        # Only clear .env AFTER Keychain stored + verified above.
        cleared = _blank_env_line("TEAMS_WEBHOOK_MFD_PAID")
        if cleared:
            print("Cleared TEAMS_WEBHOOK_MFD_PAID from .env (value now lives only "
                  "in Keychain).")
        else:
            print("Note: .env line was already empty — nothing to clear.")

    print()
    print("Verify it still posts with: `python test_teams_webhook.py --paid`")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Store automation-worker secrets in the macOS Keychain."
    )
    parser.add_argument("--teams", action="store_true",
                        help="Store the Teams webhook URL instead of the Notion secret.")
    parser.add_argument("--from-env", action="store_true",
                        help="(with --teams) Migrate the existing "
                             "TEAMS_WEBHOOK_MFD_PAID value out of .env without "
                             "re-pasting, then blank the .env line.")
    args = parser.parse_args()

    if args.from_env and not args.teams:
        print("--from-env only applies to --teams.", file=sys.stderr)
        return 2

    if sys.platform != "darwin":
        print("This script is for macOS only. On the Pi/Docker, put secrets in "
              "the environment (.env.secrets / .env.docker) instead.",
              file=sys.stderr)
        return 2

    try:
        import keyring
    except ImportError:
        print("The `keyring` package isn't installed. Run:", file=sys.stderr)
        print("  pip install -r requirements.txt", file=sys.stderr)
        return 2

    service = os.getenv("KEYSTORE_SERVICE", "proficient-automation-worker")

    if args.teams:
        return _setup_teams(keyring, service, from_env=args.from_env)
    return _setup_notion(keyring, service)


if __name__ == "__main__":
    sys.exit(main())
