"""
Diagnostic tool. Runs preflight checks and reports what's working / broken.

Runs in this order — stops at the first failure to keep output readable:

  1. Config loads (.env is present, required fields set).
  2. Secret resolves (Keychain on Mac, env on Pi). Length and prefix reported,
     NEVER the value itself.
  3. Notion auth round-trip (GET /users/me). Reports HTTP status only.
  4. Both data sources reachable. Queries each with page_size=1. Reports
     row count returned (should be >= 0) but not row contents.

Exit codes:
  0 = all green
  1 = at least one check failed
"""
from __future__ import annotations

import sys
from typing import Tuple

from config import load_config, Config
from notion_client import NotionClient, NotionError


def _ok(msg: str) -> None:
    print(f"  [OK]   {msg}")


def _fail(msg: str) -> None:
    print(f"  [FAIL] {msg}")


def _check_config() -> Tuple[bool, Config]:
    print("1. Config")
    try:
        cfg = load_config()
    except Exception as e:
        _fail(f"Could not load config: {e}")
        return False, None  # type: ignore

    _ok(f"Bid List DS ID: {cfg.bid_list_ds_id}")
    _ok(f"RP Field Log DS ID: {cfg.rp_field_log_ds_id}")
    _ok(f"CP Field Log DS ID: {cfg.cp_field_log_ds_id}")
    _ok(f"Project Plans DS ID: {cfg.project_plans_ds_id}")
    _ok(f"State dir: {cfg.state_dir}")
    _ok(f"Log dir: {cfg.log_dir}")
    _ok(f"Overlap seconds: {cfg.overlap_seconds}")
    return True, cfg


def _check_secret(cfg: Config) -> bool:
    print("\n2. Secret")
    s = cfg.notion_secret
    if not s:
        _fail("Secret resolved to empty string.")
        return False
    prefix = s[:4] if len(s) > 4 else "?"
    _ok(f"Secret length: {len(s)} chars, prefix: {prefix}... (value not displayed)")
    return True


def _check_auth(cfg: Config) -> Tuple[bool, NotionClient]:
    print("\n3. Notion auth")
    client = NotionClient(
        secret=cfg.notion_secret,
        api_base=cfg.notion_api_base,
        version=cfg.notion_version,
    )
    try:
        data = client._request("GET", "/users/me")
        bot_name = data.get("name", "(no name)")
        _ok(f"Auth OK — integration bot: {bot_name}")
        return True, client
    except NotionError as e:
        _fail(f"Auth failed: {e}")
        return False, client


def _check_data_source(client: NotionClient, label: str, ds_id: str) -> bool:
    print(f"\n4. Data source: {label}")
    try:
        # Just retrieve the data source — don't query rows. Pulls schema.
        data = client._request("POST", f"/data_sources/{ds_id}/query",
                               {"page_size": 1})
    except NotionError as e:
        _fail(f"{label} unreachable: {e}")
        return False

    count = len(data.get("results", []))
    _ok(f"{label} reachable. Sample query returned {count} row(s). "
        f"(Integration has access and schema is queryable.)")
    return True


def main() -> int:
    print("Proficient Automation Worker — doctor\n")

    ok, cfg = _check_config()
    if not ok:
        return 1

    if not _check_secret(cfg):
        return 1

    ok, client = _check_auth(cfg)
    if not ok:
        return 1

    all_ok = True
    all_ok &= _check_data_source(client, "Bid List", cfg.bid_list_ds_id)
    all_ok &= _check_data_source(client, "RP Field Log", cfg.rp_field_log_ds_id)
    all_ok &= _check_data_source(client, "CP Field Log", cfg.cp_field_log_ds_id)
    all_ok &= _check_data_source(client, "Project Plans", cfg.project_plans_ds_id)

    print()
    print("All checks passed." if all_ok else "Some checks failed.")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
