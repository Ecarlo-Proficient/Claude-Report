#!/usr/bin/env python3
"""
Invoice sync entrypoint — separate from sync.py because the cadence differs.

sync.py runs the Bid List → Field Log flow every 5 min.
run_invoice_sync.py runs the QBO → Invoice Tracker flow every 15 min
(launchd: com.proficient.invoice-sync.plist).

Usage:
    python run_invoice_sync.py            # live run
    python run_invoice_sync.py --dry-run  # preview only — no writes

Exit codes:
    0 — clean run
    1 — partial failure (one or more per-invoice errors)
    2 — fatal error (sync raised before completing)
"""
from __future__ import annotations

import argparse
import os
import sys

from config import load_config
from logger import setup_logging
from notion_client import NotionClient
from invoice_sync import (
    FLOW_NAME,
    InvoiceSyncTargets,
    sync_qbo_invoices_to_notion,
)
from export_invoices_xlsx import export_open_invoices_xlsx
from teams_notify import notify_sync_alert
from version import runtime_label


OUTCOME_CLEAN = 0
OUTCOME_PARTIAL = 1
OUTCOME_FATAL = 2


def main() -> int:
    parser = argparse.ArgumentParser(description="Proficient invoice sync (QBO → Notion)")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print planned creates/updates without calling Notion or QBO write endpoints",
    )
    args = parser.parse_args()

    rt = runtime_label()
    try:
        config = load_config()
    except Exception as e:
        print(f"FATAL: configuration error: {e}", file=sys.stderr)
        # Config failed before we could read the keystore, so the alert webhook
        # can only come from the environment (covers the Docker container).
        notify_sync_alert(
            os.getenv("TEAMS_WEBHOOK_ALERTS", "").strip(),
            runtime=rt, severity="error",
            title="FATAL — configuration error",
            detail=f"{type(e).__name__}: {e}",
        )
        return OUTCOME_FATAL

    log = setup_logging(config.log_dir)
    log.info("Starting %s [%s] dry_run=%s", FLOW_NAME, rt, args.dry_run)

    notion = NotionClient(
        secret=config.notion_secret,
        api_base=config.notion_api_base,
        version=config.notion_version,
    )

    targets = InvoiceSyncTargets(
        res_com_invoice_ds_id=config.invoice_res_com_ds_id,
        mfd_invoice_ds_id=config.invoice_mfd_ds_id,
        customer_list_ds_id=config.customer_list_ds_id,
        mfd_client_list_ds_id=config.mfd_client_list_ds_id,
    )

    log.info("---- Flow: %s ----", FLOW_NAME)
    try:
        res_com_summary, mfd_summary = sync_qbo_invoices_to_notion(
            notion=notion,
            targets=targets,
            paid_retention_months=config.invoice_paid_retention_months,
            dry_run=args.dry_run,
            teams_webhook_mfd_paid=config.teams_webhook_mfd_paid,
            state_dir=config.state_dir,
        )
    except Exception as e:
        log.exception("Fatal error in %s: %s", FLOW_NAME, e)
        notify_sync_alert(
            config.teams_webhook_alerts, runtime=rt, severity="error",
            title="FATAL — sync crashed before completing",
            detail=f"{type(e).__name__}: {e}",
        )
        return OUTCOME_FATAL

    log.info("Res/Com summary: %s", res_com_summary.as_dict())
    if res_com_summary.error_examples:
        log.warning("Res/Com error examples: %s", res_com_summary.error_examples)
    log.info("MFD summary: %s", mfd_summary.as_dict())
    if mfd_summary.error_examples:
        log.warning("MFD error examples: %s", mfd_summary.error_examples)

    # Excel export — runs even on partial failure (best-effort).
    # Skipped in dry-run since dry-run doesn't actually update Notion.
    # Skipped when SKIP_EXCEL_EXPORT=1 (Docker / Linux containers where OneDrive
    # isn't yet wired up via rclone — Mac launchd still owns Excel during Phase 1).
    import os as _os
    skip_excel = _os.getenv("SKIP_EXCEL_EXPORT", "").lower() in ("1", "true", "yes")
    if args.dry_run:
        log.info("Dry-run — skipping Excel export.")
    elif skip_excel:
        log.info("SKIP_EXCEL_EXPORT=1 — skipping Excel export (Docker / no-OneDrive mode).")
    else:
        try:
            log.info("---- Flow: export_open_invoices_xlsx ----")
            export_open_invoices_xlsx(
                notion=notion,
                res_com_ds_id=config.invoice_res_com_ds_id,
                mfd_ds_id=config.invoice_mfd_ds_id,
                customer_list_ds_id=config.customer_list_ds_id,
                mfd_client_list_ds_id=config.mfd_client_list_ds_id,
            )
        except Exception as e:
            # Don't fail the whole run if Excel export hiccups —
            # the sync's actual work (QBO→Notion) already succeeded.
            log.exception("Excel export failed (non-fatal): %s", e)

    if res_com_summary.had_errors or mfd_summary.had_errors:
        log.warning("%s partial failure.", FLOW_NAME)
        examples = (res_com_summary.error_examples or []) + (mfd_summary.error_examples or [])
        notify_sync_alert(
            config.teams_webhook_alerts, runtime=rt, severity="warning",
            title=(f"Completed with errors — "
                   f"Res/Com {res_com_summary.errors}, MFD {mfd_summary.errors}"),
            detail="; ".join(examples[:3]) or "see logs for detail",
        )
        return OUTCOME_PARTIAL

    log.info("%s clean.", FLOW_NAME)
    return OUTCOME_CLEAN


if __name__ == "__main__":
    sys.exit(main())
