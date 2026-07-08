"""
wip_sync.py — STUB during Excel pivot (2026-06-25).

The original wip_sync.py wrote QBO data into Notion WIP databases. That
approach is RETIRED. The new WIP source of truth is Excel on SharePoint
(`Company Files - WIP Report/WIP - MASTER.xlsx`), and the script will
write only to the "Test" tab via wip_excel_guard.py until Ted graduates
the rule.

This stub exists so:
  1. The launchd schedule (com.proficient.wip-sync.plist) doesn't error
     out while the new Excel-based sync is being written.
  2. The shell wrapper (run_wip_sync.sh) keeps working.
  3. Any old imports of wip_sync don't crash with ImportError.

The new implementation will land here, written from scratch against the
Excel-only target. Until then, this is a clean no-op.

Run: `python wip_sync.py` → prints message, exits 0.
"""
from __future__ import annotations

import logging
import sys


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    log = logging.getLogger("wip_sync")
    log.info("wip_sync.py: stub. Notion WIP path retired 2026-06-25. Excel sync pending.")
    log.info("Targets when implemented:")
    log.info("  READ: QBO API (Bills, Purchases, Invoices) — read-only HTTP GET")
    log.info("  WRITE: Excel WIP file 'Test' sheet ONLY (enforced by wip_excel_guard.py)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
