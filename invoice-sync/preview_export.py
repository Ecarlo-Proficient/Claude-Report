"""
preview_export.py — regenerate Open_Invoices.xlsx to a THROWAWAY path.

Read-only on Notion (the export only queries; it never writes). Safe to run any
time to eyeball the aging tabs before a real `sync-ar` touches the live OneDrive
copy. Use it to confirm the Open Balance / Total Amount columns, the per-row
data bar, and that your Excel Notes were preserved.

    cd "…/Automate Concrete Business/invoice-sync"
    python3 preview_export.py

Writes to ~/Library/Logs/Proficient/collections-notes-backup/PREVIEW_Open_Invoices.xlsx
and prints the path. Nothing in OneDrive or Notion is modified.
"""
from __future__ import annotations

import logging
import shutil
from pathlib import Path

from config import load_config
from notion_client import NotionClient
from export_invoices_xlsx import export_open_invoices_xlsx, DEFAULT_EXPORT_PATH

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

OUT = (
    Path.home()
    / "Library" / "Logs" / "Proficient" / "collections-notes-backup"
    / "PREVIEW_Open_Invoices.xlsx"
)


def main() -> int:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    # Seed the throwaway path from the LIVE file (if present) so preservation
    # reads your real Notes off it — the preview then shows BOTH the new layout
    # AND your Notes carried over. Only the throwaway OUT is ever written; the
    # live OneDrive file and Notion are untouched.
    if DEFAULT_EXPORT_PATH.exists():
        shutil.copy2(DEFAULT_EXPORT_PATH, OUT)
        print(f"Seeded preview from live file (to prove Note preservation): {DEFAULT_EXPORT_PATH}")
    cfg = load_config()
    notion = NotionClient(
        secret=cfg.notion_secret,
        api_base=cfg.notion_api_base,
        version=cfg.notion_version,
    )
    path = export_open_invoices_xlsx(
        notion=notion,
        res_com_ds_id=cfg.invoice_res_com_ds_id,
        mfd_ds_id=cfg.invoice_mfd_ds_id,
        customer_list_ds_id=cfg.customer_list_ds_id,
        mfd_client_list_ds_id=cfg.mfd_client_list_ds_id,
        output_path=OUT,
        # Show the real behaviour: cell Notes ABSORB into the Notes column and the
        # cell Note is dropped. apply_notion=False → the Notion push is only LOGGED
        # (see the "would push (dry-run)" lines above), never written.
        absorb_notes=True,
        apply_notion=False,
    )
    print(f"\nPreview written → {path}")
    print("In Excel, on the CP/MFD/RP Aging tabs:")
    print("  • 'Open Balance' then 'Total Amount' with a blue bar (no repair prompt).")
    print("  • Notes column now shows your cell-Note text (stamped '— Name, M/D'),")
    print("    and the yellow cell-Note markers are GONE (absorbed).")
    print("The lines above starting 'would push (dry-run)' are exactly what would be")
    print("written to Notion Quick Status — nothing was written this run.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
