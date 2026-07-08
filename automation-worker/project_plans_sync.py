"""
Project Plans sync — Bid List → Project Plans DB.

Single-target sync (one row per project) that keeps the Plans URL fresh on
the Project Plans database in the Field teamspace. The Notion checkbox-
automation creates the row at checkbox-time; this sync UPDATEs it whenever
the PM later edits Plans (or Job Address) on the Bid List.

Why a separate flow:
  Plans link can't live on Bid List for supers (no Bid List access for that
  role) and shouldn't live as a column on Field Log (would repeat across all
  9 stage rows per project). Project Plans is a separate Notion DB in the
  Field teamspace with one row per project: Project # (title), Job Address
  (text), Plans (URL). Supers access it directly.

Target schema (Project Plans):
  Project # (title)    — match key, set at create by automation, NOT touched here
  Job Address (text)   — copy-forward from Bid List
  Plans (url)          — copy-forward from Bid List

Source (Bid List): same gate as field_log_sync — `Send to Field Log = true`.

Match: title-derived Project # equals Project Plans row's title.

UPDATE-only. If a Send=true Bid List row has no matching Project Plans row,
log a warning — the automation should have created one when the checkbox
was toggled. Either the automation didn't fire (configuration error) or
the row was deleted manually.
"""
from __future__ import annotations

import logging
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

from notion_client import NotionClient, NotionError


log = logging.getLogger("automation_worker.project_plans_sync")


FLOW_NAME = "bid_list_to_project_plans"


# -------- Source schema (Bid List) --------
SOURCE_JOB_NAME_TITLE = "Job Name"
SOURCE_JOB_ADDRESS = "Job Address"
SOURCE_PLANS = "Plans"
SOURCE_SEND_TO_FIELD_LOG = "Send to Field Log"

# Title delimiter — split on first occurrence only (Project # can contain '-').
PROJECT_NUM_DELIMITER = " - "


# -------- Target schema (Project Plans) --------
TARGET_PROJECT_NUM = "Project #"     # title — match key, NOT written
TARGET_JOB_ADDRESS = "Job Address"   # rich_text
TARGET_PLANS = "Plans"               # url


@dataclass
class ProjectPlansSyncSummary:
    source_rows_total: int = 0
    source_rows_skipped_title: int = 0
    groups_total: int = 0                  # one per unique title-derived Project #
    target_rows_updated: int = 0
    target_rows_no_match: int = 0          # Send=true but no Project Plans row exists
    errors: list = field(default_factory=list)

    @property
    def had_errors(self) -> bool:
        return bool(self.errors)

    def as_dict(self) -> dict:
        return {
            "source_rows_total": self.source_rows_total,
            "source_rows_skipped_title": self.source_rows_skipped_title,
            "groups_total": self.groups_total,
            "target_rows_updated": self.target_rows_updated,
            "target_rows_no_match": self.target_rows_no_match,
            "errors": self.errors,
        }


# ---------- Property extraction helpers ----------

def _extract_title(page: dict, prop_name: str) -> str:
    prop = page.get("properties", {}).get(prop_name, {})
    parts = prop.get("title", []) or []
    return "".join(p.get("plain_text", "") for p in parts).strip()


def _extract_rich_text(page: dict, prop_name: str) -> str:
    prop = page.get("properties", {}).get(prop_name, {})
    parts = prop.get("rich_text", []) or []
    return "".join(p.get("plain_text", "") for p in parts).strip()


def _extract_url(page: dict, prop_name: str) -> Optional[str]:
    prop = page.get("properties", {}).get(prop_name, {})
    return prop.get("url")  # may be None


def _split_job_name(job_name: str) -> tuple[str, str]:
    """Split '<Project#> - <Description>' on the FIRST ' - '. Same logic as
    field_log_sync — Project # may contain dashes (e.g. RP5982-FTW), and the
    delimiter is space-dash-space, never plain dash."""
    if not job_name or PROJECT_NUM_DELIMITER not in job_name:
        return "", ""
    head, _, tail = job_name.partition(PROJECT_NUM_DELIMITER)
    return head.strip(), tail.strip()


def _first_non_empty(values):
    """Same shape as field_log_sync._first_non_empty. None / "" / [] / {}
    treated as empty; numeric 0 is NOT empty."""
    for v in values:
        if v is None:
            continue
        if isinstance(v, (str, list, dict)) and not v:
            continue
        return v
    return None


# ---------- Property builders (for write) ----------

def _rich_text_prop(value: str) -> dict:
    if not value:
        return {"rich_text": []}
    return {"rich_text": [{"type": "text", "text": {"content": value}}]}


def _url_prop(value: Optional[str]) -> dict:
    # Notion accepts None/null to clear a URL; empty string is rejected.
    return {"url": value or None}


# ---------- Update logic ----------

def _build_merged_properties(sold_scope_rows: list) -> dict:
    """
    Merge 1+ Bid List scope rows that share a Project # into a single
    properties dict for Project Plans. First-non-empty wins for each field
    across scope rows. Project # (title) is NOT included — the automation
    set it at create time and we never change it.
    """
    address = _first_non_empty(
        [_extract_rich_text(p, SOURCE_JOB_ADDRESS) for p in sold_scope_rows]
    )
    plans_url = _first_non_empty(
        [_extract_url(p, SOURCE_PLANS) for p in sold_scope_rows]
    )
    return {
        TARGET_JOB_ADDRESS: _rich_text_prop(address or ""),
        TARGET_PLANS: _url_prop(plans_url),
    }


def _update_project_plans_row(
    client: NotionClient,
    project_plans_ds_id: str,
    project_num: str,
    sold_scope_rows: list,
    summary: ProjectPlansSyncSummary,
    dry_run: bool,
) -> None:
    """
    Find the Project Plans row whose title equals project_num, PATCH its
    Job Address + Plans fields. UPDATE-only — if no row exists, log a
    warning so the operator can investigate (the automation should have
    created one when the checkbox was toggled).
    """
    target_row = client.query_by_title(
        project_plans_ds_id, TARGET_PROJECT_NUM, project_num
    )
    if not target_row:
        log.warning(
            "Project Plans: Send=true on Bid List but no Project Plans row "
            "for Project # %r — automation may not have fired, or the row was deleted",
            project_num,
        )
        summary.target_rows_no_match += 1
        return

    mapped = _build_merged_properties(sold_scope_rows)
    target_id = target_row.get("id", "?")

    try:
        if dry_run:
            log.info(
                "[DRY-RUN] Project Plans UPDATE row %s for Project # %r",
                target_id, project_num,
            )
        else:
            client.update_page(target_id, mapped)
        summary.target_rows_updated += 1
    except NotionError as e:
        log.error(
            "Notion error updating Project Plans row %s (Project # %r): %s",
            target_id, project_num, e,
        )
        summary.errors.append(
            {"project_num": project_num, "row_id": target_id, "error": str(e)}
        )


# ---------- Main sync ----------

def sync_bid_list_to_project_plans(
    client: NotionClient,
    bid_list_ds_id: str,
    project_plans_ds_id: str,
    dry_run: bool = False,
) -> ProjectPlansSyncSummary:
    """
    Run one full sweep over Send=true Bid List rows + UPDATE pass on Project Plans.

    Steps:
      1. Walk Bid List rows where Send to Field Log = true. For each:
         - Skip if title doesn't match "<Project#> - <Description>".
         - Bucket by title-derived Project #.
      2. For each unique Project #:
         - Merge Job Address + Plans across scope rows (first-non-empty).
         - Look up matching Project Plans row by title equality.
         - PATCH Job Address + Plans, or warn if not found.

    Returns ProjectPlansSyncSummary — caller decides what to do with errors.
    """
    summary = ProjectPlansSyncSummary()
    log.info(
        "Project Plans sync start (Send=true filter). dry_run=%s plans_ds=%s",
        dry_run, project_plans_ds_id,
    )

    sent_filter = {
        "property": SOURCE_SEND_TO_FIELD_LOG,
        "checkbox": {"equals": True},
    }
    buckets: dict[str, list] = defaultdict(list)

    for source_page in client.query_data_source(bid_list_ds_id, filter_body=sent_filter):
        summary.source_rows_total += 1

        title = _extract_title(source_page, SOURCE_JOB_NAME_TITLE)
        proj_raw, _descriptive = _split_job_name(title)
        if not proj_raw:
            summary.source_rows_skipped_title += 1
            log.warning(
                "Send=true row %s has malformed title %r — fix the title to "
                "'<Project#> - <Description>'",
                source_page.get("id", "?"), title,
            )
            continue

        # Project Plans key = title-derived Project # untouched (any -FTW stays).
        # Same convention as Field Log key — flatwork variants get their own
        # Project Plans row keyed `<base>-FTW`.
        buckets[proj_raw].append(source_page)

    summary.groups_total = len(buckets)
    log.info(
        "Scan complete: %d tracked rows (Send=true), %d unique projects",
        summary.source_rows_total,
        summary.groups_total,
    )

    for project_num in sorted(buckets.keys()):
        sold_scope_rows = buckets[project_num]
        try:
            _update_project_plans_row(
                client=client,
                project_plans_ds_id=project_plans_ds_id,
                project_num=project_num,
                sold_scope_rows=sold_scope_rows,
                summary=summary,
                dry_run=dry_run,
            )
        except NotionError as e:
            log.error("Notion error processing Project # %r: %s", project_num, e)
            summary.errors.append({"project_num": project_num, "error": str(e)})
        except Exception as e:
            log.exception("Unexpected error processing Project # %r", project_num)
            summary.errors.append({"project_num": project_num, "error": str(e)})

    return summary
