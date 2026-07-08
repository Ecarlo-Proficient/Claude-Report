"""
Tier sync — Bid List → RP Field Log + CP Field Log.

Semantics (as of 2026-04-28):
  Field Log databases are STAGE-PER-ROW. Each physical job has ~9 stage
  rows (one per build phase), all tagged with the same Project #. Job-level
  fields (Division, Builder, Superintendent, Square Foot, etc.) are
  duplicated as copy-forward columns on every stage row. The user hides
  the duplication via group-by-Project # (or group-by-Job-Address) views.

  Stage rows are CREATED by a Notion button on the Bid List ("Send to
  Field Log") — never by this sync. This sync is UPDATE-ONLY: it walks
  every stage row matching a project's Project # and refreshes the
  copy-forward columns from the current Bid List values.

  Stages are field records. Even when the Bid List row is no longer
  tracked, the stage rows stay in place (they hold real-world phase
  progress). This sync NEVER archives. Active Status + Date Completed
  are owned by Field Log (set by the supers in the field) and are NEVER
  touched by this sync.

Eligibility — `Send to Field Log = true` is the gate.
  Bid List rows are only processed if the `Send to Field Log` checkbox is
  set. The button flips that checkbox to true after creating stage rows;
  the sync only sees rows the PM has actually pushed to Field Log. This
  keeps the working set small and avoids touching rows that aren't being
  tracked yet. (Common Bid List sync was removed 2026-04-28 — Bid List
  is shared directly with clerks now, so the Tier 1 separation is gone.)

Routing:
  - Division = Residential → RP Field Log
  - Division = Commercial  → CP Field Log
  - Other / missing        → skipped + logged (can't route)

Field Log key = title-derived Project #, untouched.
  Per Ted (2026-04-27): the `-FTW` suffix on a Project # is the canonical
  signal that a row is a flatwork variant of the main job, treated as its
  own logical project in Field Log. Most Bid List rows have titles like
  "RPxxxx - <address>" with no suffix; flatwork variants are titled
  "RPxxxx-FTW - <address>". The Job Type field is NOT consulted — Job
  Types like TRACT HOMES can also have an -FTW Project # because flatwork
  is the foundation work for any tract job.

  Both the button and this sync extract Project # by splitting the title
  on the FIRST " - " (space-dash-space) and using the head as-is.
  Whatever the button copied into Project # at create time is what we
  match against on update.

Match key on source:
  Bid List has NO separate Project # property. Its TITLE is "Job Name"
  formatted "<Project#> - <Description>". We split on the FIRST " - "
  to extract Project #. The descriptive half (typically the address) is
  discarded — Job Address is its own copy-forward field on Field Log.

Mode:
  Single mode: every run is a full sweep of Sent=true Bid List rows.
  Cheap because the working set is small (only actively-tracked
  projects), and UPDATEs are no-ops in Notion when properties haven't
  changed (Notion doesn't bump Last Edited unless the property value
  differs).

Per-group errors are caught, logged, counted. One bad group does not stop
the run. If ANY group errors, the caller should treat the run as partial.
"""
from __future__ import annotations

import logging
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

from notion_client import NotionClient, NotionError


log = logging.getLogger("automation_worker.field_log_sync")


FLOW_NAME = "bid_list_to_field_logs"


# -------- Source schema (Bid List) --------
SOURCE_JOB_NAME_TITLE = "Job Name"            # TITLE — "<Project#> - <Description>"
SOURCE_DIVISION = "Division"                  # select — "Residential" / "Commercial"
SOURCE_BUILDER = "Builder"                    # relation
SOURCE_SUPERINTENDENT = "Superintendent"      # people
SOURCE_SQUARE_FOOT = "Square Foot"            # number
SOURCE_FTW_SQ_FT = "FTW Sq. Ft."              # number
SOURCE_PIER_COUNT = "Pier Count"              # number
SOURCE_JOB_ADDRESS = "Job Address"            # rich_text
SOURCE_CITY = "City"                          # select on Bid List
SOURCE_SENT_TO_FIELD_LOG = "Send to Field Log"  # checkbox — sync gating signal (Ted's naming: action verb "Send", checkmark = it's been sent)

# Title delimiter — split on first occurrence only (Project # can contain '-').
PROJECT_NUM_DELIMITER = " - "

# Division values that route to each Field Log.
DIVISION_RESIDENTIAL = "Residential"
DIVISION_COMMERCIAL = "Commercial"


# -------- Target schema (RP / CP Field Log) --------
# Both Field Logs share identical copy-forward columns.
# Job Name and Job Type were dropped 2026-04-27. Job Address is the address
# of record (Job Name was duplicating it). Division alone is enough routing
# context — Job Type added select-option drift without driving any logic.
TARGET_PROJECT_NUM = "Project #"          # rich_text — match key
TARGET_DIVISION = "Division"              # select
TARGET_BUILDER = "Builder"                # rich_text (flattened from source relation)
TARGET_SUPERINTENDENT = "Superintendent"  # people
TARGET_SQUARE_FOOT = "Square Foot"        # number
TARGET_FTW_SQ_FT = "FTW Sq. Ft."          # number
TARGET_PIER_COUNT = "Pier Count"          # number
TARGET_JOB_ADDRESS = "Job Address"        # rich_text
TARGET_CITY = "City"                      # rich_text (Bid List select → Field Log text)

# Field Log owns these — sync NEVER writes them:
#   Active Status (select) — supers set in the field
#   Date Completed (date)  — supers set in the field
# Listed here only as a reminder of why they're absent from _build_merged_properties.


@dataclass
class FieldLogSyncSummary:
    # Source-row-level counters (over Sent=true rows only)
    source_rows_total: int = 0              # rows with Send to Field Log = true
    source_rows_skipped_title: int = 0      # title doesn't match pattern
    source_rows_unrouted: int = 0           # Division empty / unrecognised
    # Group-level counters (one unit = one unique Field Log key)
    groups_residential: int = 0
    groups_commercial: int = 0
    # Stage-row-level counters (the actual writes)
    stage_rows_updated: int = 0
    stage_rows_no_match: int = 0            # group had no stage rows in target — button hasn't fired yet for this row
    errors: list = field(default_factory=list)

    @property
    def had_errors(self) -> bool:
        return bool(self.errors)

    def as_dict(self) -> dict:
        return {
            "source_rows_total": self.source_rows_total,
            "source_rows_skipped_title": self.source_rows_skipped_title,
            "source_rows_unrouted": self.source_rows_unrouted,
            "groups_residential": self.groups_residential,
            "groups_commercial": self.groups_commercial,
            "stage_rows_updated": self.stage_rows_updated,
            "stage_rows_no_match": self.stage_rows_no_match,
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


def _extract_select_name(page: dict, prop_name: str) -> Optional[str]:
    prop = page.get("properties", {}).get(prop_name, {})
    select = prop.get("select")
    return select.get("name") if select else None


def _extract_number(page: dict, prop_name: str) -> Optional[float]:
    prop = page.get("properties", {}).get(prop_name, {})
    return prop.get("number")  # may be None


def _extract_people_ids(page: dict, prop_name: str) -> list:
    prop = page.get("properties", {}).get(prop_name, {})
    people = prop.get("people", []) or []
    return [p.get("id") for p in people if p.get("id")]


def _extract_relation_titles(client: NotionClient, page: dict, prop_name: str) -> list:
    prop = page.get("properties", {}).get(prop_name, {})
    related = prop.get("relation", []) or []
    names: list = []
    for rel in related:
        rel_id = rel.get("id")
        if not rel_id:
            continue
        try:
            related_page = client.retrieve_page(rel_id)
            title_str = _find_first_title(related_page)
            if title_str:
                names.append(title_str)
        except NotionError as e:
            log.warning("Could not fetch related page %s for flatten: %s", rel_id, e)
    return names


def _find_first_title(page: dict) -> str:
    for _name, prop in (page.get("properties") or {}).items():
        if prop.get("type") == "title":
            parts = prop.get("title", []) or []
            return "".join(p.get("plain_text", "") for p in parts).strip()
    return ""


def _split_job_name(job_name: str) -> tuple[str, str]:
    """Split '<Project#> - <Description>' on the FIRST ' - '. Project # may
    contain dashes (e.g. 'RP5982-FTW') so first-occurrence split matters.

    The delimiter is space-dash-space, never plain dash, so the dash inside
    'RPxxxx-FTW' is not a split point — the head correctly comes out as
    'RPxxxx-FTW' for flatwork variants.
    """
    if not job_name or PROJECT_NUM_DELIMITER not in job_name:
        return "", ""
    head, _, tail = job_name.partition(PROJECT_NUM_DELIMITER)
    return head.strip(), tail.strip()


# ---------- Property builders (for write) ----------

def _rich_text_prop(value: str) -> dict:
    if not value:
        return {"rich_text": []}
    return {"rich_text": [{"type": "text", "text": {"content": value}}]}


def _select_prop(value: Optional[str]) -> dict:
    if not value:
        return {"select": None}
    return {"select": {"name": value}}


def _number_prop(value: Optional[float]) -> dict:
    return {"number": value}


def _people_prop(user_ids: list) -> dict:
    return {"people": [{"object": "user", "id": uid} for uid in user_ids]}


# ---------- Aggregation helpers ----------

def _dedupe_preserve_order(items: list) -> list:
    seen = set()
    out: list = []
    for x in items:
        if x in seen:
            continue
        seen.add(x)
        out.append(x)
    return out


def _first_non_empty(values: list):
    """Return first non-empty value, or None.

    Empty means: None, "", [], {}. Notably 0 is NOT empty — it's a
    legitimate value for Pier Count ("no piers needed"), FTW Sq. Ft.
    ("no flatwork on this job"), etc. Treating 0 as empty would silently
    fall through to the next scope row's value (or None), which would
    overwrite an explicit zero with a missing value.
    """
    for v in values:
        if v is None:
            continue
        if isinstance(v, (str, list, dict)) and not v:
            continue
        return v
    return None


def _build_merged_properties(
    client: NotionClient,
    sold_scope_rows: list,
    field_log_key: str,
) -> dict:
    """
    Merge 1+ scope rows that share a Field Log key into a single properties
    dict for the Field Log target.

    Active Status + Date Completed are deliberately omitted — Field Log owns.
    Job Name and Job Type were dropped 2026-04-27.
    """
    division = _first_non_empty(
        [_extract_select_name(p, SOURCE_DIVISION) for p in sold_scope_rows]
    )
    address = _first_non_empty(
        [_extract_rich_text(p, SOURCE_JOB_ADDRESS) for p in sold_scope_rows]
    )
    city = _first_non_empty(
        [_extract_select_name(p, SOURCE_CITY) for p in sold_scope_rows]
    )
    square_foot = _first_non_empty(
        [_extract_number(p, SOURCE_SQUARE_FOOT) for p in sold_scope_rows]
    )
    ftw_sq_ft = _first_non_empty(
        [_extract_number(p, SOURCE_FTW_SQ_FT) for p in sold_scope_rows]
    )
    pier_count = _first_non_empty(
        [_extract_number(p, SOURCE_PIER_COUNT) for p in sold_scope_rows]
    )

    builder_names: list = []
    for page in sold_scope_rows:
        builder_names.extend(_extract_relation_titles(client, page, SOURCE_BUILDER))
    builder_joined = ", ".join(_dedupe_preserve_order(builder_names))

    supers: list = []
    for page in sold_scope_rows:
        supers.extend(_extract_people_ids(page, SOURCE_SUPERINTENDENT))
    supers_deduped = _dedupe_preserve_order(supers)

    return {
        TARGET_PROJECT_NUM: _rich_text_prop(field_log_key),
        TARGET_DIVISION: _select_prop(division),
        TARGET_BUILDER: _rich_text_prop(builder_joined),
        TARGET_SUPERINTENDENT: _people_prop(supers_deduped),
        TARGET_SQUARE_FOOT: _number_prop(square_foot),
        TARGET_FTW_SQ_FT: _number_prop(ftw_sq_ft),
        TARGET_PIER_COUNT: _number_prop(pier_count),
        TARGET_JOB_ADDRESS: _rich_text_prop(address or ""),
        TARGET_CITY: _rich_text_prop(city or ""),
    }


# ---------- Field Log update ----------

def _query_stage_rows_for_project(
    client: NotionClient,
    field_log_ds_id: str,
    project_num: str,
) -> list:
    """Return all stage rows in field_log_ds_id whose Project # rich_text
    equals project_num. May be 0 (button never pressed) or N (one per stage)."""
    filter_body = {
        "property": TARGET_PROJECT_NUM,
        "rich_text": {"equals": project_num},
    }
    return list(client.query_data_source(field_log_ds_id, filter_body=filter_body))


def _update_field_log_group(
    client: NotionClient,
    field_log_ds_id: str,
    field_log_label: str,           # "RP" or "CP" — for log clarity only
    field_log_key: str,
    sold_scope_rows: list,
    summary: FieldLogSyncSummary,
    dry_run: bool,
) -> None:
    """
    Find every stage row in the target Field Log matching field_log_key on
    Project #, and PATCH it with the merged copy-forward properties.

    UPDATE-only — never creates, never archives. If 0 stage rows match,
    that's odd (the row is Sent=true but no stage rows exist) — log it
    so the operator can investigate.
    """
    mapped = _build_merged_properties(client, sold_scope_rows, field_log_key)

    stage_rows = _query_stage_rows_for_project(client, field_log_ds_id, field_log_key)

    if not stage_rows:
        log.warning(
            "%s Field Log: Sent=true but no stage rows for Project # %r — "
            "stages may have been deleted, or the button's Project # formula "
            "produced a different value than expected",
            field_log_label, field_log_key,
        )
        summary.stage_rows_no_match += 1
        return

    log.info(
        "%s Field Log: updating %d stage rows for Project # %r",
        field_log_label, len(stage_rows), field_log_key,
    )

    for stage_page in stage_rows:
        stage_id = stage_page.get("id", "?")
        try:
            if dry_run:
                log.info(
                    "[DRY-RUN] %s Field Log UPDATE stage %s for Project # %r",
                    field_log_label, stage_id, field_log_key,
                )
            else:
                client.update_page(stage_id, mapped)
            summary.stage_rows_updated += 1
        except NotionError as e:
            log.error(
                "Notion error updating %s Field Log stage %s (Project # %r): %s",
                field_log_label, stage_id, field_log_key, e,
            )
            summary.errors.append(
                {
                    "field_log": field_log_label,
                    "project_num": field_log_key,
                    "stage_id": stage_id,
                    "error": str(e),
                }
            )


# ---------- Main sync ----------

def sync_bid_list_to_field_logs(
    client: NotionClient,
    bid_list_ds_id: str,
    rp_field_log_ds_id: str,
    cp_field_log_ds_id: str,
    dry_run: bool = False,
) -> FieldLogSyncSummary:
    """
    Run one full sweep over Sent=true Bid List rows + UPDATE pass.

    Steps:
      1. Walk Bid List rows where Send to Field Log = true. For each:
         - Skip if title doesn't match "<Project#> - <Description>".
         - Use title-derived Project # as the Field Log key (untouched —
           any -FTW suffix is preserved as part of the key).
         - Determine routing (Division: Residential→RP, Commercial→CP, else skip+log).
         - Bucket by (target_field_log, key).
      2. For each bucket: merge scope rows, query target stages, PATCH each.

    Returns FieldLogSyncSummary — caller decides what to do with errors.
    """
    summary = FieldLogSyncSummary()
    log.info(
        "Field Log sync start (Sent=true filter). dry_run=%s rp_ds=%s cp_ds=%s",
        dry_run, rp_field_log_ds_id, cp_field_log_ds_id,
    )

    # Pass 1: walk source filtered to Sent=true, bucket by (Field Log, key).
    # The "Send to Field Log" checkbox is the gating signal — only projects
    # the PM has actually pushed to Field Log get scanned.
    sent_filter = {
        "property": SOURCE_SENT_TO_FIELD_LOG,
        "checkbox": {"equals": True},
    }
    buckets: dict[tuple[str, str, str], list] = defaultdict(list)

    for source_page in client.query_data_source(bid_list_ds_id, filter_body=sent_filter):
        summary.source_rows_total += 1

        title = _extract_title(source_page, SOURCE_JOB_NAME_TITLE)
        proj_raw, _descriptive = _split_job_name(title)
        if not proj_raw:
            summary.source_rows_skipped_title += 1
            log.warning(
                "Sent=true row %s has malformed title %r — fix the title to "
                "'<Project#> - <Description>'",
                source_page.get("id", "?"), title,
            )
            continue

        # Field Log key = title-derived Project # as-is. Any "-FTW" suffix
        # is part of the key, identifying flatwork as its own logical project.
        key = proj_raw

        division = _extract_select_name(source_page, SOURCE_DIVISION)
        if division == DIVISION_RESIDENTIAL:
            target_ds = rp_field_log_ds_id
            label = "RP"
        elif division == DIVISION_COMMERCIAL:
            target_ds = cp_field_log_ds_id
            label = "CP"
        else:
            summary.source_rows_unrouted += 1
            log.warning(
                "Sent=true row %s has unroutable Division %r — skipping (need Residential or Commercial)",
                source_page.get("id", "?"), division,
            )
            continue

        bucket_key = (target_ds, label, key)
        buckets[bucket_key].append(source_page)

    # Count unique groups per division for summary.
    for (_ds, label, _key) in buckets.keys():
        if label == "RP":
            summary.groups_residential += 1
        elif label == "CP":
            summary.groups_commercial += 1

    log.info(
        "Scan complete: %d tracked rows (Sent=true), %d unroutable, "
        "%d unique RP groups, %d unique CP groups",
        summary.source_rows_total,
        summary.source_rows_unrouted,
        summary.groups_residential,
        summary.groups_commercial,
    )

    # Pass 2: per-bucket UPDATE. Sort for deterministic logs.
    for bucket_key in sorted(buckets.keys()):
        target_ds, label, key = bucket_key
        sold_scope_rows = buckets[bucket_key]
        try:
            _update_field_log_group(
                client=client,
                field_log_ds_id=target_ds,
                field_log_label=label,
                field_log_key=key,
                sold_scope_rows=sold_scope_rows,
                summary=summary,
                dry_run=dry_run,
            )
        except NotionError as e:
            log.error("Notion error processing %s group %r: %s", label, key, e)
            summary.errors.append(
                {"field_log": label, "project_num": key, "error": str(e)}
            )
        except Exception as e:
            log.exception("Unexpected error processing %s group %r", label, key)
            summary.errors.append(
                {"field_log": label, "project_num": key, "error": str(e)}
            )

    return summary
