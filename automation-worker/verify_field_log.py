#!/usr/bin/env python3
"""
Field Log consistency verifier — READ-ONLY.

What it checks:
  1. ORPHAN STAGES        — stages in a Field Log whose Project # has no
                            corresponding Sent=true row on Bid List.
                            (Sent flag got unchecked, or Bid List row
                            was deleted, but the stage rows survived.)
  2. MISSING STAGES       — Sent=true Bid List rows with ZERO stage rows
                            in the routed Field Log. (Stage rows got
                            manually deleted after the button created them.)
  3. MISROUTED STAGES     — Project # exists in RP Field Log but Bid List
                            Division = Commercial, or vice versa.
  4. STALE COPY-FORWARD   — Stage row's copy-forward field disagrees with
                            the merged value field_log_sync would compute.
                            Reports per-field diffs with project + stage IDs.

Exit codes:
  0 — clean
  1 — issues found (any of the four categories)
  2 — fatal error (config / network)

Usage:
  python verify_field_log.py
  python verify_field_log.py --quiet           # only summary, no per-issue lines
  python verify_field_log.py --max-issues 50   # cap how many issues print

This script never writes. Safe to run anytime, including against prod.
"""
from __future__ import annotations

import argparse
import logging
import sys
from collections import defaultdict
from typing import Optional

from config import load_config
from logger import setup_logging
from notion_client import NotionClient, NotionError

import field_log_sync as fls


log = logging.getLogger("automation_worker.verify_field_log")


# ---------- Helpers ----------

def _normalise_for_compare(value):
    """Coerce empty-ish values to a single canonical form (None) so that
    e.g. '' vs None doesn't show up as a spurious diff."""
    if value in ("", [], {}):
        return None
    return value


def _extract_target_state(stage_page: dict) -> dict:
    """Read the copy-forward fields off a Field Log stage row, normalised
    for comparison against what field_log_sync would compute."""
    return {
        "Project #": _normalise_for_compare(
            fls._extract_rich_text(stage_page, fls.TARGET_PROJECT_NUM)
        ),
        "Division": _normalise_for_compare(
            fls._extract_select_name(stage_page, fls.TARGET_DIVISION)
        ),
        "Builder": _normalise_for_compare(
            fls._extract_rich_text(stage_page, fls.TARGET_BUILDER)
        ),
        "Superintendent": tuple(
            sorted(fls._extract_people_ids(stage_page, fls.TARGET_SUPERINTENDENT))
        ) or None,
        "Square Foot": fls._extract_number(stage_page, fls.TARGET_SQUARE_FOOT),
        "FTW Sq. Ft.": fls._extract_number(stage_page, fls.TARGET_FTW_SQ_FT),
        "Pier Count": fls._extract_number(stage_page, fls.TARGET_PIER_COUNT),
        "Job Address": _normalise_for_compare(
            fls._extract_rich_text(stage_page, fls.TARGET_JOB_ADDRESS)
        ),
        "City": _normalise_for_compare(
            fls._extract_rich_text(stage_page, fls.TARGET_CITY)
        ),
    }


def _extract_expected_state(properties_dict: dict) -> dict:
    """Convert a properties dict (as built by _build_merged_properties) into
    the same flat shape as _extract_target_state for comparison."""
    def _rt(prop):
        parts = prop.get("rich_text", []) or []
        text = "".join(p.get("text", {}).get("content", "") for p in parts)
        return _normalise_for_compare(text)

    def _sel(prop):
        sel = prop.get("select")
        return _normalise_for_compare(sel.get("name") if sel else None)

    def _num(prop):
        return prop.get("number")

    def _ppl(prop):
        return tuple(sorted(p.get("id") for p in prop.get("people", []) or [])) or None

    return {
        "Project #": _rt(properties_dict[fls.TARGET_PROJECT_NUM]),
        "Division": _sel(properties_dict[fls.TARGET_DIVISION]),
        "Builder": _rt(properties_dict[fls.TARGET_BUILDER]),
        "Superintendent": _ppl(properties_dict[fls.TARGET_SUPERINTENDENT]),
        "Square Foot": _num(properties_dict[fls.TARGET_SQUARE_FOOT]),
        "FTW Sq. Ft.": _num(properties_dict[fls.TARGET_FTW_SQ_FT]),
        "Pier Count": _num(properties_dict[fls.TARGET_PIER_COUNT]),
        "Job Address": _rt(properties_dict[fls.TARGET_JOB_ADDRESS]),
        "City": _rt(properties_dict[fls.TARGET_CITY]),
    }


def _diff_states(actual: dict, expected: dict) -> list[tuple[str, object, object]]:
    """Return list of (field_name, actual_value, expected_value) for any
    field that differs. Empty list = consistent."""
    diffs: list[tuple[str, object, object]] = []
    for k, exp in expected.items():
        act = actual.get(k)
        if act != exp:
            diffs.append((k, act, exp))
    return diffs


# ---------- Verifier ----------

def verify(
    client: NotionClient,
    bid_list_ds_id: str,
    rp_field_log_ds_id: str,
    cp_field_log_ds_id: str,
    quiet: bool,
    max_issues: int,
) -> int:
    """Run all four checks. Return number of issues found."""

    # --- Build "expected world" from Bid List ---
    # Same logic as field_log_sync.sync_bid_list_to_field_logs, but stops
    # before writing. We need a map: (target_label, key) → expected props.
    expected: dict[tuple[str, str], tuple[dict, list]] = {}
    # And per-key: what division *should* this be routed to?
    expected_routing: dict[str, str] = {}  # key → "RP" or "CP"

    buckets: dict[tuple[str, str], list] = defaultdict(list)

    sent_filter = {
        "property": fls.SOURCE_SENT_TO_FIELD_LOG,
        "checkbox": {"equals": True},
    }
    for source_page in client.query_data_source(bid_list_ds_id, filter_body=sent_filter):
        title = fls._extract_title(source_page, fls.SOURCE_JOB_NAME_TITLE)
        proj_raw, _descriptive = fls._split_job_name(title)
        if not proj_raw:
            continue
        # Field Log key = title-derived Project # untouched (any -FTW stays).
        key = proj_raw
        division = fls._extract_select_name(source_page, fls.SOURCE_DIVISION)
        if division == fls.DIVISION_RESIDENTIAL:
            label = "RP"
        elif division == fls.DIVISION_COMMERCIAL:
            label = "CP"
        else:
            # Sold-but-unroutable: still record so we can flag missing
            # stages and misroutes, but use a sentinel label.
            label = "?"
        bucket_key = (label, key)
        buckets[bucket_key].append(source_page)
        # Last-write-wins on routing (a key SHOULD only ever get one label
        # — if not, that itself is a misroute we'll catch below).
        expected_routing[key] = label

    log.info("Built expected world: %d Sold groups across both Field Logs", len(buckets))

    # Compute expected properties for each bucket.
    for (label, key), rows in buckets.items():
        if label == "?":
            continue  # can't compute target props if we don't know which DB
        try:
            props = fls._build_merged_properties(client, rows, key)
            expected[(label, key)] = (props, rows)
        except NotionError as e:
            log.warning("Could not build expected props for %s/%s: %s", label, key, e)

    # --- Walk actual Field Log stage rows ---
    rp_stages: dict[str, list] = defaultdict(list)
    cp_stages: dict[str, list] = defaultdict(list)
    for stage in client.query_data_source(rp_field_log_ds_id):
        proj = fls._extract_rich_text(stage, fls.TARGET_PROJECT_NUM)
        if proj:
            rp_stages[proj].append(stage)
    for stage in client.query_data_source(cp_field_log_ds_id):
        proj = fls._extract_rich_text(stage, fls.TARGET_PROJECT_NUM)
        if proj:
            cp_stages[proj].append(stage)

    log.info(
        "Loaded actual world: %d unique RP projects (%d stages), %d unique CP projects (%d stages)",
        len(rp_stages), sum(len(s) for s in rp_stages.values()),
        len(cp_stages), sum(len(s) for s in cp_stages.values()),
    )

    issues = 0
    printed = 0

    def _print_issue(msg: str):
        nonlocal printed
        if quiet or printed >= max_issues:
            return
        log.info(msg)
        printed += 1

    # --- Check 1: ORPHAN STAGES ---
    log.info("---- Check 1: orphan stages (Field Log Project # not Sold on Bid List) ----")
    eligible_keys = set(k for (_label, k) in expected.keys())
    orphans_rp = sorted(set(rp_stages.keys()) - eligible_keys)
    orphans_cp = sorted(set(cp_stages.keys()) - eligible_keys)
    for key in orphans_rp:
        issues += 1
        _print_issue(f"  ORPHAN [RP] Project # {key!r} — {len(rp_stages[key])} stage(s) but no Sold Bid List row")
    for key in orphans_cp:
        issues += 1
        _print_issue(f"  ORPHAN [CP] Project # {key!r} — {len(cp_stages[key])} stage(s) but no Sold Bid List row")
    log.info("Found %d orphan project(s) (%d RP, %d CP)",
             len(orphans_rp) + len(orphans_cp), len(orphans_rp), len(orphans_cp))

    # --- Check 2: MISSING STAGES ---
    log.info("---- Check 2: missing stages (Sold project but no stage rows in routed Field Log) ----")
    missing = 0
    for (label, key) in sorted(expected.keys()):
        actual_stages = rp_stages if label == "RP" else cp_stages
        if not actual_stages.get(key):
            missing += 1
            issues += 1
            _print_issue(
                f"  MISSING [{label}] Project # {key!r} — Sold on Bid List but no stage rows "
                f"(button never pressed)"
            )
    log.info("Found %d missing-stage project(s)", missing)

    # --- Check 3: MISROUTED STAGES ---
    log.info("---- Check 3: misrouted stages (project in wrong Field Log for its Division) ----")
    misroutes = 0
    for proj, stages in rp_stages.items():
        expected_label = expected_routing.get(proj)
        if expected_label and expected_label != "RP":
            misroutes += 1
            issues += 1
            _print_issue(
                f"  MISROUTE Project # {proj!r} has {len(stages)} stage(s) in RP Field Log "
                f"but Bid List Division routes it to {expected_label}"
            )
    for proj, stages in cp_stages.items():
        expected_label = expected_routing.get(proj)
        if expected_label and expected_label != "CP":
            misroutes += 1
            issues += 1
            _print_issue(
                f"  MISROUTE Project # {proj!r} has {len(stages)} stage(s) in CP Field Log "
                f"but Bid List Division routes it to {expected_label}"
            )
    log.info("Found %d misrouted project(s)", misroutes)

    # --- Check 4: STALE COPY-FORWARD ---
    log.info("---- Check 4: stale copy-forward fields ----")
    stale = 0
    for (label, key), (expected_props, _rows) in sorted(expected.items()):
        actual_stages = rp_stages if label == "RP" else cp_stages
        stages_for_key = actual_stages.get(key, [])
        if not stages_for_key:
            continue  # already counted as missing
        expected_state = _extract_expected_state(expected_props)
        for stage in stages_for_key:
            actual_state = _extract_target_state(stage)
            diffs = _diff_states(actual_state, expected_state)
            if diffs:
                stale += 1
                issues += 1
                fields_str = "; ".join(
                    f"{f}: actual={a!r} expected={e!r}" for (f, a, e) in diffs
                )
                _print_issue(
                    f"  STALE [{label}] Project # {key!r} stage {stage.get('id', '?')[:8]}…  "
                    f"→ {fields_str}"
                )
    log.info("Found %d stale stage row(s)", stale)

    # --- Summary ---
    log.info(
        "Verification complete: %d total issue(s) [orphans=%d, missing=%d, misroutes=%d, stale=%d]",
        issues, len(orphans_rp) + len(orphans_cp), missing, misroutes, stale,
    )
    if printed < issues and not quiet:
        log.info("(Output truncated at %d issues — pass --max-issues N to show more)", max_issues)

    return issues


def main() -> int:
    parser = argparse.ArgumentParser(description="Field Log consistency verifier (read-only)")
    parser.add_argument("--quiet", action="store_true", help="Only print summary lines")
    parser.add_argument(
        "--max-issues", type=int, default=20,
        help="Maximum issues to print per category before truncating (default 20)",
    )
    args = parser.parse_args()

    try:
        config = load_config()
    except Exception as e:
        print(f"FATAL: configuration error: {e}", file=sys.stderr)
        return 2

    setup_logging(config.log_dir)

    client = NotionClient(
        secret=config.notion_secret,
        api_base=config.notion_api_base,
        version=config.notion_version,
    )

    try:
        issues = verify(
            client=client,
            bid_list_ds_id=config.bid_list_ds_id,
            rp_field_log_ds_id=config.rp_field_log_ds_id,
            cp_field_log_ds_id=config.cp_field_log_ds_id,
            quiet=args.quiet,
            max_issues=args.max_issues,
        )
    except Exception as e:
        log.exception("Fatal error during verification: %s", e)
        return 2

    return 1 if issues else 0


if __name__ == "__main__":
    sys.exit(main())
