"""
State store — tracks last_successful_run per sync flow.

File-based (JSON). One file per state store instance. Keyed by flow name
(e.g. "invoice_cdc_deletions") so multiple sync flows can share one store.

Advancement rule: timestamp is only advanced by the caller on fully clean runs.
On any per-row or fatal error, the caller leaves the timestamp alone — the next
run re-queries the same window. Upsert is idempotent, so re-processing is safe.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


class StateStore:
    def __init__(self, state_file: Path):
        self.state_file = state_file
        self._data: dict = self._load()

    def _load(self) -> dict:
        if not self.state_file.exists():
            return {}
        with self.state_file.open("r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                # Corrupt state file — treat as empty. Next save overwrites.
                return {}

    def _save(self) -> None:
        # Atomic-ish write: write to tmp then rename
        tmp = self.state_file.with_suffix(".tmp")
        with tmp.open("w", encoding="utf-8") as f:
            json.dump(self._data, f, indent=2)
        tmp.replace(self.state_file)

    def get_last_run(self, flow_name: str) -> Optional[str]:
        """Returns ISO 8601 UTC string, or None if flow has never run cleanly."""
        return self._data.get(flow_name, {}).get("last_successful_run")

    def set_last_run(self, flow_name: str, ts: datetime) -> None:
        """ts must be timezone-aware; stored as UTC ISO 8601."""
        if ts.tzinfo is None:
            raise ValueError("timestamp must be timezone-aware (use datetime.now(timezone.utc))")
        self._data.setdefault(flow_name, {})["last_successful_run"] = (
            ts.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        )
        self._save()
