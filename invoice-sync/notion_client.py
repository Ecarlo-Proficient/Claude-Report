"""
Thin HTTP wrapper for the Notion API.

Only implements what tier sync needs: query data source, retrieve page,
create page, update page. Handles auth + versioning + retry on 429.

Why not the notion-sdk-py package? Keeping dependencies minimal — requests
+ dotenv is all we need. Add the SDK later if a feature is genuinely missing.
"""
from __future__ import annotations

import logging
import time
from typing import Any, Iterator, Optional

import requests


log = logging.getLogger("automation_worker.notion")


class NotionError(Exception):
    """Raised on non-retryable Notion API errors."""


class NotionClient:
    def __init__(self, secret: str, api_base: str, version: str):
        self._headers = {
            "Authorization": f"Bearer {secret}",
            "Notion-Version": version,
            "Content-Type": "application/json",
        }
        self._api_base = api_base
        self._session = requests.Session()

    # ------------ low-level request helpers ------------

    # Retry budget — 429 gets a much bigger budget than transient errors
    # because Notion's rate limiter can hold the gate closed for a while
    # if multiple integrations share the workspace.
    _MAX_ATTEMPTS_429 = 10
    _MAX_ATTEMPTS_OTHER = 4
    _MAX_BACKOFF_S = 30

    def _request(self, method: str, path: str, json_body: Optional[dict] = None) -> dict:
        """Single request with 429 retry + 5xx retry. Raises NotionError on non-retryable failure."""
        url = f"{self._api_base}{path}"
        attempt = 0
        while True:
            attempt += 1
            try:
                resp = self._session.request(
                    method, url, headers=self._headers, json=json_body, timeout=30
                )
            except requests.RequestException as e:
                if attempt >= self._MAX_ATTEMPTS_OTHER:
                    raise NotionError(f"Network error after {attempt} attempts: {e}") from e
                backoff = min(2 ** attempt, self._MAX_BACKOFF_S)
                log.warning("Network error on %s %s (attempt %d): %s — retrying in %ds",
                            method, path, attempt, e, backoff)
                time.sleep(backoff)
                continue

            if resp.status_code == 429:
                # Respect Retry-After if Notion provides one, but apply exp
                # backoff on repeat 429s — Notion's Retry-After is often a
                # minimum, not a guarantee the next call will succeed.
                retry_after = int(resp.headers.get("Retry-After", "1"))
                backoff = min(max(retry_after, 2 ** attempt), self._MAX_BACKOFF_S)
                log.warning("Notion 429 (attempt %d/%d) — sleeping %ds",
                            attempt, self._MAX_ATTEMPTS_429, backoff)
                if attempt >= self._MAX_ATTEMPTS_429:
                    raise NotionError(f"Rate-limited after {attempt} attempts")
                time.sleep(backoff)
                continue

            if 500 <= resp.status_code < 600:
                if attempt >= self._MAX_ATTEMPTS_OTHER:
                    raise NotionError(f"Notion {resp.status_code}: {resp.text}")
                backoff = min(2 ** attempt, self._MAX_BACKOFF_S)
                log.warning("Notion %d on %s %s (attempt %d) — retrying in %ds",
                            resp.status_code, method, path, attempt, backoff)
                time.sleep(backoff)
                continue

            if not resp.ok:
                # 4xx — don't retry, surface immediately
                raise NotionError(
                    f"Notion {resp.status_code} on {method} {path}: {resp.text}"
                )

            return resp.json()

    # ------------ public surface ------------

    def query_data_source(
        self,
        ds_id: str,
        filter_body: Optional[dict] = None,
        sorts: Optional[list] = None,
        page_size: int = 100,
    ) -> Iterator[dict]:
        """
        Yields ALL pages from a data source matching the filter, handling pagination.
        """
        body: dict = {"page_size": page_size}
        if filter_body:
            body["filter"] = filter_body
        if sorts:
            body["sorts"] = sorts

        start_cursor = None
        while True:
            if start_cursor:
                body["start_cursor"] = start_cursor
            data = self._request("POST", f"/data_sources/{ds_id}/query", body)
            for page in data.get("results", []):
                yield page
            if not data.get("has_more"):
                return
            start_cursor = data.get("next_cursor")
            body.pop("start_cursor", None)  # rebuild on next iter

    def query_by_title(self, ds_id: str, title_prop_name: str, title_value: str) -> Optional[dict]:
        """
        Returns the first page in ds_id whose title_prop_name title equals title_value,
        or None. Used for upsert match-by-Project #.
        """
        filter_body = {
            "property": title_prop_name,
            "title": {"equals": title_value},
        }
        for page in self.query_data_source(ds_id, filter_body=filter_body, page_size=5):
            return page
        return None

    def retrieve_page(self, page_id: str) -> dict:
        return self._request("GET", f"/pages/{page_id}")

    def create_page(self, ds_id: str, properties: dict) -> dict:
        body = {
            "parent": {"type": "data_source_id", "data_source_id": ds_id},
            "properties": properties,
        }
        return self._request("POST", "/pages", body)

    def update_page(self, page_id: str, properties: dict) -> dict:
        return self._request("PATCH", f"/pages/{page_id}", {"properties": properties})

    def archive_page(self, page_id: str) -> dict:
        """
        Soft-delete (archive) a page. Notion does not hard-delete via the API;
        archived pages move to Trash and can be restored manually for 30 days.
        Currently unused — kept for future sync flows that need to remove
        derived rows when the source row is no longer eligible.
        """
        return self._request("PATCH", f"/pages/{page_id}", {"archived": True})
