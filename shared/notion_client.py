"""
notion_client.py — the shared Notion API client.

A thin HTTP wrapper (query / create / update pages against /data_sources),
the same pattern invoice-sync uses, graduated to shared/ so the ledger's
sync_actions.py can push action items to Notion. invoice-sync keeps its own
tool-local copy on purpose (historical, per CLAUDE.md); NEW tools use THIS one.

Auth: the Notion integration secret from the Keychain (service
'proficient-automation-worker', key 'notion') via `keyring`, or NOTION_SECRET.
API version 2025-09-03 (required for the /data_sources/{id}/query endpoint).
"""
from __future__ import annotations

import os
import time
from typing import Iterator, Optional

import requests

API_BASE = "https://api.notion.com/v1"
VERSION = "2025-09-03"


class NotionError(Exception):
    """Raised on non-retryable Notion API errors."""


def load_secret() -> str:
    """NOTION_SECRET env var, else the Keychain (proficient-automation-worker/notion)."""
    env = os.getenv("NOTION_SECRET")
    if env:
        return env
    service = os.getenv("KEYSTORE_SERVICE", "proficient-automation-worker")
    key = os.getenv("KEYSTORE_KEY_NOTION", "notion")
    try:
        import keyring
        secret = keyring.get_password(service, key)
    except Exception as e:  # noqa: BLE001
        raise NotionError(f"Could not read Notion secret from Keychain ({service}/{key}): {e}")
    if not secret:
        raise NotionError(f"No Notion secret in Keychain ({service}/{key}); set it or NOTION_SECRET.")
    return secret


class NotionClient:
    _MAX_ATTEMPTS_429 = 8
    _MAX_ATTEMPTS_OTHER = 4
    _MAX_BACKOFF_S = 30

    def __init__(self, secret: Optional[str] = None, api_base: str = API_BASE, version: str = VERSION):
        self._headers = {
            "Authorization": f"Bearer {secret or load_secret()}",
            "Notion-Version": version,
            "Content-Type": "application/json",
        }
        self._api_base = api_base
        self._session = requests.Session()

    def _request(self, method: str, path: str, json_body: Optional[dict] = None) -> dict:
        url = f"{self._api_base}{path}"
        attempt = 0
        while True:
            attempt += 1
            try:
                r = self._session.request(method, url, headers=self._headers, json=json_body, timeout=30)
            except requests.RequestException as e:
                if attempt >= self._MAX_ATTEMPTS_OTHER:
                    raise NotionError(f"Network error after {attempt} attempts: {e}") from e
                time.sleep(min(2 ** attempt, self._MAX_BACKOFF_S)); continue
            if r.status_code == 429:
                if attempt >= self._MAX_ATTEMPTS_429:
                    raise NotionError(f"Rate-limited after {attempt} attempts")
                time.sleep(min(max(int(r.headers.get("Retry-After", "1")), 2 ** attempt), self._MAX_BACKOFF_S)); continue
            if 500 <= r.status_code < 600:
                if attempt >= self._MAX_ATTEMPTS_OTHER:
                    raise NotionError(f"Notion {r.status_code}: {r.text[:300]}")
                time.sleep(min(2 ** attempt, self._MAX_BACKOFF_S)); continue
            if not r.ok:
                raise NotionError(f"Notion {r.status_code} on {method} {path}: {r.text[:400]}")
            return r.json()

    def query_data_source(self, ds_id: str, filter_body: Optional[dict] = None,
                          sorts: Optional[list] = None, page_size: int = 100) -> Iterator[dict]:
        body: dict = {"page_size": page_size}
        if filter_body:
            body["filter"] = filter_body
        if sorts:
            body["sorts"] = sorts
        start = None
        while True:
            if start:
                body["start_cursor"] = start
            data = self._request("POST", f"/data_sources/{ds_id}/query", body)
            for page in data.get("results", []):
                yield page
            if not data.get("has_more"):
                return
            start = data.get("next_cursor")
            body.pop("start_cursor", None)

    def query_by_property(self, ds_id: str, prop: str, kind: str, value) -> Optional[dict]:
        """First page where property `prop` (of Notion type `kind`, e.g. 'rich_text',
        'title') equals `value`, else None. Used for upsert-by-Action-Key."""
        flt = {"property": prop, kind: {"equals": value}}
        for page in self.query_data_source(ds_id, filter_body=flt, page_size=3):
            return page
        return None

    def retrieve_page(self, page_id: str) -> dict:
        return self._request("GET", f"/pages/{page_id}")

    def create_page(self, ds_id: str, properties: dict, children: Optional[list] = None) -> dict:
        body = {"parent": {"type": "data_source_id", "data_source_id": ds_id}, "properties": properties}
        if children:
            body["children"] = children
        return self._request("POST", "/pages", body)

    def update_page(self, page_id: str, properties: dict) -> dict:
        return self._request("PATCH", f"/pages/{page_id}", {"properties": properties})
