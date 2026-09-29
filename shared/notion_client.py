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


try:
    from shared import qbo_vault as kc
except ImportError:  # imported with shared/ itself on sys.path
    import qbo_vault as kc  # type: ignore


class NotionError(Exception):
    """Raised on non-retryable Notion API errors."""


def load_secret() -> str:
    """The Notion integration token: NOTION_SECRET env var, else the key library via Key Helper
    (shared/qbo_vault.get_secret - moved there from the older Keychain store 09/29/2026)."""
    try:
        secret = kc.get_secret("NOTION_SECRET")
    except kc.SecretsError as e:
        raise NotionError(f"Could not read the Notion token: {e}")
    if not secret:
        raise NotionError("No Notion token in the key library (NOTION_SECRET) - run "
                          "shared/setup_qbo.py --rotate NOTION_SECRET, or set NOTION_SECRET.")
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

    def block_children(self, block_id: str, page_size: int = 100) -> Iterator[dict]:
        """Yield the child blocks of a page/block (its body content), paginated.
        Used by load_customers.py to read the 'History of interactions' notes."""
        start = None
        while True:
            q = f"?page_size={page_size}" + (f"&start_cursor={start}" if start else "")
            data = self._request("GET", f"/blocks/{block_id}/children{q}")
            for block in data.get("results", []):
                yield block
            if not data.get("has_more"):
                return
            start = data.get("next_cursor")

    def create_page(self, ds_id: str, properties: dict, children: Optional[list] = None) -> dict:
        body = {"parent": {"type": "data_source_id", "data_source_id": ds_id}, "properties": properties}
        if children:
            body["children"] = children
        return self._request("POST", "/pages", body)

    def update_page(self, page_id: str, properties: dict) -> dict:
        return self._request("PATCH", f"/pages/{page_id}", {"properties": properties})

    def update_page_full(self, page_id: str, body: dict) -> dict:
        """PATCH a page with any top-level fields (properties, icon, archived…)."""
        return self._request("PATCH", f"/pages/{page_id}", body)

    def retrieve_data_source(self, ds_id: str) -> dict:
        return self._request("GET", f"/data_sources/{ds_id}")

    def append_children(self, block_id: str, children: list, after: Optional[str] = None) -> dict:
        """Append blocks under a page/block (max 100 per call, 2 nesting levels),
        optionally right after the sibling block `after`."""
        body: dict = {"children": children}
        if after:
            body["after"] = after
        return self._request("PATCH", f"/blocks/{block_id}/children", body)

    def update_block(self, block_id: str, body: dict) -> dict:
        return self._request("PATCH", f"/blocks/{block_id}", body)

    def delete_block(self, block_id: str) -> dict:
        """Move a block (and its children) to trash - recoverable from Notion's Trash."""
        return self._request("DELETE", f"/blocks/{block_id}")

    def upload_file(self, path, content_type: str = "application/octet-stream") -> str:
        """Single-part Notion file upload (files up to 20 MB). Returns the
        file_upload id to attach within the hour: {"type": "file_upload",
        "file_upload": {"id": ...}, "name": ...} on a files property."""
        from pathlib import Path as _P
        p = _P(path)
        up = self._request("POST", "/file_uploads",
                           {"filename": p.name, "content_type": content_type})
        headers = {k: v for k, v in self._headers.items() if k != "Content-Type"}
        with p.open("rb") as fh:
            r = self._session.post(up["upload_url"], headers=headers,
                                   files={"file": (p.name, fh, content_type)}, timeout=120)
        if not r.ok:
            raise NotionError(f"Notion upload {r.status_code}: {r.text[:300]}")
        return up["id"]
