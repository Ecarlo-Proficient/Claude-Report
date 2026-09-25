"""
qbo_cache.py — one QBO session, one pull per entity, shared by consumers.

The company-health tools were each opening their own QBO session and
re-downloading the same tables (Invoice pulled 3×, Bill 3×, BillPayment 3×,
Payment/Purchase 2× …), plus money_bleeds made ONE API CALL PER PROJECT to get
a project's invoices. This caches per (entity, where) for the life of a run and
exposes pre-built indexes, so a caller asks for what it needs and pays for the
download at most once (the user 2026-07-24).

Scope: the company-health audit tools. It does not change how invoice-sync,
bill-tracker or the WIP readers authenticate — they keep their own paths.

USAGE
    from shared.qbo_cache import QboCache
    qc = QboCache()                       # lazy: no auth until first use
    access, cid = qc.credentials()        # one Touch ID for the whole run
    open_invoices = qc.invoices(where="Balance > '0'")
    by_cust = qc.invoices_by_customer()   # {customer_id: [invoice, …]}
    pmap = qc.project_customer_map()      # {'MFD177': {id, name, …}, …}

Every accessor returns the SAME list object on repeat calls — treat results as
read-only. Call `qc.stats()` to see what was pulled vs served from cache.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

from shared import qbo_api


class QboCache:
    """Per-run QBO pull cache. Not thread-safe; one instance per run."""

    def __init__(self, access: Optional[str] = None,
                 company_id: Optional[str] = None, verbose: bool = False):
        self._access = access
        self._cid = company_id
        self.verbose = verbose
        self._entities: Dict[Tuple[str, str], List[dict]] = {}
        self._derived: Dict[str, Any] = {}
        self._pulls = 0
        self._hits = 0

    # ── auth ──
    def credentials(self) -> Tuple[str, str]:
        """Authenticate once; every later call reuses the same token."""
        if self._access is None or self._cid is None:
            self._access, self._cid = qbo_api.load_credentials()
        return self._access, self._cid

    # ── generic entity pull ──
    def entity(self, name: str, where: str = "") -> List[dict]:
        """All rows of a QBO entity, cached by (entity, where)."""
        key = (name, where or "")
        if key in self._entities:
            self._hits += 1
            return self._entities[key]
        access, cid = self.credentials()
        rows = qbo_api.query_all(access, cid, name, where)
        self._entities[key] = rows
        self._pulls += 1
        if self.verbose:
            print(f"      [qbo_cache] pulled {len(rows):>5} {name}"
                  f"{' where ' + where if where else ''}")
        return rows

    # ── convenience accessors (thin, so intent reads clearly) ──
    def invoices(self, where: str = "") -> List[dict]:
        return self.entity("Invoice", where)

    def bills(self, where: str = "") -> List[dict]:
        return self.entity("Bill", where)

    def bill_payments(self, where: str = "") -> List[dict]:
        return self.entity("BillPayment", where)

    def payments(self, where: str = "") -> List[dict]:
        return self.entity("Payment", where)

    def purchases(self, where: str = "") -> List[dict]:
        return self.entity("Purchase", where)

    def purchase_orders(self, where: str = "") -> List[dict]:
        return self.entity("PurchaseOrder", where)

    # ── derived indexes ──
    def project_customer_map(self) -> Dict[str, dict]:
        """{'MFD177': {id, name, …}} — strict project-# match (shared/qbo_api)."""
        if "proj_map" not in self._derived:
            access, cid = self.credentials()
            self._derived["proj_map"] = qbo_api.build_project_customer_map(access, cid)
            self._pulls += 1
        else:
            self._hits += 1
        return self._derived["proj_map"]

    def invoices_for_customer(self, customer_id: str) -> List[dict]:
        """All invoices for ONE customer/project, cached per customer id.

        Deliberately per-customer, NOT a bulk pull: the draw checks need each
        project's full invoice history (cumulative billed), and this company has
        ~33k invoices all-time — downloading the whole table costs far more than
        a handful of narrow queries (measured 2026-07-28: bulk pull turned a
        ~6-minute run into 1h40m). The cache makes repeat asks free, which is
        what the MFD and CP checks need.
        """
        cid_key = f"inv_cust:{customer_id}"
        if cid_key in self._derived:
            self._hits += 1
            return self._derived[cid_key]
        access, cid = self.credentials()
        rows = qbo_api.fetch_customer_invoices(access, cid, str(customer_id))
        self._derived[cid_key] = rows
        self._pulls += 1
        return rows

    def open_invoices(self) -> List[dict]:
        """Open invoices only (Balance > 0) — a small, cheap set, unlike the
        full invoice table."""
        return self.entity("Invoice", "Balance > '0'")

    # ── diagnostics ──
    def stats(self) -> dict:
        return {"pulls": self._pulls, "cache_hits": self._hits,
                "entities": {f"{n}{'|' + w if w else ''}": len(rows)
                             for (n, w), rows in self._entities.items()}}

    def summary(self) -> str:
        s = self.stats()
        return (f"{s['pulls']} QBO pull(s), {s['cache_hits']} served from cache"
                + (f" — {', '.join(f'{k}:{v}' for k, v in s['entities'].items())}"
                   if s["entities"] else ""))
