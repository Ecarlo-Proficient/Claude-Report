#!/usr/bin/env python3
"""
refresh_mirror.py - keep the raw QBO mirror current (shared/qbo_mirror).

  python3 ledger/refresh_mirror.py                 refresh: the change feed since the last stamp
  python3 ledger/refresh_mirror.py --seed          full pull (first time, or a rebuild; ~20 min)
  python3 ledger/refresh_mirror.py --reconcile     COUNT(*) per entity vs the mirror, sweep drift
  python3 ledger/refresh_mirror.py --status        counts, stamps, size - no QBO call
  --entities Bill,Invoice                          limit --seed to some entities

Read-only on QBO. Writes only the mirror file (outside the repo). Never prints
the realm. Step 0 of sync-all once the tools read the mirror.
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import qbo_api                 # noqa: E402
from shared import qbo_mirror as mirror    # noqa: E402


def _status() -> int:
    s = mirror.status()
    print(f"mirror: {s['db']} ({s['size_mb']} MB)")
    con = mirror.connect()
    try:
        enc, plain = mirror.encryption_status(con)
    finally:
        con.close()
    print(f"  at rest: {enc:,} rows encrypted (AES-256-GCM, key in the Keychain library)"
          + (f" · {plain:,} PLAIN - run --encrypt" if plain else ""))
    print(f"  seeded {s['seeded_at'] or '-'} · last refresh {s['last_refresh'] or '-'}")
    tot = 0
    for e, r in s["entities"].items():
        tot += r["alive"]
        print(f"  {e:14s} {r['alive']:>8,} alive  {r['deleted']:>6,} deleted   newest {r['newest'] or '-'}")
    print(f"  {'TOTAL':14s} {tot:>8,}")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--seed", action="store_true", help="full pull of every entity")
    g.add_argument("--reconcile", action="store_true", help="count check per entity, sweep drift")
    g.add_argument("--status", action="store_true", help="what the mirror holds (no QBO call)")
    g.add_argument("--encrypt", action="store_true",
                   help="encrypt any plain rows with the Keychain MIRROR_KEY (one-time migration; no QBO call)")
    ap.add_argument("--entities", default="", help="comma list, --seed only")
    a = ap.parse_args(argv)
    if a.status:
        return _status()
    if a.encrypt:
        con = mirror.connect()
        try:
            n = mirror.encrypt_existing(con)
            enc, plain = mirror.encryption_status(con)
        finally:
            con.close()
        print(f"done: {n:,} rows encrypted this run · {enc:,} encrypted · {plain:,} plain")
        return 0 if plain == 0 else 1
    access = qbo_api.refresh_access()
    _, cid = qbo_api.load_credentials()
    con = mirror.connect()
    t0 = time.time()
    try:
        if a.seed:
            ents = [e.strip() for e in a.entities.split(",") if e.strip()] or None
            bad = [e for e in (ents or []) if e not in mirror.ENTITIES]
            if bad:
                print(f"unknown entities: {bad}; choose from {', '.join(mirror.ENTITIES)}")
                return 2
            print(f"SEED {'all entities' if not ents else ', '.join(ents)}")
            res = mirror.seed(con, access, cid, entities=ents)
        elif a.reconcile:
            print("RECONCILE")
            res = mirror.reconcile(con, access, cid)
        else:
            print("REFRESH")
            res = mirror.refresh(con, access, cid)
    finally:
        con.close()
    print(f"done: {res['mode']} · {res.get('upserted', res.get('fetched', 0)):,} written · "
          f"{res.get('deleted', 0):,} deleted · {time.time() - t0:.0f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
