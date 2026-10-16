"""ledger/draw_periods.py - the period read, the memo rewrite, the overlap rule, and the save on the ledger (the only
write there is: the owner 2026-10-11, "save but only on the ledger"; QuickBooks is never touched)."""
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import draw_periods as D   # noqa: E402

MEMO = "MFD325 - Mesquite - Briarwood - August Draw 2026\n(Period:07/02/2026 - 08/01/2026)"


def test_period_read_and_rewritten_in_place():
    assert D.period_of(MEMO) == ("2026-07-02", "2026-08-01")
    assert D.period_of("no window here") == (None, None)
    out = D.memo_with(MEMO, "2026-07-05", "2026-08-04")
    assert out == "MFD325 - Mesquite - Briarwood - August Draw 2026\n(Period:07/05/2026 - 08/04/2026)"
    assert D.memo_with("Draw #4", "2026-01-01", "2026-01-31") == "Draw #4\n(Period:01/01/2026 - 01/31/2026)"
    assert D.memo_with("", "2026-01-01", "2026-01-31") == "(Period:01/01/2026 - 01/31/2026)"


def test_check_refuses_overlap_and_backwards_windows():
    ok = [{"id": "1", "label": "July", "start": "2026-06-02", "end": "2026-07-01"},
          {"id": "2", "label": "August", "start": "2026-07-02", "end": "2026-08-01"},
          {"id": "3", "label": "August", "start": "2026-07-02", "end": "2026-08-01"}]   # a second invoice on the same draw
    assert D.check(ok) == ""
    bad = [dict(ok[0]), dict(ok[1], start="2026-07-01")]
    assert "overlaps" in D.check(bad)
    assert "before the start" in D.check([{"id": "1", "label": "x", "start": "2026-07-02", "end": "2026-07-01"}])
    assert "longer than" in D.check([{"id": "1", "label": "x", "start": "2026-01-01", "end": "2026-12-31"}])
    assert "needed" in D.check([{"id": "1", "label": "x", "start": "", "end": "2026-07-01"}])


def test_save_keeps_only_windows_off_the_invoice_and_refuses_overlap():
    con = sqlite3.connect(":memory:")
    inv = [{"id": "1", "label": "July", "start": "2026-06-02", "end": "2026-07-01", "inv_start": "2026-06-02", "inv_end": "2026-07-01"},
           {"id": "2", "label": "August", "start": "2026-07-05", "end": "2026-08-04", "inv_start": "2026-07-02", "inv_end": "2026-08-01"},
           {"id": "3", "label": "August", "start": "2026-07-05", "end": "2026-08-04", "inv_start": "2026-07-02", "inv_end": "2026-08-01"}]
    r = D.save(con, "mfd325", inv)
    assert r["ok"] and r["saved"] == 2
    got = D.saved(con, "MFD325")
    assert set(got) == {"2", "3"} and got["2"][:2] == ("2026-07-05", "2026-08-04")
    assert "overlaps" in D.save(con, "MFD325", [dict(inv[0], end="2026-07-06"), inv[1]])["error"]
    assert set(D.saved(con, "MFD325")) == {"2", "3"}                       # a refused save changes nothing
    back = [dict(p, start=p["inv_start"], end=p["inv_end"]) for p in inv]
    r2 = D.save(con, "MFD325", back)
    assert r2["ok"] and r2["saved"] == 0                                    # back to as invoiced = nothing kept
    assert D.saved(con, "MFD325") == {}
    assert D.saved(sqlite3.connect(":memory:"), "MFD325") == {}            # no table yet: nothing held
