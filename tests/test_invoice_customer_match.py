"""invoice-sync customer matcher: one shared word is not a match (audit 2026-09-25).

Pure logic - a hand-built customer cache, no Notion, no QBO. The cases are the
real QBO -> Notion pairs the audit found, company names only.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))                    # repo root -> shared
sys.path.insert(0, str(ROOT / "invoice-sync"))   # the module under test

import invoice_sync as s  # noqa: E402


def _cache(*names):
    return [s.CustomerCacheEntry(page_id=f"p{i}", raw_name=n, keywords=s._keywords(n),
                                 compressed=s._compressed_form(n))
            for i, n in enumerate(names)]


def _match(qbo_name, *notion_names):
    m = s._lookup_customer_id(qbo_name, _cache(*notion_names))
    return m.matched_name if m else None


# -- wrong links the old matcher made: now left unlinked (the sync warns) --
def test_one_shared_word_is_not_a_match():
    assert _match("NATIONAL HOME CORPORATION", "GGC NATIONAL CONTRACTORS") is None
    assert _match("Tri-Star Construction, Inc.", "Tri-C Construction") is None
    assert _match("North Forest Development- DFW, LLC", "DFW CUSTOM HOMES") is None
    assert _match("URBAN ESSENTIALS, LLC", "CLASSIC URBAN HOMES") is None
    assert _match("KNIGHT PRIVATE FUND, LLC", "KNIGHT CONTRACTING GROUP") is None
    assert _match("MCR DEL ROY, LLC", "MCR BUILDERS, LLC") is None


def test_two_shared_words_with_a_stray_word_is_not_a_match():
    assert _match("NAVARRO COUNTY HABITAT", "HABITAT FOR HUMANITY OF COLLIN COUNTY") is None


# -- right links the old matcher made: kept --
def test_typo_and_plural_drift_still_match():
    assert _match("FOURTEEN CONSTRUCTION", "FOURTEN CONSTRUCTION, LLC") == "FOURTEN CONSTRUCTION, LLC"
    assert _match("Peterson Constuction, Inc", "PETERSON CONSTRUCTION") == "PETERSON CONSTRUCTION"


def test_all_qbo_words_found_with_two_shared_still_matches():
    assert _match("Perry Guest Construction, LLC", "PERRY GUEST COMPANIES") == "PERRY GUEST COMPANIES"
    assert _match("HGR GENERAL CONTRACTORS", "HGR GENERAL CONTRACTORS, LP") == "HGR GENERAL CONTRACTORS, LP"


def test_notion_name_leading_the_qbo_name_still_matches():
    assert _match("DHI Communities Construction of Texas LLC", "DHI") == "DHI"
    assert _match("Embrey Builders LLC-Champions Way DFW LP",
                  "Embrey Builders LLC") == "Embrey Builders LLC"


def test_all_stopword_names_unchanged():
    assert (_match("Development & Construction Services LLC", "DEVELOPMENT & CONSTRUCTION SERVICE")
            == "DEVELOPMENT & CONSTRUCTION SERVICE")


def test_right_customer_wins_over_a_one_word_decoy():
    assert _match("GGC National Contractors", "NATIONAL HOME CORPORATION",
                  "GGC NATIONAL CONTRACTORS") == "GGC NATIONAL CONTRACTORS"
