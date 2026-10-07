"""Process guides show people's names only at serve time (ledger/registry_view, the user 2026-10-07).

The guide on disk carries roles (data-role slots); the ledger fills them from the gitignored roster.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import registry_view as rv  # noqa: E402

ROSTER = """| Name | Also called | Handle | Division |
| ---- | ----------- | ------ | -------- |
|      |             | `the user` | all |
| Pat Doe | PD | `the payroll manager` | all |
| A & B | | `RP Operations Assistant` | RP |
| One | | `the super on the job` | CP |
| Two | | `the super on the job` | RP |
"""


def test_roster_names_skips_blank_and_shared_handles(tmp_path):
    p = tmp_path / "ROSTER.md"
    p.write_text(ROSTER)
    names = rv.roster_names(p)
    assert names == {"the payroll manager": "Pat Doe", "rp operations assistant": "A & B"}


def test_fill_names_escapes_and_leaves_unknown_blank():
    html = ('<span class="ln" data-role="the payroll manager"></span>'
            '<span class="ln" data-role="RP Operations Assistant"></span>'
            '<span class="ln" data-role="the super on the job"></span>')
    out = rv.fill_names(html, {"the payroll manager": "Pat Doe", "rp operations assistant": "A & B"})
    assert '>Pat Doe</span>' in out and '>A &amp; B</span>' in out
    assert 'data-role="the super on the job"></span>' in out
