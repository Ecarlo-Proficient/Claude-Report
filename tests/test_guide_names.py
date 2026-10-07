"""Process guides show people's names from names.js beside them (ledger/registry_view, the user 2026-10-07).

The guide on disk carries roles only; the ledger inlines the gitignored names.js when it serves the page.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import registry_view as rv  # noqa: E402

PAGE = '<head><script src="names.js"></script></head><b data-role="payroll manager">Payroll Manager</b>'


def test_inline_names_puts_the_file_in_the_page(tmp_path):
    (tmp_path / "names.js").write_text('window.GUIDE_NAMES = {"payroll manager": "Pat </script> Doe"};')
    out = rv.inline_names(PAGE, tmp_path)
    assert 'src="names.js"' not in out
    assert "Pat <\\/script> Doe" in out and out.count("</script>") == 1


def test_no_names_file_leaves_the_page_alone(tmp_path):
    assert rv.inline_names(PAGE, tmp_path) == PAGE
