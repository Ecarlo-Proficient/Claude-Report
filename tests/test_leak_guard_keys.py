"""The key/token half of .github/leak_guard.sh (security review 09/29/2026): each kind of secret committed
to a file is caught, and the guard never prints the value. Fake values are built at runtime so this file
itself never matches."""
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
GUARD = ROOT / ".github" / "leak_guard.sh"

FAKES = {
    "notion": "ntn_" + "a1" * 22,
    "teams": "https://x.webhook.office.com/webhookb2/" + "b2" * 15,
    "flow": "https://prod-01.westus.logic.azure.com/workflows/x/triggers/manual/paths/invoke?sig=" + "c3" * 15,
    "pem": "-----BEGIN " + "RSA PRIVATE KEY-----",
    "github": "ghp_" + "d4" * 18,
    "aws": "AKIA" + "E5F6" * 4,
    "intuit": "AB11" + "7" * 10 + "f7" * 14,
    "named": "QBO_CLIENT_SECRET" + "=" + "g8" * 15,
}


def run_guard(tmp_path, text):
    repo = tmp_path / "repo"
    repo.mkdir()
    subprocess.run(["git", "init", "-q"], cwd=repo, check=True)
    (repo / "config.py").write_text(text + "\n")
    subprocess.run(["git", "add", "config.py"], cwd=repo, check=True)
    return subprocess.run(["bash", str(GUARD)], cwd=repo, capture_output=True, text=True)


@pytest.mark.skipif(shutil.which("git") is None, reason="git missing")
@pytest.mark.parametrize("kind", sorted(FAKES))
def test_each_kind_is_caught_without_printing_it(tmp_path, kind):
    r = run_guard(tmp_path, f'VALUE = "{FAKES[kind]}"')
    assert r.returncode == 1, (kind, r.stdout)
    assert "key or token" in r.stdout
    assert FAKES[kind] not in r.stdout + r.stderr


def test_ordinary_code_passes(tmp_path):
    r = run_guard(tmp_path, 'token = kc.get_secret("NOTION_SECRET")\nQBO_CLIENT_SECRET = ""')
    assert r.returncode == 0, r.stdout
