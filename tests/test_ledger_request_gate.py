"""The ledger's request gate + open-path guards (security review 09/29/2026): another website must not be able
to drive the local server (cross-site POST, DNS rebinding, <img> GETs), and the Finder helpers must never
launch an app. Pure functions plus one real server round-trip on a free port - no ledger data touched."""
import http.client
import sys
import threading
from http.server import ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "ledger"))

import dashboard as d   # noqa: E402

PORT = 8787
OK_HOST = {"Host": f"127.0.0.1:{PORT}"}


def gate(method, path, **h):
    return d.request_block_reason(method, path, {**OK_HOST, **h}, PORT)


def test_ledger_page_and_scripts_pass():
    assert gate("GET", "/") is None                                              # typed / bookmarked
    assert gate("GET", "/api/data", **{"Sec-Fetch-Site": "same-origin"}) is None
    assert gate("POST", "/api/sync", Origin=f"http://127.0.0.1:{PORT}", **{"Sec-Fetch-Site": "same-origin"}) is None
    assert gate("POST", "/api/pay-run/clear") is None                             # curl / a local script: no browser headers
    assert d.request_block_reason("GET", "/api/data", {"Host": f"localhost:{PORT}"}, PORT) is None


def test_dns_rebinding_refused():
    assert d.request_block_reason("GET", "/api/data", {"Host": f"evil.example:{PORT}"}, PORT) == "host"
    assert d.request_block_reason("POST", "/api/checkdrift/reapply", {"Host": "evil.example"}, PORT) == "host"
    assert d.request_block_reason("GET", "/", {}, PORT) == "host"


def test_cross_site_refused():
    assert gate("POST", "/api/sync", Origin="https://evil.example") == "origin"
    assert gate("POST", "/api/sync", Origin="null") == "origin"                 # sandboxed iframe / file:
    assert gate("POST", "/api/sync", **{"Sec-Fetch-Site": "cross-site"}) == "cross-site"
    assert gate("GET", "/api/checkdrift/reapply", **{"Sec-Fetch-Site": "cross-site"}) == "cross-site"
    assert gate("GET", "/api/data", Origin="https://evil.example") == "origin"
    assert gate("GET", "/", **{"Sec-Fetch-Site": "cross-site"}) is None          # a link to the page itself is harmless


def test_launchables_refused(tmp_path):
    assert d._launchable("/Applications/Terminal.app")
    assert d._launchable("/Applications/Terminal.app/Contents/MacOS/Terminal")
    assert d._launchable("/Users/x/Downloads/run.command")
    prog = tmp_path / "tool"
    prog.write_text("#!/bin/sh\n")
    prog.chmod(0o755)
    assert d._launchable(prog)
    assert not d._launchable("/Volumes/Common/Jobs/RP7608_5011 BOURQUIN/Profit and Loss/Project_PnL_RP7608.xlsx")
    assert not d._launchable(tmp_path)
    assert d._os_open("/Applications/Calculator.app") == "refused: not a folder or document"


def test_reveal_root_check(tmp_path):
    root = tmp_path / "Common"
    (root / "Job").mkdir(parents=True)
    (tmp_path / "CommonX").mkdir()
    assert d._under(root / "Job", [root])
    assert d._under(root, [root])
    assert not d._under(root / ".." / ".." / "Applications", [root])            # the old startswith() let this through
    assert not d._under(tmp_path / "CommonX", [root])
    link = root / "Job" / "escape"
    link.symlink_to(tmp_path / "CommonX")
    assert not d._under(link, [root])


def test_realm_scrubbed_from_errors():
    assert d._scrub("/v3/company/1234567890/query -> 401") == "/v3/company/<id>/query -> 401"


def test_real_server_refuses_foreign_host():
    srv = ThreadingHTTPServer(("127.0.0.1", 0), d.Handler)
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    port = srv.server_address[1]
    try:
        c = http.client.HTTPConnection("127.0.0.1", port, timeout=5)
        c.request("POST", "/api/pay-run/clear", headers={"Host": f"evil.example:{port}"})
        r = c.getresponse()
        assert r.status == 403 and b"did not come from the ledger page" in r.read()
        c.request("POST", "/api/pay-run/clear", headers={"Origin": "https://evil.example"})
        r = c.getresponse()
        assert r.status == 403
        r.read()
    finally:
        srv.shutdown()
        srv.server_close()
