#!/usr/bin/env python3
"""
python-env/lock.py - build requirements.lock: EVERY package (dependencies included), exact version, with
every sha256 PyPI publishes for it, and a platform marker where the Mac and Linux CI differ.

requirements.txt stays the human list (top-level pins). The lock is generated from it and is what gets
installed - `pip install --require-hashes -r requirements.lock` - so a tampered or swapped package on PyPI
(or in a mirror / cache) fails the install instead of running (security review 09/29/2026).

    python3 python-env/lock.py          regenerate after editing requirements.txt
    python3 python-env/lock.py --check  fail if the lock is missing or older than requirements.txt

How: pip resolves the full set for THIS Mac (`pip install --dry-run --report`, nothing installed); the
dependency tree is then walked again under Linux's markers (pip judges markers by the machine it runs on, so a
Linux-only dependency such as keyring's SecretStorage is invisible to it here) and any Linux-only package is
taken from PyPI at the highest version that satisfies the tree. Hashes come from PyPI's JSON API.
"""
from __future__ import annotations

import hashlib
import json
import re
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

from pip._vendor.packaging.requirements import Requirement
from pip._vendor.packaging.specifiers import SpecifierSet
from pip._vendor.packaging.utils import canonicalize_name
from pip._vendor.packaging.version import Version

HERE = Path(__file__).resolve().parent
REQS = HERE / "requirements.txt"
LOCK = HERE / "requirements.lock"
PYVER = (HERE / "PYTHON_VERSION").read_text().strip()

LINUX = {"sys_platform": "linux", "platform_system": "Linux", "os_name": "posix",
         "platform_machine": "x86_64", "implementation_name": "cpython",
         "platform_python_implementation": "CPython", "python_version": PYVER,
         "python_full_version": PYVER + ".0", "extra": ""}
TOOLS = {"pip"}                      # pinned + hashed too: setup.sh installs pip from the lock


def reqs_digest() -> str:
    return hashlib.sha256(REQS.read_bytes() + (HERE / "PYTHON_VERSION").read_bytes()).hexdigest()


def pypi(name: str, version: str = "") -> dict:
    url = f"https://pypi.org/pypi/{name}/{version + '/' if version else ''}json"
    with urllib.request.urlopen(url, timeout=60) as r:
        return json.load(r)


def mac_resolve() -> dict:
    """{canonical name: (version, requires_dist)} for this Mac - pip's own resolver, nothing installed."""
    with tempfile.TemporaryDirectory() as t:
        rep = Path(t) / "report.json"
        # Dependencies nobody pinned keep the version this environment already runs (proven, not whatever
        # PyPI published today - a brand-new release is exactly what a lock is for); a pin change that needs a
        # newer one fails here, and then it gets pinned in requirements.txt on purpose.
        cons = Path(t) / "constraints.txt"
        cons.write_text("\n".join(installed_unpinned()) + "\n")
        subprocess.run([sys.executable, "-m", "pip", "install", "--dry-run", "--ignore-installed", "--quiet",
                        "--disable-pip-version-check", "--report", str(rep), "-r", str(REQS), "-c", str(cons),
                        f"pip=={pip_version()}"], check=True)
        report = json.loads(rep.read_text())
    out = {}
    for item in report["install"]:
        md = item["metadata"]
        out[canonicalize_name(md["name"])] = (md["version"], md.get("requires_dist") or [], md["name"])
    return out


def installed_unpinned() -> list:
    """name==version for every package in this environment that requirements.txt does not pin."""
    from importlib.metadata import distributions
    pinned = {canonicalize_name(r.name) for r in top_level()}
    return sorted({f"{d.metadata['Name']}=={d.version}" for d in distributions()
                   if canonicalize_name(d.metadata["Name"]) not in pinned})


def pip_version() -> str:
    import pip
    return pip.__version__


def top_level() -> list:
    rows = []
    for line in REQS.read_text().splitlines():
        line = line.split("#", 1)[0].strip()
        if line:
            rows.append(Requirement(line))
    rows.append(Requirement(f"pip=={pip_version()}"))
    return rows


def linux_walk(mac: dict) -> dict:
    """{canonical name: version} needed on Linux: the tree re-walked under Linux markers; a package the Mac
    resolve does not have is taken from PyPI (highest version satisfying every specifier seen)."""
    need, specs, seen = {}, {}, set()
    queue = [r for r in top_level() if r.marker is None or r.marker.evaluate(LINUX)]
    extra_meta = {}
    while queue:
        r = queue.pop()
        n = canonicalize_name(r.name)
        specs[n] = specs.get(n, SpecifierSet()) & r.specifier
        if n in seen:
            continue
        seen.add(n)
        if n in mac:
            version, requires, _ = mac[n]
        else:
            info = pypi(r.name)
            ok = [v for v in info["releases"] if info["releases"][v] and not Version(v).is_prerelease
                  and Version(v) in specs[n]]
            if not ok:
                raise SystemExit(f"no Linux release of {r.name} satisfies {specs[n]}")
            version = str(max(ok, key=Version))
            requires = pypi(r.name, version)["info"].get("requires_dist") or []
            extra_meta[n] = r.name
        need[n] = version
        for d in requires:
            dr = Requirement(d)
            env = dict(LINUX, extra="")
            if dr.marker is None or dr.marker.evaluate(env):
                queue.append(dr)
    for n, v in need.items():
        if Version(v) not in specs.get(n, SpecifierSet()):
            raise SystemExit(f"{n}=={v} does not satisfy {specs[n]} on Linux - pin it in requirements.txt")
    return need, extra_meta


def hashes(name: str, version: str) -> list:
    files = pypi(name, version)["urls"]
    return sorted({f["digests"]["sha256"] for f in files})


def build() -> None:
    mac = mac_resolve()
    linux, linux_names = linux_walk(mac)
    rows = []
    names = sorted(set(mac) | set(linux))
    for n in names:
        mv = mac[n][0] if n in mac else None
        lv = linux.get(n)
        display = mac[n][2] if n in mac else linux_names.get(n, n)
        if mv and lv and mv != lv:
            raise SystemExit(f"{display}: Mac resolves {mv}, Linux {lv} - pin it in requirements.txt")
        version = mv or lv
        marker = "" if (mv and lv) else ('sys_platform == "darwin"' if mv else 'sys_platform == "linux"')
        rows.append((display, version, marker, hashes(display, version)))
    out = [
        "# python-env/requirements.lock - GENERATED by python-env/lock.py from requirements.txt. Do not edit.",
        "# Every package, dependencies included, exact version + every sha256 PyPI publishes for it.",
        "# Installed with: pip install --require-hashes -r python-env/requirements.lock",
        f"# requirements-sha256: {reqs_digest()}",
        "",
    ]
    for display, version, marker, hs in rows:
        head = f"{display}=={version}" + (f" ; {marker}" if marker else "")
        out.append(head + " \\")
        out.extend(f"    --hash=sha256:{h}" + (" \\" if i < len(hs) - 1 else "") for i, h in enumerate(hs))
    LOCK.write_text("\n".join(out) + "\n")
    print(f"wrote {LOCK.name}: {len(rows)} packages "
          f"({sum(1 for r in rows if not r[2])} both, {sum(1 for r in rows if 'darwin' in r[2])} Mac-only, "
          f"{sum(1 for r in rows if 'linux' in r[2])} Linux-only)")


def check() -> int:
    if not LOCK.exists():
        print("   python-env/requirements.lock is missing - run: python3 python-env/lock.py")
        return 1
    m = re.search(r"^# requirements-sha256: ([0-9a-f]{64})$", LOCK.read_text(), re.M)
    if not m or m.group(1) != reqs_digest():
        print("   requirements.txt / PYTHON_VERSION changed since requirements.lock was built - run: "
              "python3 python-env/lock.py")
        return 1
    return 0


if __name__ == "__main__":
    if "--check" in sys.argv:
        sys.exit(check())
    build()
