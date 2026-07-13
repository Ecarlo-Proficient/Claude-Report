"""
version.py — release identity for the invoice sync.

Two release lines:
  • Docker — the v1 "true release". Version "1.0.0", injected via the APP_VERSION
    env var baked into the image (see docker/Dockerfile). This is the production
    target once testing is signed off.
  • Mac-only lineage — "mvN" (mac version N): the manual `sync-ar` / visual-viewer
    path that predates the container. Bump MAC_VERSION on Mac-side changes.

Resolution: APP_VERSION (set by Docker) wins; otherwise the Mac version string.
`runtime_label()` is what you show in logs / alerts so it's always obvious which
instance is talking — important while Mac and Docker run side by side in testing.
"""
from __future__ import annotations

import os
from pathlib import Path

# Mac-only lineage version — bump this on Mac-side-only changes (e.g. "mv2").
MAC_VERSION = "mv1"

# The true v1 release ships in Docker; the image sets APP_VERSION=1.0.0.
DOCKER_DEFAULT = "1.0.0"


def in_docker() -> bool:
    """True when running inside the container (image sets APP_VERSION; Docker
    also creates /.dockerenv)."""
    return os.environ.get("APP_VERSION") is not None or Path("/.dockerenv").exists()


def version() -> str:
    """The version string: APP_VERSION if set (Docker), else the Mac version."""
    return os.environ.get("APP_VERSION") or MAC_VERSION


def runtime_label() -> str:
    """Human label for logs/alerts, e.g. 'v1.0.0 (docker)' or 'mv1 (mac)'."""
    return f"v{version()} (docker)" if in_docker() else f"{version()} (mac)"
