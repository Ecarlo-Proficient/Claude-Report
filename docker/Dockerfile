# ─────────────────────────────────────────────────────────────────────────────
# Dockerfile — invoice-sync (Phase 1)
#
# Phase 1 scope: QBO → Notion invoice sync only. Excel/OneDrive export is
# disabled inside the container via SKIP_EXCEL_EXPORT=1 (the Mac side still
# owns Excel until Phase 2 adds rclone-based OneDrive sync).
#
# Build context = project root (one level up from this Dockerfile). The build
# command lives in docker/README.md; do not run `docker build` from here.
# ─────────────────────────────────────────────────────────────────────────────
FROM python:3.11-slim AS base

# Container metadata. v1.1.0 = the 2026-07 repo restructure (code identical,
# paths renamed: automation-worker/ → invoice-sync/, vault → shared/). The
# Mac-only runtime is the separate "mvN" lineage (see invoice-sync/version.py).
LABEL org.opencontainers.image.title="invoice-sync"
LABEL org.opencontainers.image.description="QBO → Notion invoice sync worker"
LABEL org.opencontainers.image.source="invoice-sync/"
LABEL org.opencontainers.image.version="1.1.0"

# Avoid Python bytecode and unbuffered stdout (logs flush immediately to docker logs)
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    TZ=America/Chicago

# OS-level deps: tzdata for log timestamps, ca-certificates for HTTPS (QBO, Notion)
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
        tzdata \
        ca-certificates \
 && rm -rf /var/lib/apt/lists/*

# Non-root user — IT may want to override UID/GID at build time to match volume perms
ARG APP_UID=1000
ARG APP_GID=1000
RUN groupadd --gid ${APP_GID} app \
 && useradd  --uid ${APP_UID} --gid ${APP_GID} --shell /bin/bash --create-home app

WORKDIR /app

# Install Python deps first (better layer caching)
COPY invoice-sync/requirements.txt /app/requirements.txt
RUN pip install -r /app/requirements.txt

# Copy the shared package (vault, paths, qbo_api) and the sync code
COPY shared/                     /app/shared/
COPY invoice-sync/               /app/invoice-sync/
COPY docker/entrypoint.sh        /app/entrypoint.sh
RUN chmod +x /app/entrypoint.sh

# Persistent state on the /data volume:
#   /data/qbo_secrets.json — rotated QBO refresh token
#   /data/state            — sync + CDC-deletion watermarks (must survive recreate)
RUN mkdir -p /data/state \
 && chown -R app:app /app /data

USER app
WORKDIR /app/invoice-sync

# Defaults — override per environment via .env file or docker-compose env vars.
# INVOICE_EXPORT_PATH points at a volume-mounted folder that Synology Cloud Sync
# (or equivalent) bridges to OneDrive. See docker/README.md "OneDrive bridge".
# Excel/OneDrive export is OFF in the container for v1 (Mac still owns the Excel
# mirror during coexistence testing). Flip SKIP_EXCEL_EXPORT to "" once the
# OneDrive bridge (Cloud Sync → /data/onedrive) is validated. INVOICE_EXPORT_PATH
# is pre-staged for that day.
ENV APP_VERSION=1.1.0 \
    QBO_SECRETS_FILE=/data/qbo_secrets.json \
    STATE_DIR=/data/state \
    SKIP_EXCEL_EXPORT=1 \
    INVOICE_EXPORT_PATH=/data/onedrive/Open_Invoices.xlsx \
    SYNC_INTERVAL_SECONDS=900

ENTRYPOINT ["/app/entrypoint.sh"]
