# ─────────────────────────────────────────────────────────────────────────────
# Dockerfile - the office server (Synology), 2026-09-30.
#
# ONE container, ONE process (docker/scheduler.py): the QBO mirror every 3 min,
# the AP/AR job every 15 min, a nightly count check. It carries READ code only -
# the allow-list below; tests/test_office_server_image.py fails the build if a
# QuickBooks writer (pay bills, re-apply, recode, bulk close, reclass) is ever
# copied in. Build context = the repo root; .dockerignore lets in only the list.
#
# Python and packages are the SAME pins as the Mac and CI (python-env/): the base
# image is pinned by digest, the packages by version + sha256 (--require-hashes).
# ─────────────────────────────────────────────────────────────────────────────
FROM python:3.14-slim@sha256:51dafde81dbdb6ebde285137a295cf18a47ca95234fe388a343719cb97305b3d

LABEL org.opencontainers.image.title="office-server" \
      org.opencontainers.image.description="QBO mirror + AP/AR sync for the office server (read code only)"

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    TZ=America/Chicago

RUN apt-get update \
 && apt-get install -y --no-install-recommends tzdata ca-certificates \
 && rm -rf /var/lib/apt/lists/*

# a plain user; IT may set these to the Synology automation user's ids at build time
ARG APP_UID=1000
ARG APP_GID=1000
RUN groupadd --gid ${APP_GID} app \
 && useradd --uid ${APP_UID} --gid ${APP_GID} --shell /usr/sbin/nologin --create-home app

WORKDIR /app
COPY python-env/requirements.lock /app/python-env/requirements.lock
RUN pip install --require-hashes -r /app/python-env/requirements.lock

# ── THE ALLOW-LIST (read code only) ─────────────────────────────────────────
COPY shared/                  /app/shared/
COPY bill-tracker/            /app/bill-tracker/
COPY invoice-sync/            /app/invoice-sync/
COPY ledger/refresh_mirror.py /app/ledger/refresh_mirror.py
COPY docker/scheduler.py      /app/docker/scheduler.py
# ────────────────────────────────────────────────────────────────────────────

RUN mkdir -p /data && chown -R app:app /data
USER app

ENTRYPOINT ["python", "/app/docker/scheduler.py"]
