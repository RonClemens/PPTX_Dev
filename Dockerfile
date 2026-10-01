# syntax=docker/dockerfile:1

# --- Stage 1: build the React frontend into static assets -----------------
FROM node:22-slim AS frontend-build
WORKDIR /app/frontend
COPY frontend/package.json frontend/package-lock.json ./
# Optional build-time CA bundle for networks that inspect/re-sign TLS (set
# PPTX_DEV_CA_BUNDLE, see README). Mounted as a secret so it is never stored
# in an image layer; ignored when empty/absent.
RUN --mount=type=secret,id=ca,required=false \
    if [ -s /run/secrets/ca ]; then export NODE_EXTRA_CA_CERTS=/run/secrets/ca; fi; \
    npm ci
COPY frontend/ ./
RUN npm run build

# --- Stage 2: Python backend, serving the built frontend too ---------------
FROM python:3.11-slim AS backend
WORKDIR /app/backend

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# The app never runs as root: docker-entrypoint.sh drops to this user (and,
# when started as root, first makes sure the data directory is writable by it
# -- e.g. a root-owned Render disk or bind mount).
RUN groupadd --system --gid 10001 app \
 && useradd --system --uid 10001 --gid app --no-create-home --shell /usr/sbin/nologin app

COPY backend/requirements.txt ./
RUN --mount=type=secret,id=ca,required=false \
    if [ -s /run/secrets/ca ]; then export PIP_CERT=/run/secrets/ca; fi; \
    pip install -r requirements.txt

COPY backend/app ./app
COPY --from=frontend-build /app/frontend/dist /app/frontend_dist
COPY docker-entrypoint.sh /usr/local/bin/docker-entrypoint.sh

ENV PPTX_DEV_FRONTEND_DIST=/app/frontend_dist \
    PPTX_DEV_DATA_DIR=/data

# Everything the app persists lives here; mount a volume on it.
RUN mkdir -p /data && chown app:app /data && chmod 0755 /usr/local/bin/docker-entrypoint.sh
VOLUME ["/data"]

# Render (and most container platforms) inject PORT at runtime; default to
# 8000 for local `docker run`.
ENV PORT=8000
EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
  CMD ["python", "-c", "import os,urllib.request; urllib.request.urlopen('http://127.0.0.1:%s/api/health' % os.environ.get('PORT','8000'), timeout=4)"]

# One worker, deliberately: the API key/base URL/model set from the in-app
# Settings modal are held in that single process's memory.
ENTRYPOINT ["docker-entrypoint.sh"]
CMD ["sh", "-c", "exec uvicorn app.main:app --host 0.0.0.0 --port ${PORT}"]
