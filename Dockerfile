# syntax=docker/dockerfile:1

# --- Stage 1: build the React frontend into static assets -----------------
FROM node:22-slim AS frontend-build
WORKDIR /app/frontend
COPY frontend/package.json frontend/package-lock.json ./
RUN npm ci
COPY frontend/ ./
RUN npm run build

# --- Stage 2: Python backend, serving the built frontend too ---------------
FROM python:3.11-slim AS backend
WORKDIR /app/backend

COPY backend/requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY backend/app ./app

COPY --from=frontend-build /app/frontend/dist /app/frontend_dist

ENV PPTX_DEV_FRONTEND_DIST=/app/frontend_dist \
    PPTX_DEV_DATA_DIR=/data \
    PYTHONUNBUFFERED=1

# Render (and most container platforms) inject PORT at runtime; default to
# 8000 for local `docker run`.
ENV PORT=8000
EXPOSE 8000

CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT}"]
