#!/bin/sh
# Start the app as the unprivileged "app" user.
#
# - Started as root (plain `docker run`, Render): make sure the data
#   directory is writable by "app" -- only recursing when its top level isn't
#   already owned by it, so normal restarts are instant -- then drop
#   privileges before exec'ing the server.
# - Started as a non-root user (docker-compose.yml's `user:`): nothing to
#   fix or drop; the volume must already be writable by that user.
set -e

DATA_DIR="${PPTX_DEV_DATA_DIR:-/data}"

if [ "$(id -u)" = "0" ]; then
  mkdir -p "$DATA_DIR"
  app_uid="$(id -u app)"
  if [ "$(stat -c %u "$DATA_DIR")" != "$app_uid" ]; then
    # Best effort: some bind-mount filesystems (Docker Desktop on Windows/macOS)
    # don't support chown; they're typically already writable anyway.
    chown -R app:app "$DATA_DIR" 2>/dev/null || echo "warning: could not chown $DATA_DIR" >&2
  fi
  exec setpriv --reuid=app --regid=app --init-groups "$@"
fi

exec "$@"
