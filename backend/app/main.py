import datetime as dt
import os
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from . import config
from .api import documents, edits, comments, ai, settings
from .storage import store

app = FastAPI(title="PPTX Review AI Assistant")

# Process start time -- exposed via /api/health so it's possible to tell
# "the container restarted" (a fresh, earlier timestamp on every request
# right after a deploy or a spin-down/wake cycle) apart from an actual
# application bug. On a platform with an ephemeral filesystem, a restart
# also silently empties PPTX_DEV_DATA_DIR, which looks identical to data
# loss from the browser's side -- this timestamp is the way to tell them
# apart after the fact.
_PROCESS_STARTED_AT = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(documents.router)
app.include_router(edits.router)
app.include_router(comments.router)
app.include_router(ai.router)
app.include_router(ai.doc_router)
app.include_router(settings.router)


@app.get("/api/health")
def health():
    data_dir = str(config.DATA_DIR)
    return {
        "status": "ok",
        # If this jumps to a recent/current time between two requests spaced
        # only a few minutes apart, the server process restarted in between
        # (a redeploy, a crash, or -- on a platform with idle spin-down --
        # simply having been asleep) rather than anything the app did wrong.
        "serverStartedAt": _PROCESS_STARTED_AT,
        "dataDir": data_dir,
        # True only when DATA_DIR is a real mounted filesystem (e.g. Render's
        # persistent disk); False means it's just an ordinary directory in
        # the container's own (ephemeral, wiped on every restart) filesystem
        # -- the single most direct way to confirm persistent storage is
        # actually wired up, short of checking the Render dashboard itself.
        "dataDirIsMountPoint": os.path.ismount(data_dir),
        "documentCount": len(store.list_documents()),
    }


# In a single-container deployment (e.g. Render), the built frontend
# (frontend/dist) is copied into the image and served by this same process,
# so there's no separate static host and no cross-origin requests to worry
# about. In local dev, PPTX_DEV_FRONTEND_DIST is unset and the frontend runs
# separately via `npm run dev` (see vite.config.ts's dev-only /api proxy),
# so this mount is skipped entirely -- it never interferes with pytest or
# `uvicorn --reload` on a checkout with no `frontend/dist`.
# The mount is registered last: FastAPI/Starlette tries routes in the order
# they were added, so the explicit /api/* routes above always match first,
# and this catch-all only serves paths nothing else claimed.
_frontend_dist = os.environ.get("PPTX_DEV_FRONTEND_DIST", "")
if _frontend_dist and Path(_frontend_dist).is_dir():
    app.mount("/", StaticFiles(directory=_frontend_dist, html=True), name="frontend")
