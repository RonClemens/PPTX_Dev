"""AI connection settings -- STATELESS on the server.

The user's API key / token, base URL, auth mode, extra headers and model are
kept in the user's own browser and sent with each AI request (see
app/ai/credentials.py); this server stores none of them. So there is
deliberately no "save settings" endpoint. What remains:

* GET  /api/settings/anthropic-key       -- what the *server's environment*
  supplies as fallback defaults (nothing, unless an operator set env vars).
* POST /api/settings/anthropic-key/test  -- one minimal real request using
  whatever credentials THIS request carries, so "Test Connection" proves the
  exact configuration the browser will use for AI features.
"""
from fastapi import APIRouter, HTTPException

from .. import config
from ..ai import AdjudicationError, test_connection

router = APIRouter(prefix="/api/settings", tags=["settings"])


@router.get("/anthropic-key")
def get_server_defaults():
    """Environment-supplied defaults only; never any user-entered value (the
    server doesn't have them). The key itself is never returned."""
    return {
        "storesCredentials": False,
        "serverKeyConfigured": bool(config.ANTHROPIC_API_KEY),
        "baseUrl": config.ANTHROPIC_BASE_URL or None,
        "authMode": config.AUTH_MODE,
        "extraHeaders": sorted(config.EXTRA_HEADERS),  # names only
        "model": config.AI_MODEL,
        "defaultModel": config.DEFAULT_AI_MODEL,
    }


@router.post("/anthropic-key/test")
def test_anthropic_connection():
    """Make one minimal real request (max_tokens=1) with the credentials
    this request carries -- the exact same client construction every other AI
    feature uses -- so a green result means "Ask AI" will actually work."""
    try:
        result = test_connection()
    except AdjudicationError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc
    return {"ok": True, **result}
