"""Runtime settings -- the Anthropic API key and, optionally, an alternate
API base URL and model name.

This is a dev/testing convenience: it lets you set these (and, on a CUI
deployment, point the client at an Anthropic-API-compatible gateway in
front of AWS Bedrock instead of the public Anthropic API, whose model IDs
may not match the direct API's) from the webapp instead of restarting the
backend with env vars. All three are held in process memory only (never
written to disk, never logged, the API key never included in any API
response) and are lost on restart. In a real GovCloud deployment, these
should instead be injected as env vars (ANTHROPIC_API_KEY,
ANTHROPIC_BASE_URL, PPTX_DEV_AI_MODEL) from your container platform's
secrets manager -- they're shared org config, not something individual
users should be able to view or change via a browser modal that any
client of the API can reach. This endpoint exists purely so a
developer/operator can test the AI adjudication flow without a redeploy.
"""
from fastapi import APIRouter, HTTPException

from .. import config
from ..ai import test_connection, AdjudicationError
from ..models import SetApiKeyRequest

router = APIRouter(prefix="/api/settings", tags=["settings"])


def _masked(key: str) -> str:
    if len(key) <= 8:
        return "*" * len(key)
    return f"{key[:6]}…{key[-4:]}"


def _status() -> dict:
    configured = bool(config.ANTHROPIC_API_KEY)
    return {
        "configured": configured,
        "masked": _masked(config.ANTHROPIC_API_KEY) if configured else None,
        "source": "runtime" if config.ANTHROPIC_API_KEY_SET_AT_RUNTIME else "env",
        # Not a secret -- shown in full so a reviewer can confirm which gateway
        # (e.g. an internal Anthropic-API-compatible Bedrock proxy) is in use.
        "baseUrl": config.ANTHROPIC_BASE_URL or None,
        "baseUrlSource": "runtime" if config.ANTHROPIC_BASE_URL_SET_AT_RUNTIME else "env",
        "model": config.AI_MODEL,
        "modelSource": "runtime" if config.AI_MODEL_SET_AT_RUNTIME else "env",
        "defaultModel": config.DEFAULT_AI_MODEL,
    }


@router.get("/anthropic-key")
def get_anthropic_key_status():
    return _status()


@router.put("/anthropic-key")
def set_anthropic_key(body: SetApiKeyRequest):
    if body.api_key is not None:
        key = body.api_key.strip()
        config.ANTHROPIC_API_KEY = key
        config.ANTHROPIC_API_KEY_SET_AT_RUNTIME = bool(key)

    if body.base_url is not None:
        base_url = config.normalize_base_url(body.base_url)
        config.ANTHROPIC_BASE_URL = base_url
        config.ANTHROPIC_BASE_URL_SET_AT_RUNTIME = bool(base_url)

    if body.model is not None:
        model = body.model.strip()
        if model:
            config.AI_MODEL = model
            config.AI_MODEL_SET_AT_RUNTIME = True
        else:
            config.AI_MODEL = config.DEFAULT_AI_MODEL
            config.AI_MODEL_SET_AT_RUNTIME = False

    return _status()


@router.post("/anthropic-key/test")
def test_anthropic_connection():
    """Make one minimal real request (max_tokens=1) with whichever key/base
    URL/model are currently active -- the exact same client construction
    every other AI feature uses -- so a green result here means "Ask AI"
    will actually work, not just that the key is present or well-formed."""
    try:
        result = test_connection()
    except AdjudicationError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc
    return {"ok": True, **result}
