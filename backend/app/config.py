import os
from pathlib import Path

BACKEND_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ.get("PPTX_DEV_DATA_DIR", BACKEND_DIR.parent / "data"))
DOCUMENTS_DIR = DATA_DIR / "documents"
INDEX_PATH = DATA_DIR / "index.json"

DOCUMENTS_DIR.mkdir(parents=True, exist_ok=True)

def normalize_base_url(url: str) -> str:
    """Strip a trailing "/v1" (with or without a trailing slash) from a
    user-supplied base URL. The Anthropic SDK always requests "/v1/messages"
    itself, so a base URL that already ends in "/v1" -- a very natural thing
    to include, since plenty of other OpenAI/Anthropic-compatible tooling
    *does* require the "/v1" suffix in its base URL -- silently produces
    ".../v1/v1/messages" and a 404 from any real gateway."""
    url = url.strip().rstrip("/")
    if url.endswith("/v1"):
        url = url[: -len("/v1")]
    return url


ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip()
ANTHROPIC_API_KEY_SET_AT_RUNTIME = False  # True once /api/settings/anthropic-key overrides it
ANTHROPIC_BASE_URL = normalize_base_url(os.environ.get("ANTHROPIC_BASE_URL", ""))
ANTHROPIC_BASE_URL_SET_AT_RUNTIME = False  # True once /api/settings/anthropic-key overrides it
AI_MODEL = os.environ.get("PPTX_DEV_AI_MODEL", "claude-sonnet-5").strip()
AI_MODEL_SET_AT_RUNTIME = False  # True once /api/settings/anthropic-key overrides it
DEFAULT_AI_MODEL = AI_MODEL  # what "reset to default" reverts to
AI_ASSISTANT_AUTHOR = "AI Assistant"
