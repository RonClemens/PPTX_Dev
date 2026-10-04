import json
import os
import re
from pathlib import Path

BACKEND_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ.get("PPTX_DEV_DATA_DIR", BACKEND_DIR.parent / "data"))
DOCUMENTS_DIR = DATA_DIR / "documents"
INDEX_PATH = DATA_DIR / "index.json"

# Browser origins allowed to call the API cross-origin (comma-separated, or
# "*"). The built UI is served by this same process, so it never needs CORS;
# set this to an empty string to disable CORS entirely (the local container
# deployment does). Default "*" keeps `npm run dev` + other tooling working.
CORS_ORIGINS = [o.strip() for o in os.environ.get("PPTX_DEV_CORS_ORIGINS", "*").split(",") if o.strip()]

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


AUTH_MODES = ("api_key", "bearer", "both")
_HEADER_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9-]*$")
# Headers that would corrupt the HTTP request itself; never accepted as "extra".
_FORBIDDEN_HEADERS = {"host", "content-length", "transfer-encoding", "connection", "content-type"}


def normalize_auth_mode(value: str) -> str:
    """"api_key" (default) sends the key as the `x-api-key` header, exactly as
    api.anthropic.com expects. "bearer" sends it as `Authorization: Bearer
    <token>` instead -- what most corporate/Bedrock/LiteLLM-style gateways
    want for a "work token" (the same thing Claude Code calls
    ANTHROPIC_AUTH_TOKEN). "both" sends both headers."""
    value = (value or "").strip().lower().replace("-", "_")
    if value in ("", "default"):
        return "api_key"
    if value in ("x_api_key", "apikey"):
        return "api_key"
    if value in ("token", "auth_token", "authorization"):
        return "bearer"
    if value not in AUTH_MODES:
        raise ValueError(f"auth mode must be one of: {', '.join(AUTH_MODES)}")
    return value


def parse_extra_headers(text: str) -> dict[str, str]:
    """Extra HTTP headers to send with every AI request, for gateways that
    need more than a key (a tenant/subscription id, a proxy auth header...).
    Accepts one "Name: value" per line, or a JSON object. Raises ValueError
    with a human-readable message on anything malformed."""
    text = (text or "").strip()
    if not text:
        return {}
    pairs: list[tuple[str, str]] = []
    if text.startswith("{"):
        try:
            obj = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError(f"extra headers look like JSON but don't parse: {exc}") from exc
        if not isinstance(obj, dict):
            raise ValueError("extra headers JSON must be an object of name: value")
        pairs = [(str(k), str(v)) for k, v in obj.items()]
    else:
        for n, line in enumerate(text.splitlines(), start=1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            name, sep, value = line.partition(":")
            if not sep:
                raise ValueError(f'extra headers line {n} must look like "Header-Name: value"')
            pairs.append((name.strip(), value.strip()))
    headers: dict[str, str] = {}
    for name, value in pairs:
        if not _HEADER_NAME_RE.match(name):
            raise ValueError(f"invalid header name: {name!r}")
        if name.lower() in _FORBIDDEN_HEADERS:
            raise ValueError(f"header {name} can't be overridden")
        if "\r" in value or "\n" in value or len(value) > 4096:
            raise ValueError(f"invalid value for header {name}")
        headers[name] = value
    if len(headers) > 20:
        raise ValueError("at most 20 extra headers")
    return headers


# SERVER-SIDE DEFAULTS FROM THE ENVIRONMENT ONLY. Credentials a user enters in
# the browser are never stored here: they travel with each AI request and live
# only for that request (see app/ai/credentials.py). These env vars exist for
# operators who deliberately want a shared, server-held key.
# ANTHROPIC_AUTH_TOKEN is accepted as an alias (it is what Claude Code uses for
# gateways); when only it is set, the token is sent as a Bearer token.
_env_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
_env_token = os.environ.get("ANTHROPIC_AUTH_TOKEN", "").strip()
ANTHROPIC_API_KEY = _env_key or _env_token
ANTHROPIC_BASE_URL = normalize_base_url(os.environ.get("ANTHROPIC_BASE_URL", ""))
AUTH_MODE = normalize_auth_mode(
    os.environ.get("PPTX_DEV_AUTH_MODE", "") or ("bearer" if _env_token and not _env_key else "")
)
try:
    # A one-line env var can't hold real newlines, so a literal backslash-n
    # between headers is accepted too.
    EXTRA_HEADERS = parse_extra_headers(os.environ.get("PPTX_DEV_EXTRA_HEADERS", "").replace("\\n", "\n"))
except ValueError as _exc:  # a bad env value shouldn't stop the app booting
    print(f"warning: ignoring PPTX_DEV_EXTRA_HEADERS: {_exc}")
    EXTRA_HEADERS = {}
AI_MODEL = os.environ.get("PPTX_DEV_AI_MODEL", "claude-sonnet-5").strip()
DEFAULT_AI_MODEL = AI_MODEL
AI_ASSISTANT_AUTHOR = "AI Assistant"
