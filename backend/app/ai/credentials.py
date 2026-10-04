"""Per-request AI credentials. Nothing here is ever stored.

The browser keeps the user's API key / token, base URL, auth mode, extra
headers and model on the user's own device and attaches them, as request
headers, to the few requests that actually call the AI. A middleware parses
them into a `ContextVar` for the duration of that single request; the AI
client reads them from there; when the response is sent the context is reset
and they're gone. They are never written to disk, never kept in a global or
cache, never put in the audit logs, and are scrubbed from any error text
(`redact`).

Environment variables (ANTHROPIC_API_KEY, ...) remain as an optional,
operator-chosen fallback for any field the request doesn't supply.
"""
from __future__ import annotations

import base64
import ipaddress
import json
import re
from contextvars import ContextVar
from dataclasses import dataclass
from urllib.parse import urlsplit

from .. import config

HEADER_KEY = "x-ai-key"
HEADER_BASE_URL = "x-ai-base-url"
HEADER_AUTH_MODE = "x-ai-auth-mode"
HEADER_EXTRA_HEADERS = "x-ai-extra-headers"  # base64(JSON object) -- header values can't carry newlines
HEADER_MODEL = "x-ai-model"
REQUEST_HEADERS = (HEADER_KEY, HEADER_BASE_URL, HEADER_AUTH_MODE, HEADER_EXTRA_HEADERS, HEADER_MODEL)


class CredentialError(ValueError):
    """The request carried a malformed credential setting."""


@dataclass(frozen=True)
class AiCredentials:
    api_key: str
    base_url: str
    auth_mode: str
    extra_headers: dict[str, str]
    model: str

    def __repr__(self) -> str:  # never leak secrets via logging/debugging
        return f"AiCredentials(key={'set' if self.api_key else 'unset'}, base_url={self.base_url!r}, mode={self.auth_mode!r})"


@dataclass(frozen=True)
class _RequestOverrides:
    api_key: str | None = None
    base_url: str | None = None
    auth_mode: str | None = None
    extra_headers: dict[str, str] | None = None
    model: str | None = None


_current: ContextVar[_RequestOverrides | None] = ContextVar("pptx_dev_ai_credentials", default=None)


def _check_base_url(url: str) -> str:
    url = config.normalize_base_url(url)
    if not url:
        return ""
    parts = urlsplit(url)
    if parts.scheme not in ("http", "https") or not parts.hostname:
        raise CredentialError("base URL must start with http:// or https://")
    host = parts.hostname.lower()
    # This server makes the request on the user's behalf, so refuse the one
    # target that is dangerous on every cloud host: the metadata service.
    if host == "metadata.google.internal" or host == "metadata":
        raise CredentialError("that base URL is not allowed")
    try:
        ip = ipaddress.ip_address(host)
    except ValueError:
        ip = None  # a hostname, not a literal IP
    if ip is not None and ip.is_link_local:
        raise CredentialError("that base URL is not allowed")
    return url


def from_request_headers(headers) -> _RequestOverrides | None:
    """Parse the x-ai-* request headers (case-insensitive mapping). Returns
    None when the request carries none of them. Raises CredentialError on a
    malformed value."""
    get = lambda name: (headers.get(name) or "").strip()  # noqa: E731
    if not any(get(h) for h in REQUEST_HEADERS):
        return None
    mode = get(HEADER_AUTH_MODE)
    extra: dict[str, str] | None = None
    raw_extra = get(HEADER_EXTRA_HEADERS)
    if raw_extra:
        try:
            decoded = base64.b64decode(raw_extra, validate=True).decode("utf-8")
        except Exception as exc:  # noqa: BLE001
            raise CredentialError("extra headers were not valid base64 UTF-8") from exc
        try:
            extra = config.parse_extra_headers(decoded)
        except ValueError as exc:
            raise CredentialError(str(exc)) from exc
    try:
        auth_mode = config.normalize_auth_mode(mode) if mode else None
    except ValueError as exc:
        raise CredentialError(str(exc)) from exc
    return _RequestOverrides(
        api_key=get(HEADER_KEY) or None,
        base_url=_check_base_url(get(HEADER_BASE_URL)) if get(HEADER_BASE_URL) else None,
        auth_mode=auth_mode,
        extra_headers=extra,
        model=get(HEADER_MODEL)[:200] or None,
    )


def set_for_request(overrides: _RequestOverrides | None):
    return _current.set(overrides)


def reset(token) -> None:
    _current.reset(token)


def effective() -> AiCredentials:
    """The credentials for the current request: each field from the request
    if it supplied one, otherwise the environment default."""
    o = _current.get() or _RequestOverrides()
    return AiCredentials(
        api_key=o.api_key if o.api_key is not None else config.ANTHROPIC_API_KEY,
        base_url=o.base_url if o.base_url is not None else config.ANTHROPIC_BASE_URL,
        auth_mode=o.auth_mode if o.auth_mode is not None else config.AUTH_MODE,
        extra_headers=dict(o.extra_headers if o.extra_headers is not None else config.EXTRA_HEADERS),
        model=o.model if o.model is not None else config.AI_MODEL,
    )


_USERINFO_RE = re.compile(r"(?<=//)[^/@\s]+@")


def redact(text: str) -> str:
    """Scrub the current request's secrets (key, extra-header values) and any
    user:password@ in a URL out of text that is about to be shown or logged."""
    c = effective()
    for secret in [c.api_key, *c.extra_headers.values()]:
        if secret and len(secret) >= 4:
            text = text.replace(secret, "***")
    return _USERINFO_RE.sub("***@", text)
