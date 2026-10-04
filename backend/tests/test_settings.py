"""AI connection settings are STATELESS on the server: the browser sends them
as x-ai-* headers on the requests that call the AI, and the server keeps
none of them. These tests pin that guarantee down."""
import base64
import importlib
import json
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from fastapi.testclient import TestClient

SECRET = "sk-secret-token-0123456789"
BASE = "https://gateway.example.mil/anthropic"


def _fresh_app(tmp_path, monkeypatch, **env):
    monkeypatch.setenv("PPTX_DEV_DATA_DIR", str(tmp_path))
    for var in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_BASE_URL", "PPTX_DEV_AI_MODEL",
                "PPTX_DEV_AUTH_MODE", "PPTX_DEV_EXTRA_HEADERS"):
        monkeypatch.delenv(var, raising=False)
    for k, v in env.items():
        monkeypatch.setenv(k, v)
    for mod in list(sys.modules):
        if mod.startswith("app"):
            del sys.modules[mod]
    return importlib.import_module("app.main")


@pytest.fixture()
def client(tmp_path, monkeypatch):
    return TestClient(_fresh_app(tmp_path, monkeypatch).app)


def ai_headers(key=None, base_url=None, auth_mode=None, extra=None, model=None) -> dict:
    h = {}
    if key is not None:
        h["x-ai-key"] = key
    if base_url is not None:
        h["x-ai-base-url"] = base_url
    if auth_mode is not None:
        h["x-ai-auth-mode"] = auth_mode
    if extra is not None:
        h["x-ai-extra-headers"] = base64.b64encode(extra.encode()).decode()
    if model is not None:
        h["x-ai-model"] = model
    return h


def outgoing(headers: dict) -> dict:
    """The credential headers the real SDK client would send for a request
    carrying `headers` (lower-cased)."""
    import app.ai.assistant as assistant_mod
    from app.ai import credentials

    token = credentials.set_for_request(credentials.from_request_headers(headers))
    try:
        c = assistant_mod._client()
        return {k.lower(): v for k, v in {**c.auth_headers, **c.default_headers}.items()}
    finally:
        credentials.reset(token)


# -- the server stores nothing -----------------------------------------------------------


def test_server_status_never_contains_user_credentials_and_says_it_stores_none(client):
    body = client.get("/api/settings/anthropic-key").json()
    assert body == {
        "storesCredentials": False, "serverKeyConfigured": False, "baseUrl": None, "authMode": "api_key",
        "extraHeaders": [], "model": "claude-sonnet-5", "defaultModel": "claude-sonnet-5",
    }


def test_there_is_no_endpoint_that_saves_credentials(client):
    for method in ("put", "post", "patch"):
        res = getattr(client, method)("/api/settings/anthropic-key", json={"api_key": SECRET, "base_url": BASE})
        assert res.status_code in (404, 405), (method, res.status_code)
    # and nothing leaked into the status either
    assert SECRET not in client.get("/api/settings/anthropic-key").text


def test_credentials_exist_only_for_the_request_that_carried_them(client):
    """A request with a key must not make later requests (from anyone) see it."""
    with_key = client.post("/api/settings/anthropic-key/test", headers=ai_headers(key=SECRET, base_url="http://127.0.0.1:1"))
    assert with_key.status_code == 502 and "could not connect" in with_key.json()["detail"]  # it tried to use them
    without = client.post("/api/settings/anthropic-key/test")
    assert without.status_code == 502 and "No API key available" in without.json()["detail"]
    assert client.get("/api/settings/anthropic-key").json()["serverKeyConfigured"] is False


def test_nothing_is_written_to_disk_or_module_state(tmp_path, monkeypatch):
    main = _fresh_app(tmp_path, monkeypatch)
    c = TestClient(main.app)
    c.post("/api/documents", files={"file": ("sample.pptx", (Path(__file__).parent / "fixtures" / "sample.pptx").read_bytes(), "x")})
    c.post("/api/settings/anthropic-key/test", headers=ai_headers(
        key=SECRET, base_url="http://127.0.0.1:1/private-path", extra="X-Tenant: tenant-secret-77", model="m"))
    # 1. no file under the data directory mentions any of it
    for f in tmp_path.rglob("*"):
        if f.is_file():
            data = f.read_bytes()
            assert SECRET.encode() not in data and b"private-path" not in data and b"tenant-secret-77" not in data, f
    # 2. no module-level state in the app holds it
    import app.config as cfg

    assert SECRET not in repr(vars(cfg)) and "private-path" not in repr(vars(cfg))
    assert cfg.ANTHROPIC_API_KEY == "" and cfg.ANTHROPIC_BASE_URL == ""


def test_credential_object_repr_never_shows_the_secret():
    from app.ai import credentials

    token = credentials.set_for_request(credentials.from_request_headers(ai_headers(key=SECRET)))
    try:
        assert SECRET not in repr(credentials.effective())
    finally:
        credentials.reset(token)


# -- how credentials are applied --------------------------------------------------------


def test_default_mode_sends_x_api_key_only():
    h = outgoing(ai_headers(key=SECRET))
    assert h["x-api-key"] == SECRET and "authorization" not in h


def test_bearer_mode_sends_authorization_bearer_and_no_x_api_key():
    h = outgoing(ai_headers(key="work-token-abc", auth_mode="bearer"))
    assert h["authorization"] == "Bearer work-token-abc" and "x-api-key" not in h


def test_both_mode_sends_both_headers():
    h = outgoing(ai_headers(key="tok", auth_mode="both"))
    assert h["authorization"] == "Bearer tok" and h["x-api-key"] == "tok"


def test_bearer_mode_never_leaks_an_env_api_key_as_x_api_key(tmp_path, monkeypatch):
    """With api_key=None the SDK silently reads ANTHROPIC_API_KEY from the
    environment and prefers it -- which would send x-api-key instead of the
    Bearer token the gateway needs."""
    _fresh_app(tmp_path, monkeypatch)
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-from-env-should-not-be-sent")
    h = outgoing(ai_headers(key="work-token", auth_mode="bearer"))
    assert h["authorization"] == "Bearer work-token" and "x-api-key" not in h


def test_auth_mode_aliases():
    for alias, expected in {"Auth-Token": "bearer", "x-api-key": "api_key", "BOTH": "both", "authorization": "bearer"}.items():
        h = outgoing(ai_headers(key="k", auth_mode=alias))
        assert ("authorization" in h) == (expected in ("bearer", "both"))


def test_extra_headers_are_sent_with_the_request():
    h = outgoing(ai_headers(key="k", extra="X-Tenant-Id: acme-42\n# comment\nOcp-Apim-Subscription-Key: s3cr3t"))
    assert h["x-tenant-id"] == "acme-42" and h["ocp-apim-subscription-key"] == "s3cr3t"
    h = outgoing(ai_headers(key="k", extra='{"X-One": "1"}'))
    assert h["x-one"] == "1"


def test_environment_is_only_a_per_field_fallback(tmp_path, monkeypatch):
    _fresh_app(tmp_path, monkeypatch, ANTHROPIC_AUTH_TOKEN="env-work-token", PPTX_DEV_EXTRA_HEADERS="X-From-Env: 1\\nX-Two: 2")
    h = outgoing({})  # no x-ai-* headers at all: pure environment
    assert h["authorization"] == "Bearer env-work-token" and "x-api-key" not in h
    assert h["x-from-env"] == "1" and h["x-two"] == "2"
    # a request-supplied token replaces the env one; unspecified fields still fall back
    h = outgoing(ai_headers(key="my-own-token"))
    assert h["authorization"] == "Bearer my-own-token" and h["x-from-env"] == "1"


# -- validation ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "hdrs, needle",
    [
        (ai_headers(auth_mode="oauth"), "auth mode"),
        (ai_headers(base_url="ftp://x.example"), "http://"),
        (ai_headers(base_url="not a url"), "http://"),
        (ai_headers(base_url="http://169.254.169.254/latest"), "not allowed"),
        (ai_headers(base_url="http://metadata.google.internal"), "not allowed"),
        ({"x-ai-extra-headers": "%%%not-base64%%%"}, "base64"),
        (ai_headers(extra="no colon here"), "Header-Name"),
        (ai_headers(extra="Bad Name: x"), "invalid header name"),
        (ai_headers(extra="Host: evil.example"), "can't be overridden"),
        (ai_headers(extra="Content-Length: 1"), "can't be overridden"),
        (ai_headers(extra='{"a": '), "JSON"),
        (ai_headers(extra="[1, 2]"), "Header-Name"),
    ],
)
def test_malformed_settings_are_rejected_with_400(client, hdrs, needle):
    res = client.post("/api/settings/anthropic-key/test", headers={**hdrs})
    assert res.status_code == 400, res.text
    assert "Invalid AI setting" in res.json()["detail"] and needle in res.json()["detail"]


def test_localhost_gateways_are_allowed_for_local_use(client):
    res = client.post("/api/settings/anthropic-key/test", headers=ai_headers(key="k", base_url="http://127.0.0.1:9"))
    assert res.status_code == 502  # allowed; simply nothing listening


# -- end to end against a fake gateway -------------------------------------------------------


@pytest.fixture()
def fake_gateway():
    """A tiny local HTTP server standing in for a corporate Anthropic-compatible
    gateway: records every request and answers with a minimal Messages response."""
    import threading
    from http.server import BaseHTTPRequestHandler, HTTPServer

    seen: list[dict] = []
    behavior = {"status": 200, "error_text": "invalid bearer token"}

    class Handler(BaseHTTPRequestHandler):
        def do_POST(self):
            length = int(self.headers.get("content-length", 0))
            body = self.rfile.read(length)
            seen.append({"path": self.path, "headers": {k.lower(): v for k, v in self.headers.items()}, "body": body})
            if behavior["status"] != 200:
                payload = json.dumps({"type": "error", "error": {"type": "authentication_error", "message": behavior["error_text"]}})
                self.send_response(behavior["status"])
            else:
                payload = json.dumps({
                    "id": "msg_test", "type": "message", "role": "assistant", "model": "gw-model",
                    "content": [{"type": "text", "text": "ok"}], "stop_reason": "end_turn",
                    "usage": {"input_tokens": 1, "output_tokens": 1},
                })
                self.send_response(200)
            self.send_header("content-type", "application/json")
            self.send_header("content-length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload.encode())

        def log_message(self, *args):
            pass

    server = HTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    yield SimpleNamespace(url=f"http://127.0.0.1:{server.server_port}", seen=seen, behavior=behavior)
    server.shutdown()


def test_test_connection_reaches_a_gateway_with_bearer_token_extra_header_and_path_prefix(client, fake_gateway):
    res = client.post("/api/settings/anthropic-key/test", headers=ai_headers(
        key="work-token-xyz", auth_mode="bearer", model="gw-model",
        base_url=f"{fake_gateway.url}/anthropic/", extra="X-Tenant-Id: acme"))
    assert res.status_code == 200, res.text
    assert res.json()["ok"] is True and res.json()["model"] == "gw-model"
    req = fake_gateway.seen[0]
    assert req["path"] == "/anthropic/v1/messages"  # path prefix on the base URL is honored
    assert req["headers"]["authorization"] == "Bearer work-token-xyz" and "x-api-key" not in req["headers"]
    assert req["headers"]["x-tenant-id"] == "acme"
    assert b'"model":"gw-model"' in req["body"].replace(b" ", b"")


def test_ai_features_use_the_request_credentials_end_to_end(client, fake_gateway):
    """A real AI route (adjudicate) with credentials only in headers."""
    fake_gateway.behavior["status"] = 200
    doc = client.post("/api/documents", files={"file": ("sample.pptx", (Path(__file__).parent / "fixtures" / "sample.pptx").read_bytes(), "x")}).json()["meta"]["id"]
    res = client.post(f"/api/documents/{doc}/ai-review", json={}, headers=ai_headers(
        key="k-123", base_url=fake_gateway.url, auth_mode="bearer"))
    # The fake gateway returns plain text (no tool call) -> the app reports that cleanly,
    # but the point is the request reached the gateway carrying the right credential.
    assert res.status_code == 502 and "did not return" in res.json()["detail"]
    assert fake_gateway.seen[0]["headers"]["authorization"] == "Bearer k-123"
    assert fake_gateway.seen[0]["path"] == "/v1/messages"


def test_gateway_rejection_surfaces_url_status_body_and_hint(client, fake_gateway):
    fake_gateway.behavior["status"] = 401
    res = client.post("/api/settings/anthropic-key/test", headers=ai_headers(key="wrong", base_url=fake_gateway.url))
    assert res.status_code == 502
    detail = res.json()["detail"]
    assert f"POST {fake_gateway.url}/v1/messages returned HTTP 401" in detail
    assert "invalid bearer token" in detail and "Authentication mode" in detail


def test_secrets_are_scrubbed_from_error_text(client, fake_gateway):
    """A gateway that echoes the credential back in its error body must not
    cause it to be displayed or logged by this app."""
    fake_gateway.behavior.update(status=401, error_text=f"bad token {SECRET} / tenant tenant-secret-77")
    res = client.post("/api/settings/anthropic-key/test", headers=ai_headers(
        key=SECRET, base_url=fake_gateway.url, extra="X-Tenant: tenant-secret-77"))
    assert res.status_code == 502
    assert SECRET not in res.text and "tenant-secret-77" not in res.text and "***" in res.text


def test_credentials_in_a_base_url_are_scrubbed_from_error_text(client):
    res = client.post("/api/settings/anthropic-key/test", headers=ai_headers(
        key="k", base_url="http://someuser:somepassword@127.0.0.1:1"))
    assert res.status_code == 502
    assert "somepassword" not in res.text and "someuser" not in res.text


# -- error descriptions (unit) --------------------------------------------------------------


def test_errors_name_the_url_status_and_likely_cause():
    import anthropic
    import httpx

    import app.ai.assistant as assistant_mod

    req = httpx.Request("POST", "https://gateway.example.mil/v1/messages")
    resp = httpx.Response(401, request=req, json={"error": {"message": "invalid bearer token"}})
    msg = assistant_mod.describe_api_error(anthropic.AuthenticationError("401", response=resp, body=None))
    assert "POST https://gateway.example.mil/v1/messages returned HTTP 401" in msg
    assert "invalid bearer token" in msg and "Authentication mode" in msg

    resp404 = httpx.Response(404, request=req, text="model not found")
    msg = assistant_mod.describe_api_error(anthropic.NotFoundError("404", response=resp404, body=None))
    assert "HTTP 404" in msg and "base URL" in msg and "model" in msg

    conn = anthropic.APIConnectionError(request=req)
    assert "could not connect to https://gateway.example.mil/v1/messages" in assistant_mod.describe_api_error(conn)


def test_adjudicate_comment_wraps_anthropic_sdk_errors_into_adjudication_error():
    import anthropic
    import httpx

    import app.ai.assistant as assistant_mod

    fake = SimpleNamespace(messages=SimpleNamespace(create=lambda **kw: (_ for _ in ()).throw(
        anthropic.APIConnectionError(request=httpx.Request("POST", "https://api.anthropic.com")))))
    from unittest.mock import patch

    with patch.object(assistant_mod, "_client", return_value=fake):
        with pytest.raises(assistant_mod.AdjudicationError):
            assistant_mod.adjudicate_comment(comment_id="0-1", comment_author="J", comment_text="?", context_markdown="x")
