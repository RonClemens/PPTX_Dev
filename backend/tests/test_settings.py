import importlib
import sys
from types import SimpleNamespace

import pytest
from fastapi.testclient import TestClient


@pytest.fixture()
def client(tmp_path, monkeypatch):
    monkeypatch.setenv("PPTX_DEV_DATA_DIR", str(tmp_path))
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    monkeypatch.delenv("ANTHROPIC_BASE_URL", raising=False)
    monkeypatch.delenv("PPTX_DEV_AI_MODEL", raising=False)
    for var in ("ANTHROPIC_AUTH_TOKEN", "PPTX_DEV_AUTH_MODE", "PPTX_DEV_EXTRA_HEADERS"):
        monkeypatch.delenv(var, raising=False)
    for mod in list(sys.modules):
        if mod.startswith("app"):
            del sys.modules[mod]
    main = importlib.import_module("app.main")
    return TestClient(main.app)


def test_status_unconfigured_by_default(client):
    res = client.get("/api/settings/anthropic-key")
    assert res.status_code == 200
    body = res.json()
    assert body == {
        "configured": False,
        "masked": None,
        "source": "env",
        "baseUrl": None,
        "baseUrlSource": "env",
        "model": "claude-sonnet-5",
        "modelSource": "env",
        "defaultModel": "claude-sonnet-5",
        "authMode": "api_key",
        "authModeSource": "env",
        "extraHeaders": [],
        "extraHeadersSource": "env",
    }


def test_set_key_then_status_reflects_it_masked(client):
    res = client.put("/api/settings/anthropic-key", json={"api_key": "sk-ant-abcdefghijklmnop"})
    assert res.status_code == 200
    body = res.json()
    assert body["configured"] is True
    assert body["masked"] == "sk-ant…mnop"
    assert "sk-ant-abcdefghijklmnop" not in res.text  # raw key never echoed back

    status = client.get("/api/settings/anthropic-key").json()
    assert status["configured"] is True
    assert status["source"] == "runtime"
    assert status["masked"] == "sk-ant…mnop"


def test_set_empty_key_clears_it(client):
    client.put("/api/settings/anthropic-key", json={"api_key": "sk-ant-abcdefghijklmnop"})
    res = client.put("/api/settings/anthropic-key", json={"api_key": "   "})
    body = res.json()
    assert body["configured"] is False
    assert body["masked"] is None
    assert client.get("/api/settings/anthropic-key").json()["configured"] is False


def test_set_base_url_is_independent_of_api_key(client):
    """Saving the base URL (from its own form in the settings modal) must not
    touch an already-configured API key, and vice versa."""
    client.put("/api/settings/anthropic-key", json={"api_key": "sk-ant-abcdefghijklmnop"})

    res = client.put(
        "/api/settings/anthropic-key",
        json={"base_url": "https://bedrock-gateway.example.mil"},
    )
    assert res.status_code == 200
    body = res.json()
    assert body["baseUrl"] == "https://bedrock-gateway.example.mil"
    assert body["baseUrlSource"] == "runtime"
    assert body["configured"] is True  # API key untouched
    assert body["masked"] == "sk-ant…mnop"

    # Saving the key again shouldn't clear the base URL either.
    res = client.put("/api/settings/anthropic-key", json={"api_key": "sk-ant-zyxwvutsrqponm"})
    assert res.json()["baseUrl"] == "https://bedrock-gateway.example.mil"


def test_set_base_url_strips_trailing_v1(client):
    """The Anthropic SDK always requests "/v1/messages" itself, so a base URL
    that already ends in "/v1" -- an easy mistake, since several other
    OpenAI/Anthropic-compatible tools *do* want "/v1" in their base URL --
    must be normalized or it silently 404s against a real gateway."""
    res = client.put(
        "/api/settings/anthropic-key",
        json={"base_url": "https://bedrock-gateway.example.mil/v1"},
    )
    assert res.json()["baseUrl"] == "https://bedrock-gateway.example.mil"


def test_clear_base_url_reverts_to_default(client):
    client.put("/api/settings/anthropic-key", json={"base_url": "https://bedrock-gateway.example.mil"})
    res = client.put("/api/settings/anthropic-key", json={"base_url": "   "})
    body = res.json()
    assert body["baseUrl"] is None
    assert body["baseUrlSource"] == "env"


def test_set_model_is_independent_of_key_and_base_url(client):
    """A Bedrock gateway's model IDs commonly differ from the direct API's,
    so the model needs to be overridable at runtime too, independently of
    the key/base URL forms."""
    client.put("/api/settings/anthropic-key", json={"api_key": "sk-ant-abcdefghijklmnop"})

    res = client.put(
        "/api/settings/anthropic-key",
        json={"model": "anthropic.claude-sonnet-4-5-20250929-v1:0"},
    )
    assert res.status_code == 200
    body = res.json()
    assert body["model"] == "anthropic.claude-sonnet-4-5-20250929-v1:0"
    assert body["modelSource"] == "runtime"
    assert body["configured"] is True  # API key untouched
    assert body["baseUrl"] is None  # base URL untouched


def test_clear_model_reverts_to_default(client):
    client.put("/api/settings/anthropic-key", json={"model": "anthropic.claude-sonnet-4-5-20250929-v1:0"})
    res = client.put("/api/settings/anthropic-key", json={"model": "   "})
    body = res.json()
    assert body["model"] == body["defaultModel"] == "claude-sonnet-5"
    assert body["modelSource"] == "env"


def test_runtime_key_and_base_url_are_used_by_adjudication(client, monkeypatch):
    """Setting the key/base URL at runtime should flow through to
    app.ai.assistant's client construction, not just the /settings status
    endpoint."""
    client.put(
        "/api/settings/anthropic-key",
        json={"api_key": "sk-ant-testkeyvalue1234", "base_url": "https://bedrock-gateway.example.mil"},
    )

    import app.ai.assistant as assistant_mod

    captured = {}

    class FakeAnthropic:
        def __init__(self, api_key, base_url=None):
            captured["api_key"] = api_key
            captured["base_url"] = base_url

    monkeypatch.setattr(assistant_mod, "Anthropic", FakeAnthropic)
    assistant_mod._client()
    assert captured["api_key"] == "sk-ant-testkeyvalue1234"
    assert captured["base_url"] == "https://bedrock-gateway.example.mil"


def test_test_connection_without_api_key_returns_502(client):
    res = client.post("/api/settings/anthropic-key/test")
    assert res.status_code == 502
    assert "ANTHROPIC_API_KEY" in res.json()["detail"]


def test_test_connection_success_uses_active_key_model_and_base_url(client, monkeypatch):
    """A successful test must use whatever is *currently active* (runtime
    override if set, else env), make one minimal real request, and report
    back the model/base URL it used -- not just say "ok"."""
    client.put(
        "/api/settings/anthropic-key",
        json={
            "api_key": "sk-ant-testkeyvalue1234",
            "base_url": "https://bedrock-gateway.example.mil",
            "model": "anthropic.claude-sonnet-4-5-20250929-v1:0",
        },
    )

    import app.ai.assistant as assistant_mod

    captured = {}

    class FakeMessages:
        def create(self, **kwargs):
            captured.update(kwargs)
            return SimpleNamespace(id="msg_test123")

    class FakeAnthropic:
        def __init__(self, api_key, base_url=None):
            captured["api_key"] = api_key
            captured["base_url"] = base_url
            self.messages = FakeMessages()

    monkeypatch.setattr(assistant_mod, "Anthropic", FakeAnthropic)
    res = client.post("/api/settings/anthropic-key/test")

    assert res.status_code == 200, res.text
    body = res.json()
    assert body["ok"] is True
    assert body["model"] == "anthropic.claude-sonnet-4-5-20250929-v1:0"
    assert body["baseUrl"] == "https://bedrock-gateway.example.mil"
    assert isinstance(body["latencyMs"], int)
    assert body["responseId"] == "msg_test123"

    assert captured["api_key"] == "sk-ant-testkeyvalue1234"
    assert captured["model"] == "anthropic.claude-sonnet-4-5-20250929-v1:0"
    assert captured["max_tokens"] == 1


def test_test_connection_reports_real_api_error(client, monkeypatch):
    client.put("/api/settings/anthropic-key", json={"api_key": "sk-ant-invalid"})

    import anthropic
    import httpx

    import app.ai.assistant as assistant_mod

    class FakeMessages:
        def create(self, **kwargs):
            raise anthropic.APIConnectionError(request=httpx.Request("POST", "https://api.anthropic.com"))

    class FakeAnthropic:
        def __init__(self, api_key, base_url=None):
            self.messages = FakeMessages()

    monkeypatch.setattr(assistant_mod, "Anthropic", FakeAnthropic)
    res = client.post("/api/settings/anthropic-key/test")

    assert res.status_code == 502
    assert "Claude API error" in res.json()["detail"]



# -- authentication mode / extra headers -------------------------------------------------


def _outgoing_headers(key="sk-test-123"):
    """The credential headers the real SDK client would send, per current settings."""
    import app.ai.assistant as assistant_mod

    client = assistant_mod._client()
    headers = {**client.auth_headers, **client.default_headers}
    return {k.lower(): v for k, v in headers.items()}


def _set(client, **body):
    return client.put("/api/settings/anthropic-key", json=body)


def test_default_mode_sends_x_api_key_only(client):
    _set(client, api_key="sk-test-123")
    h = _outgoing_headers()
    assert h["x-api-key"] == "sk-test-123" and "authorization" not in h


def test_bearer_mode_sends_authorization_bearer_and_no_x_api_key(client):
    res = _set(client, api_key="work-token-abc", auth_mode="bearer")
    assert res.json()["authMode"] == "bearer" and res.json()["authModeSource"] == "runtime"
    h = _outgoing_headers()
    assert h["authorization"] == "Bearer work-token-abc"
    assert "x-api-key" not in h


def test_both_mode_sends_both_headers(client):
    _set(client, api_key="tok", auth_mode="both")
    h = _outgoing_headers()
    assert h["authorization"] == "Bearer tok" and h["x-api-key"] == "tok"


def test_auth_mode_aliases_and_reset(client):
    assert _set(client, auth_mode="Auth-Token").json()["authMode"] == "bearer"
    assert _set(client, auth_mode="x-api-key").json()["authMode"] == "api_key"
    _set(client, auth_mode="bearer")
    res = _set(client, auth_mode="")
    assert res.json()["authMode"] == "api_key" and res.json()["authModeSource"] == "env"


def test_invalid_auth_mode_is_400_and_changes_nothing(client):
    _set(client, auth_mode="bearer")
    res = _set(client, auth_mode="oauth")
    assert res.status_code == 400 and "auth mode" in res.json()["detail"]
    assert client.get("/api/settings/anthropic-key").json()["authMode"] == "bearer"


def test_extra_headers_are_sent_but_values_are_never_returned(client):
    _set(client, api_key="k")
    res = _set(client, extra_headers="X-Tenant-Id: acme-secret-42\n# comment\nOcp-Apim-Subscription-Key: s3cr3t")
    body = res.json()
    assert body["extraHeaders"] == ["Ocp-Apim-Subscription-Key", "X-Tenant-Id"]
    assert body["extraHeadersSource"] == "runtime"
    assert "acme-secret-42" not in res.text and "s3cr3t" not in res.text
    h = _outgoing_headers()
    assert h["x-tenant-id"] == "acme-secret-42" and h["ocp-apim-subscription-key"] == "s3cr3t"


def test_extra_headers_accept_json_and_can_be_cleared(client):
    res = _set(client, extra_headers='{"X-One": "1", "X-Two": "2"}')
    assert res.json()["extraHeaders"] == ["X-One", "X-Two"]
    res = _set(client, extra_headers="")
    assert res.json()["extraHeaders"] == [] and res.json()["extraHeadersSource"] == "env"


@pytest.mark.parametrize(
    "bad",
    ["no colon here", "Bad Name: x", "X-Ok: fine\n: nothing", '{"a": ', "[1, 2]", "X-Evil: a\rb"],
)
def test_malformed_extra_headers_are_400_and_change_nothing(client, bad):
    _set(client, extra_headers="X-Keep: yes")
    res = _set(client, extra_headers=bad)
    assert res.status_code == 400
    assert client.get("/api/settings/anthropic-key").json()["extraHeaders"] == ["X-Keep"]


def test_env_auth_token_alias_defaults_to_bearer(tmp_path, monkeypatch):
    monkeypatch.setenv("PPTX_DEV_DATA_DIR", str(tmp_path))
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    monkeypatch.setenv("ANTHROPIC_AUTH_TOKEN", "env-work-token")
    monkeypatch.setenv("PPTX_DEV_EXTRA_HEADERS", "X-From-Env: 1")
    for mod in list(sys.modules):
        if mod.startswith("app"):
            del sys.modules[mod]
    c = TestClient(importlib.import_module("app.main").app)
    body = c.get("/api/settings/anthropic-key").json()
    assert body["configured"] is True and body["authMode"] == "bearer" and body["extraHeaders"] == ["X-From-Env"]
    h = _outgoing_headers()
    assert h["authorization"] == "Bearer env-work-token" and "x-api-key" not in h


def test_errors_name_the_url_status_and_likely_cause():
    import anthropic
    import httpx

    import app.ai.assistant as assistant_mod

    req = httpx.Request("POST", "https://gateway.example.mil/v1/messages")
    resp = httpx.Response(401, request=req, json={"error": {"message": "invalid bearer token"}})
    msg = assistant_mod.describe_api_error(
        anthropic.AuthenticationError("401", response=resp, body=None)
    )
    assert "POST https://gateway.example.mil/v1/messages returned HTTP 401" in msg
    assert "invalid bearer token" in msg and "Authentication mode" in msg

    resp404 = httpx.Response(404, request=req, text="model not found")
    msg = assistant_mod.describe_api_error(anthropic.NotFoundError("404", response=resp404, body=None))
    assert "HTTP 404" in msg and "base URL" in msg and "model" in msg

    conn = anthropic.APIConnectionError(request=req)
    assert "could not connect to https://gateway.example.mil/v1/messages" in assistant_mod.describe_api_error(conn)


def test_bearer_mode_never_leaks_an_env_api_key_as_x_api_key(client, monkeypatch):
    """With api_key=None the SDK silently reads ANTHROPIC_API_KEY from the
    environment and prefers it -- which would send x-api-key instead of the
    Bearer token the gateway needs."""
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-from-env-should-not-be-sent")
    _set(client, api_key="work-token", auth_mode="bearer")
    h = _outgoing_headers()
    assert h["authorization"] == "Bearer work-token" and "x-api-key" not in h


# -- end to end against a fake gateway -----------------------------------------------------


@pytest.fixture()
def fake_gateway():
    """A tiny local HTTP server standing in for a corporate Anthropic-compatible
    gateway: records every request and answers with a minimal Messages response."""
    import json
    import threading
    from http.server import BaseHTTPRequestHandler, HTTPServer

    seen: list[dict] = []
    behavior = {"status": 200}

    class Handler(BaseHTTPRequestHandler):
        def do_POST(self):
            length = int(self.headers.get("content-length", 0))
            body = self.rfile.read(length)
            seen.append({"path": self.path, "headers": {k.lower(): v for k, v in self.headers.items()}, "body": body})
            if behavior["status"] != 200:
                payload = json.dumps({"type": "error", "error": {"type": "authentication_error", "message": "invalid bearer token"}})
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
    _set(client, api_key="work-token-xyz", auth_mode="bearer", model="gw-model",
         base_url=f"{fake_gateway.url}/anthropic/", extra_headers="X-Tenant-Id: acme")
    res = client.post("/api/settings/anthropic-key/test")
    assert res.status_code == 200, res.text
    assert res.json()["ok"] is True and res.json()["model"] == "gw-model"

    req = fake_gateway.seen[0]
    assert req["path"] == "/anthropic/v1/messages"  # path prefix on the base URL is honored
    assert req["headers"]["authorization"] == "Bearer work-token-xyz"
    assert "x-api-key" not in req["headers"]
    assert req["headers"]["x-tenant-id"] == "acme"
    assert b'"model":"gw-model"' in req["body"].replace(b" ", b"")


def test_gateway_rejection_surfaces_url_status_body_and_hint(client, fake_gateway):
    fake_gateway.behavior["status"] = 401
    _set(client, api_key="wrong", base_url=fake_gateway.url)
    res = client.post("/api/settings/anthropic-key/test")
    assert res.status_code == 502
    detail = res.json()["detail"]
    assert f"POST {fake_gateway.url}/v1/messages returned HTTP 401" in detail
    assert "invalid bearer token" in detail and "Authentication mode" in detail


def test_env_extra_headers_accept_literal_backslash_n_separators(tmp_path, monkeypatch):
    monkeypatch.setenv("PPTX_DEV_DATA_DIR", str(tmp_path))
    monkeypatch.setenv("PPTX_DEV_EXTRA_HEADERS", "X-One: 1\\nX-Two: 2")
    for mod in list(sys.modules):
        if mod.startswith("app"):
            del sys.modules[mod]
    c = TestClient(importlib.import_module("app.main").app)
    assert c.get("/api/settings/anthropic-key").json()["extraHeaders"] == ["X-One", "X-Two"]
