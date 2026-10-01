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
