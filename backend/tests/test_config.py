"""Regression tests for two easy-to-hit config mistakes:

1. A stray leading/trailing newline or space in the ANTHROPIC_API_KEY (or
   ANTHROPIC_BASE_URL) env var -- an easy mistake when copy-pasting into a
   platform's environment variable UI -- must not be sent to Anthropic
   verbatim; the SDK doesn't trim it, so an otherwise-valid key with e.g. a
   trailing "\n" fails auth with a confusing 401.

2. A base URL that already ends in "/v1" -- natural to include, since
   several other OpenAI/Anthropic-compatible tools *do* want it in their
   base URL -- must not be sent to the Anthropic SDK as-is: the SDK always
   requests "/v1/messages" itself, so the request becomes
   ".../v1/v1/messages" and a real gateway 404s.
"""
import importlib
import sys


def _reload_config(monkeypatch, **env):
    for key, value in env.items():
        monkeypatch.setenv(key, value)
    for mod in list(sys.modules):
        if mod == "app.config" or (mod.startswith("app.") and mod != "app.config"):
            del sys.modules[mod]
    if "app" in sys.modules:
        del sys.modules["app"]
    return importlib.import_module("app.config")


def test_api_key_env_var_is_stripped_of_whitespace(monkeypatch):
    config = _reload_config(monkeypatch, ANTHROPIC_API_KEY="  sk-ant-abcdefghijklmnop\n")
    assert config.ANTHROPIC_API_KEY == "sk-ant-abcdefghijklmnop"


def test_base_url_env_var_is_stripped_of_whitespace(monkeypatch):
    config = _reload_config(monkeypatch, ANTHROPIC_BASE_URL="  https://bedrock-gateway.example.mil \n")
    assert config.ANTHROPIC_BASE_URL == "https://bedrock-gateway.example.mil"


def test_base_url_env_var_trailing_v1_is_stripped(monkeypatch):
    config = _reload_config(monkeypatch, ANTHROPIC_BASE_URL="https://bedrock-gateway.example.mil/v1")
    assert config.ANTHROPIC_BASE_URL == "https://bedrock-gateway.example.mil"


def test_normalize_base_url_strips_trailing_v1_and_slash():
    from app.config import normalize_base_url

    assert normalize_base_url("https://gw.example.mil/v1") == "https://gw.example.mil"
    assert normalize_base_url("https://gw.example.mil/v1/") == "https://gw.example.mil"
    assert normalize_base_url("https://gw.example.mil/") == "https://gw.example.mil"
    assert normalize_base_url("https://gw.example.mil") == "https://gw.example.mil"
    assert normalize_base_url("  https://gw.example.mil/v1  ") == "https://gw.example.mil"
    assert normalize_base_url("") == ""
    # A path that merely ends with "v1" as part of a longer segment isn't touched.
    assert normalize_base_url("https://gw.example.mil/apiv1") == "https://gw.example.mil/apiv1"
