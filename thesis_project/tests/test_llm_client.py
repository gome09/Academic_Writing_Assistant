# -*- coding: utf-8 -*-
from src import llm_enhancer


def test_client_default_timeout(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.delenv("LLM_TIMEOUT", raising=False)
    captured = {}

    class FakeOpenAI:
        def __init__(self, **kw):
            captured.update(kw)

    monkeypatch.setattr("openai.OpenAI", FakeOpenAI)
    llm_enhancer._client()
    assert captured["timeout"] == 60.0
    assert captured["max_retries"] == 1


def test_client_timeout_env_override(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("LLM_TIMEOUT", "30")
    captured = {}

    class FakeOpenAI:
        def __init__(self, **kw):
            captured.update(kw)

    monkeypatch.setattr("openai.OpenAI", FakeOpenAI)
    llm_enhancer._client()
    assert captured["timeout"] == 30.0
