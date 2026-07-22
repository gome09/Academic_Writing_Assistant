# -*- coding: utf-8 -*-
from src import llm_enhancer


class _FakeResp:
    def __init__(self, content):
        msg = type("M", (), {"content": content})()
        self.choices = [type("C", (), {"message": msg})()]


def test_chat_passes_response_format(monkeypatch):
    calls = []

    class FakeCompletions:
        @staticmethod
        def create(**kw):
            calls.append(kw)
            return _FakeResp('{"a": 1}')

    fake_client = type("Cl", (), {})()
    fake_client.chat = type("Ch", (), {"completions": FakeCompletions})()
    monkeypatch.setattr(llm_enhancer, "_client", lambda: fake_client)

    llm_enhancer._chat("s", "u", json_mode=True)
    assert calls[0]["response_format"] == {"type": "json_object"}


def test_chat_plain_has_no_response_format(monkeypatch):
    calls = []

    class FakeCompletions:
        @staticmethod
        def create(**kw):
            calls.append(kw)
            return _FakeResp("ok")

    fake_client = type("Cl", (), {})()
    fake_client.chat = type("Ch", (), {"completions": FakeCompletions})()
    monkeypatch.setattr(llm_enhancer, "_client", lambda: fake_client)

    llm_enhancer._chat("s", "u")
    assert "response_format" not in calls[0]


def test_chat_falls_back_when_endpoint_rejects_json_mode(monkeypatch):
    calls = []

    class FakeCompletions:
        @staticmethod
        def create(**kw):
            calls.append(kw)
            if "response_format" in kw:
                raise TypeError("unexpected keyword")
            return _FakeResp('{"a": 1}')

    fake_client = type("Cl", (), {})()
    fake_client.chat = type("Ch", (), {"completions": FakeCompletions})()
    monkeypatch.setattr(llm_enhancer, "_client", lambda: fake_client)

    assert llm_enhancer._chat("s", "u", json_mode=True) == '{"a": 1}'
    assert len(calls) == 2 and "response_format" not in calls[1]


def test_chat_json_requests_json_mode(monkeypatch):
    seen = {}

    def fake_chat(system, user, json_mode=False):
        seen["json_mode"] = json_mode
        return '{"a": 1}'

    monkeypatch.setattr(llm_enhancer, "_chat", fake_chat)
    assert llm_enhancer._chat_json("s", "u") == {"a": 1}
    assert seen["json_mode"] is True
