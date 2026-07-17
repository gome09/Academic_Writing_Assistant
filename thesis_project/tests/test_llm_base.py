# -*- coding: utf-8 -*-
import pytest
from src import llm_enhancer


def test_not_available_without_key(monkeypatch):
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    assert llm_enhancer.is_available() is False


def test_available_with_key(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    assert llm_enhancer.is_available() is True


def test_chat_json_plain(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat", lambda s, u: '{"a": 1}')
    assert llm_enhancer._chat_json("s", "u") == {"a": 1}


def test_chat_json_fenced(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat",
                        lambda s, u: '```json\n{"a": [1, 2]}\n```')
    assert llm_enhancer._chat_json("s", "u") == {"a": [1, 2]}


def test_chat_json_with_surrounding_prose(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat",
                        lambda s, u: '好的，结果如下：{"a": 1} 以上就是结果。')
    assert llm_enhancer._chat_json("s", "u") == {"a": 1}


def test_chat_json_no_json_raises(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat", lambda s, u: "抱歉，我无法处理。")
    with pytest.raises(ValueError):
        llm_enhancer._chat_json("s", "u")
