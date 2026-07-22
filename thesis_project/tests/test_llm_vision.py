# -*- coding: utf-8 -*-
import pytest
from src import llm_vision


def test_unavailable_without_vision_model(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.delenv("LLM_VISION_MODEL", raising=False)
    assert llm_vision.is_vision_available() is False


def test_available_with_both_envs(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("LLM_VISION_MODEL", "qwen-vl-plus")
    assert llm_vision.is_vision_available() is True


def test_describe_image_parses_json(monkeypatch):
    captured = {}

    def fake_chat_vision(prompt, image_b64, mime):
        captured["mime"] = mime
        return '{"caption": "系统架构图", "summary": "三层架构：表现层…"}'

    monkeypatch.setattr(llm_vision, "_chat_vision", fake_chat_vision)
    out = llm_vision.describe_image(b"\x89PNG...", ".png")
    assert out == {"caption": "系统架构图", "summary": "三层架构：表现层…"}
    assert captured["mime"] == "image/png"


def test_describe_image_unknown_ext_defaults_png(monkeypatch):
    captured = {}

    def fake_chat_vision(prompt, image_b64, mime):
        captured["mime"] = mime
        return '{"caption": "c", "summary": "s"}'

    monkeypatch.setattr(llm_vision, "_chat_vision", fake_chat_vision)
    llm_vision.describe_image(b"x", ".tiff")
    assert captured["mime"] == "image/png"


def test_describe_image_bad_json_raises(monkeypatch):
    monkeypatch.setattr(llm_vision, "_chat_vision", lambda *a: "看不懂")
    with pytest.raises(ValueError):
        llm_vision.describe_image(b"x", ".png")
