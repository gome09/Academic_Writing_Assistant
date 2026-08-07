# -*- coding: utf-8 -*-
"""T3-1: 多轮润色测试——三档改写、标记、可降级。"""
from __future__ import annotations

from src import llm_enhancer
from src.llm_enhancer import POLISH_MARK, polish_paragraphs


def _thesis(chapters):
    return {"title": "T", "author": "A", "abstract": "", "abstract_en": "",
            "keywords": [], "keywords_en": [], "chapters": chapters}


def _ch(title, paras):
    return {"title": title, "level": 1, "paras": paras, "subs": [],
            "tables": [], "images": [], "blocks": [], "section_role": "body"}


def test_polish_adds_marker(monkeypatch):
    """润色后的段落带 POLISH_MARK。"""
    thesis = _thesis([_ch("绪论", [
        "本研究探讨深度学习在校园垃圾图像分类中的应用，具有重要的理论意义和实践价值。" * 2,
    ])])

    def fake_chat_json(sys_prompt, user):
        return {"0": "本研究深入探讨深度学习在校园垃圾图像分类领域的应用，"
                     "具有深远的理论意义与广泛的实践价值。"}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    n = polish_paragraphs(thesis, "standard")
    assert n == 1
    assert POLISH_MARK in thesis["chapters"][0]["paras"][0]


def test_polish_preserves_on_failure(monkeypatch, capsys):
    """LLM 失败时保留原文，不中断。"""
    original = "这是一段足够长的正文段落用于测试润色功能的降级行为。" * 2
    thesis = _thesis([_ch("绪论", [original])])

    def boom(s, u):
        raise RuntimeError("api down")

    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    n = polish_paragraphs(thesis, "standard")
    assert n == 0
    assert thesis["chapters"][0]["paras"][0] == original
    assert "润色失败" in capsys.readouterr().out


def test_polish_skips_short_and_placeholder(monkeypatch):
    """过短段落和占位符跳过。"""
    thesis = _thesis([_ch("绪论", [
        "<请填写>",           # 占位符
        "短",                  # 过短
        "这是一段足够长的正文段落用于测试润色功能是否正确跳过短段落。" * 2,  # 可润色
    ])])

    calls = []

    def fake_chat_json(sys_prompt, user):
        calls.append(user)
        return {"2": "润色后的长段落内容。" * 10}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    n = polish_paragraphs(thesis, "standard")
    assert n == 1
    assert len(calls) == 1
    # 短段落和占位符未被修改
    assert thesis["chapters"][0]["paras"][0] == "<请填写>"
    assert thesis["chapters"][0]["paras"][1] == "短"


def test_polish_skips_already_marked(monkeypatch):
    """已有 AI 标记的段落不再润色。"""
    marked = f"已有标记的段落内容{llm_enhancer.AI_MARK}" + "x" * 30
    thesis = _thesis([_ch("绪论", [marked])])

    def should_not_call(s, u):
        raise AssertionError("不应调用已标记段落的润色")

    monkeypatch.setattr(llm_enhancer, "_chat_json", should_not_call)
    n = polish_paragraphs(thesis, "standard")
    assert n == 0


def test_polish_level_selects_prompt(monkeypatch):
    """不同 level 使用不同系统提示词。"""
    received_sys = []

    def fake_chat_json(sys_prompt, user):
        received_sys.append(sys_prompt)
        return {"0": "润色结果" + "x" * 30}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    para = "这是一段足够长的正文段落用于测试不同润色级别的系统提示词选择。" * 2

    polish_paragraphs(_thesis([_ch("绪论", [para])]), "conservative")
    polish_paragraphs(_thesis([_ch("绪论", [para])]), "strong")
    assert "保守" in received_sys[0]
    assert "强力" in received_sys[1]


def test_polish_invalid_level_falls_back_to_standard(monkeypatch):
    """未知 level 回退 standard。"""
    received_sys = []

    def fake_chat_json(sys_prompt, user):
        received_sys.append(sys_prompt)
        return {"0": "x" * 40}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    para = "这是一段足够长的正文段落用于测试未知润色级别的回退行为。" * 2
    polish_paragraphs(_thesis([_ch("绪论", [para])]), "nonexistent")
    assert "标准" in received_sys[0]
