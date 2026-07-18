# -*- coding: utf-8 -*-
"""LLM 失败 + 原文足量时，bullets 应该是原文摘要而非 <待补充要点>。"""
from src import llm_enhancer


def test_safe_bullets_returns_paragraph_summary_when_llm_fails(monkeypatch):
    """LLM 异常时，bullets 不再退化为单一占位，而是原文 80 字内摘要。"""
    long = "本章系统介绍了总体架构设计、数据流图与模块划分，" \
           "对核心算法的输入输出与实现细节进行详细说明。"

    def boom(s, u):
        raise RuntimeError("api down")

    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    bullets = llm_enhancer._safe_bullets([long])
    assert len(bullets) >= 1
    assert bullets[0] != "<待补充要点>"
    assert all(len(b) <= 41 for b in bullets)
    # 摘要应回显原文前 40 字
    assert "总体架构" in bullets[0] or "模块" in bullets[0]


def test_safe_bullets_short_text_uses_rules(monkeypatch):
    """段落太短时仍然走 _to_bullets，不抛错。"""

    def boom(s, u):
        raise RuntimeError("api down")

    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    bullets = llm_enhancer._safe_bullets(["太短。"])
    assert bullets
