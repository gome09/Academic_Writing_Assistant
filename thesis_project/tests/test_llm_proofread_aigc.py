# -*- coding: utf-8 -*-
"""T3-2/T3-5: AI 校对 + AIGC 率提示测试。"""
from __future__ import annotations

from src import llm_enhancer
from src.llm_enhancer import AI_MARK, POLISH_MARK, ai_text_ratio, proofread


def _thesis(chapters):
    return {"title": "T", "author": "A", "abstract": "", "abstract_en": "",
            "keywords": [], "keywords_en": [], "chapters": chapters}


def _ch(title, paras):
    return {"title": title, "level": 1, "paras": paras, "subs": [],
            "tables": [], "images": [], "blocks": [], "section_role": "body"}


# ---------------------------------------------------------------------------
#  T3-2: AI 校对
# ---------------------------------------------------------------------------
def test_proofread_returns_issues(monkeypatch):
    """校对返回建议列表，不改原文。"""
    para = "这是是一段有错别字的文本，用于测试校对功能。" * 2
    thesis = _thesis([_ch("绪论", [para])])
    original = thesis["chapters"][0]["paras"][0]

    def fake_chat_json(sys_prompt, user):
        return {"issues": ["'这是是'应为'这是'", "建议统一术语"]}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    issues = proofread(thesis)
    assert len(issues) == 2
    assert "绪论" in issues[0]
    # 原文未被修改
    assert thesis["chapters"][0]["paras"][0] == original


def test_proofread_failure_returns_empty(monkeypatch, capsys):
    """校对失败返回空列表，不中断。"""
    para = "这是一段足够长的正文段落用于测试校对失败时的降级行为。" * 2
    thesis = _thesis([_ch("绪论", [para])])

    def boom(s, u):
        raise RuntimeError("api down")

    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    issues = proofread(thesis)
    assert issues == []
    assert "校对失败" in capsys.readouterr().out


def test_proofread_skips_short_and_placeholder(monkeypatch):
    """过短段落和占位符跳过。"""
    thesis = _thesis([_ch("绪论", ["短", "<请填写>"])])

    def should_not_call(s, u):
        raise AssertionError("不应校验过短段落")

    monkeypatch.setattr(llm_enhancer, "_chat_json", should_not_call)
    issues = proofread(thesis)
    assert issues == []


# ---------------------------------------------------------------------------
#  T3-5: AIGC 率提示
# ---------------------------------------------------------------------------
def test_ai_text_ratio_no_marks():
    """无 AI 标记 -> ratio=0。"""
    thesis = _thesis([_ch("绪论", ["普通正文段落。" * 10])])
    r = ai_text_ratio(thesis)
    assert r["ratio"] == 0
    assert r["ai_marked_segments"] == 0
    assert r["total_chars"] > 0


def test_ai_text_ratio_with_marks():
    """有 AI 标记 -> ratio > 0。"""
    thesis = _thesis([_ch("绪论", [
        f"AI 生成段落 {AI_MARK}",
        "普通段落" * 20,
        f"润色段落 {POLISH_MARK}",
    ])])
    r = ai_text_ratio(thesis)
    assert r["ratio"] > 0
    assert r["ai_marked_segments"] == 2
    assert r["ai_chars"] > 0


def test_ai_text_ratio_empty_thesis():
    """空论文 -> ratio=0。"""
    r = ai_text_ratio({"chapters": []})
    assert r["ratio"] == 0
    assert r["total_chars"] == 0


def test_ai_text_ratio_abstract_included():
    """摘要中的 AI 标记也计入。"""
    thesis = {"chapters": [_ch("绪论", ["正文" * 20])],
              "abstract": f"摘要 {AI_MARK}"}
    r = ai_text_ratio(thesis)
    assert r["ai_marked_segments"] >= 1
    assert r["ai_chars"] > 0
