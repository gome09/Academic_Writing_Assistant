# -*- coding: utf-8 -*-
"""classify_chapters 标题归一化：NFKC + 去空白。"""
import unicodedata

from src import llm_enhancer


def _norm(s):
    s = unicodedata.normalize("NFKC", s).strip()
    return s.replace(" ", "").replace("\u3000", "")


def test_classify_chapters_normalizes_keys(monkeypatch):
    """LLM 返回带前后空白/全角空格的标题，classify_chapters 内部归一化后依然命中。"""

    def fake_chat_json(system, user):
        return {"  实验环境  ": "result", "结论与展望": "conclusion"}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    titles = ["实验环境", "结论与展望"]
    result = llm_enhancer.classify_chapters(titles)
    # 归一化键查询
    assert result.get(_norm("实验环境")) == "result"
    assert result.get(_norm("结论与展望")) == "conclusion"


def test_classify_chapters_invalid_buckets_filtered(monkeypatch):
    """非法 bucket 被过滤。"""

    def fake_chat_json(system, user):
        return {"实验环境": "noise", "结论与展望": "conclusion"}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    result = llm_enhancer.classify_chapters(["实验环境", "结论与展望"])
    # 「noise」非法桶，整条不进入
    assert "实验环境" not in result and _norm("实验环境") not in result
    # 若归一化键查询包含应仍可命中
    assert result.get(_norm("结论与展望")) == "conclusion"
