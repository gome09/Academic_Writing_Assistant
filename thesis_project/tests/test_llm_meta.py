# -*- coding: utf-8 -*-
from src import llm_enhancer
from src.llm_enhancer import AI_MARK
from src.organizer import PLACEHOLDER
from tests.factories import p, doc


def _thesis(**over):
    t = {"title": PLACEHOLDER, "author": PLACEHOLDER,
         "abstract": PLACEHOLDER, "abstract_en": PLACEHOLDER,
         "keywords": [PLACEHOLDER], "keywords_en": [PLACEHOLDER],
         "chapters": [], "references": [], "auto_skeleton": False}
    t.update(over)
    return t


def test_refine_meta_fills_placeholders(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "title": "某垃圾分类系统研究", "author": "张三",
        "abstract": "本文研究了垃圾分类。", "keywords": ["深度学习", "分类"]})
    t = _thesis()
    llm_enhancer.refine_meta(t, [doc([p("正文片段。")])])
    assert t["title"] == "某垃圾分类系统研究"
    assert t["author"] == "张三"
    assert t["abstract"] == f"本文研究了垃圾分类。 {AI_MARK}"
    assert t["keywords"] == ["深度学习", "分类"]


def test_refine_meta_keeps_existing_values(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "title": "LLM乱给的题目", "author": "LLM乱给的作者",
        "abstract": "x", "keywords": ["x"]})
    t = _thesis(title="真题目", author="真作者",
                abstract="真摘要", keywords=["真关键词"])
    llm_enhancer.refine_meta(t, [doc([p("正文。")])])
    assert t["title"] == "真题目"
    assert t["author"] == "真作者"
    assert t["abstract"] == "真摘要"
    assert t["keywords"] == ["真关键词"]


def test_refine_meta_skips_llm_when_nothing_needed(monkeypatch):
    def boom(s, u):
        raise AssertionError("不应调用 LLM")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    t = _thesis(title="a", author="b", abstract="c", keywords=["d"])
    llm_enhancer.refine_meta(t, [doc([p("正文。")])])


def test_translate_abstract(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "abstract_en": "This paper studies waste sorting.",
        "keywords_en": ["deep learning", "classification"]})
    t = _thesis(abstract="本文研究了垃圾分类。", keywords=["深度学习"])
    llm_enhancer.translate_abstract(t)
    assert t["abstract_en"] == f"This paper studies waste sorting. {AI_MARK}"
    assert t["keywords_en"] == ["deep learning", "classification"]


def test_translate_abstract_skipped_when_no_abstract(monkeypatch):
    def boom(s, u):
        raise AssertionError("不应调用 LLM")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    t = _thesis()  # abstract 仍是占位符
    llm_enhancer.translate_abstract(t)
    assert t["abstract_en"] == PLACEHOLDER
