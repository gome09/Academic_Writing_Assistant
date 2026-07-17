# -*- coding: utf-8 -*-
from src import llm_enhancer
from src.organizer import PLACEHOLDER


def _thesis():
    return {"title": "某系统研究", "author": "张三",
            "abstract": "摘要。", "abstract_en": PLACEHOLDER,
            "keywords": ["k"], "keywords_en": [PLACEHOLDER],
            "chapters": [
                {"title": "绪论", "level": 1,
                 "paras": ["研究背景很重要。研究意义也很大。"], "subs": []},
                {"title": "某个古怪标题", "level": 1,
                 "paras": ["做了实验，效果不错。"], "subs": []},
            ],
            "references": [], "auto_skeleton": False}


def test_llm_bullets_used(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        lambda s, u: {"bullets": ["要点一", "要点二"]})
    assert llm_enhancer._llm_bullets(["原文段落。"]) == ["要点一", "要点二"]


def test_llm_bullets_empty_paras_placeholder(monkeypatch):
    def boom(s, u):
        raise AssertionError("空内容不应调用 LLM")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    assert llm_enhancer._llm_bullets([]) == ["<待补充要点>"]
    assert llm_enhancer._llm_bullets([PLACEHOLDER]) == ["<待补充要点>"]


def test_safe_bullets_falls_back_on_error(monkeypatch):
    def boom(s, u):
        raise RuntimeError("api down")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    got = llm_enhancer._safe_bullets(["规则截取的句子够长可以入选。"])
    assert got == ["规则截取的句子够长可以入选"]


def test_rebuild_deck_uses_llm_classify_and_bullets(monkeypatch):
    def fake_chat_json(system, user):
        if "分类" in system or "background" in system:
            return {"某个古怪标题": "result"}
        return {"bullets": ["LLM要点"]}
    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    deck = llm_enhancer.rebuild_deck(_thesis())
    slides = deck["slides"]
    sec_idx = next(i for i, s in enumerate(slides)
                   if s["type"] == "section" and s["title"] == "研究成果")
    assert slides[sec_idx + 1]["title"] == "某个古怪标题"
    assert slides[sec_idx + 1]["bullets"] == ["LLM要点"]


def test_rebuild_deck_classify_failure_falls_back(monkeypatch):
    def fake_chat_json(system, user):
        if "background" in system:
            raise RuntimeError("api down")
        return {"bullets": ["LLM要点"]}
    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    deck = llm_enhancer.rebuild_deck(_thesis())
    # 分类失败 -> 规则分类："某个古怪标题" 命不中关键词 -> method
    slides = deck["slides"]
    sec_idx = next(i for i, s in enumerate(slides)
                   if s["type"] == "section" and s["title"] == "研究方法与过程")
    titles_after = [s["title"] for s in slides[sec_idx + 1:sec_idx + 3]]
    assert "某个古怪标题" in titles_after
