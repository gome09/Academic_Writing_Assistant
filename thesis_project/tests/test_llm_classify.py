# -*- coding: utf-8 -*-
from src import llm_enhancer
from src.organizer import _build_deck


def _ch(title, paras=("内容一二三。",)):
    # 注意：要点须 >=4 字，否则被 _to_bullets 的短句过滤丢弃
    return {"title": title, "level": 1, "paras": list(paras), "subs": []}


def test_build_deck_accepts_custom_classifier():
    deck_meta = {"title": "T", "author": "A"}
    deck = _build_deck(deck_meta, [_ch("某个古怪标题")],
                       classify_fn=lambda t: "result")
    slides = deck["slides"]
    sec_idx = next(i for i, s in enumerate(slides)
                   if s["type"] == "section" and s["title"] == "研究成果")
    assert slides[sec_idx + 1]["type"] == "content"
    assert slides[sec_idx + 1]["title"] == "某个古怪标题"


def test_build_deck_accepts_custom_bullets_fn():
    deck = _build_deck({"title": "T", "author": "A"}, [_ch("绪论")],
                       bullets_fn=lambda paras: ["自定义要点"])
    contents = [s for s in deck["slides"] if s["type"] == "content"]
    assert contents[0]["bullets"] == ["自定义要点"]


def test_build_deck_default_unchanged():
    deck = _build_deck({"title": "T", "author": "A"}, [_ch("绪论")])
    contents = [s for s in deck["slides"] if s["type"] == "content"]
    assert contents[0]["bullets"] == ["内容一二三"]


def test_classify_chapters_filters_invalid(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "绪论": "background", "怪章": "nonsense", "实验": "result"})
    got = llm_enhancer.classify_chapters(["绪论", "怪章", "实验"])
    assert got == {"绪论": "background", "实验": "result"}
