# -*- coding: utf-8 -*-
from src import synthesizer
from config.format_spec import REFS_SPEC

TOPIC = {"title": "基于X的Y系统", "author": "张三", "background": "研究背景……"}
CARDS = [{"title": "文A", "topic": "主题A", "source": "a.pdf"},
         {"title": "文B", "topic": "主题B", "source": "b.pdf"}]


def test_build_outline_valid(monkeypatch):
    def fake_chat_json(system, user):
        return {"chapters": [
            {"title": "绪论", "kind": "intro", "cards": []},
            {"title": "相关技术综述", "kind": "review", "cards": [0, 1]},
            {"title": "系统设计", "kind": "core", "cards": [0]},
            {"title": "总结", "kind": "conclusion", "cards": []},
        ]}
    monkeypatch.setattr(synthesizer, "_chat_json", fake_chat_json)
    outline = synthesizer.build_outline(TOPIC, CARDS, [])
    assert [c["kind"] for c in outline] == ["intro", "review", "core",
                                            "conclusion"]
    assert outline[1]["cards"] == [0, 1]


def test_build_outline_drops_bad_kind_and_card_ids(monkeypatch):
    def fake_chat_json(system, user):
        return {"chapters": [
            {"title": "绪论", "kind": "intro", "cards": []},
            {"title": "怪章", "kind": "weird", "cards": []},      # 非法 kind 丢弃
            {"title": "综述", "kind": "review", "cards": [0, 9]},  # 越界编号过滤
            {"title": "设计", "kind": "core", "cards": []},
            {"title": "总结", "kind": "conclusion", "cards": []},
        ]}
    monkeypatch.setattr(synthesizer, "_chat_json", fake_chat_json)
    outline = synthesizer.build_outline(TOPIC, CARDS, [])
    assert all(c["kind"] in synthesizer._VALID_KINDS for c in outline)
    assert outline[1]["cards"] == [0]


def test_build_outline_llm_failure_uses_default_skeleton(monkeypatch, capsys):
    def boom(system, user):
        raise RuntimeError("接口挂了")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    outline = synthesizer.build_outline(TOPIC, CARDS, [])
    assert [c["title"] for c in outline] == \
        [c["title"] for c in REFS_SPEC["default_outline"]]
    # 默认骨架下所有卡都关联到各综述章，素材不至于丢
    review = [c for c in outline if c["kind"] == "review"]
    assert all(c["cards"] == [0, 1] for c in review)
    assert "大纲" in capsys.readouterr().out


def test_build_outline_too_few_chapters_falls_back(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "chapters": [{"title": "只有一章", "kind": "core", "cards": []}]})
    outline = synthesizer.build_outline(TOPIC, CARDS, [])
    assert [c["title"] for c in outline] == \
        [c["title"] for c in REFS_SPEC["default_outline"]]
