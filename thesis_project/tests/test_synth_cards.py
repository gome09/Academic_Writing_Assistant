# -*- coding: utf-8 -*-
from src import synthesizer


def _ref(source, text):
    return {"source": source, "type": "pdf", "meta": {},
            "blocks": [{"kind": "paragraph", "level": 0, "text": text,
                        "rows": None}]}


def test_make_cards_one_per_doc(monkeypatch):
    def fake_chat_json(system, user):
        return {"title": "YOLO综述", "authors": ["李四"], "year": "2023",
                "topic": "目标检测", "method": "综述", "conclusion": "有效",
                "quotes": ["单阶段检测器速度快"]}
    monkeypatch.setattr(synthesizer, "_chat_json", fake_chat_json)
    cards = synthesizer.make_cards([_ref("a.pdf", "正文A"), _ref("b.pdf", "正文B")])
    assert len(cards) == 2
    assert cards[0]["title"] == "YOLO综述"
    assert cards[0]["source"] == "a.pdf"
    assert cards[0]["quotes"] == ["单阶段检测器速度快"]


def test_make_cards_failed_doc_falls_back_to_text_card(monkeypatch, capsys):
    calls = []

    def flaky(system, user):
        calls.append(1)
        if len(calls) == 1:
            raise RuntimeError("接口超时")
        return {"title": "B文", "authors": [], "year": "", "topic": "t",
                "method": "", "conclusion": "", "quotes": []}
    monkeypatch.setattr(synthesizer, "_chat_json", flaky)
    cards = synthesizer.make_cards(
        [_ref("坏.pdf", "这篇失败了" * 100), _ref("好.pdf", "ok")])
    assert len(cards) == 2                       # 失败篇不丢
    assert cards[0]["title"] == "坏.pdf"          # 退化卡用文件名当标题
    assert cards[0]["fallback_text"].startswith("这篇失败了")
    assert "摘要卡" in capsys.readouterr().out    # 打了降级告警


def test_make_cards_normalizes_missing_fields(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json",
                        lambda s, u: {"title": "X"})   # 其余字段全缺
    card = synthesizer.make_cards([_ref("x.pdf", "t")])[0]
    assert card["authors"] == [] and card["quotes"] == []
    assert card["year"] == "" and card["method"] == ""
