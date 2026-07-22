# -*- coding: utf-8 -*-
from src import synthesizer
from src.llm_enhancer import AI_MARK

TOPIC = {"title": "题目X", "author": "张三", "background": "背景"}
CARDS = [{"title": "文A", "topic": "tA", "method": "mA", "conclusion": "cA",
          "quotes": ["观点A"], "source": "a.pdf", "authors": [], "year": ""},
         {"title": "文B", "topic": "tB", "method": "mB", "conclusion": "cB",
          "quotes": [], "source": "b.pdf", "authors": [], "year": ""}]
CH = {"title": "相关技术综述", "kind": "review", "cards": [0, 1]}


def test_write_review_paras_carry_ai_mark(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "paras": ["目标检测方法分为两类[1]。", "近年趋势是端到端[2]。"]})
    paras = synthesizer.write_review(CH, TOPIC, CARDS)
    assert len(paras) == 2
    assert all(p.endswith(AI_MARK) for p in paras)
    assert "[1]" in paras[0]


def test_write_review_citation_ids_are_one_based_positions(monkeypatch):
    captured = {}

    def fake_chat_json(system, user):
        captured["user"] = user
        return {"paras": ["p"]}
    monkeypatch.setattr(synthesizer, "_chat_json", fake_chat_json)
    synthesizer.write_review({"title": "综述", "kind": "review", "cards": [1]},
                             TOPIC, CARDS)
    # 只喂关联卡；提示词中的引用编号是全局 1-based（文B -> [2]）
    assert "[2] 文B" in captured["user"]
    assert "文A" not in captured["user"]


def test_write_review_failure_returns_material_fallback(monkeypatch, capsys):
    def boom(system, user):
        raise RuntimeError("挂了")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    paras = synthesizer.write_review(CH, TOPIC, CARDS)
    # 降级：素材摘录（卡片要点罗列）+ 占位符，无 AI 正文
    assert paras[-1] == synthesizer.PLACEHOLDER
    assert any("文A" in p for p in paras)
    assert "综述" in capsys.readouterr().out
