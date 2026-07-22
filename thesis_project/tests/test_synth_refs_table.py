# -*- coding: utf-8 -*-
from src import synthesizer

CARDS = [{"title": "文A", "authors": ["李四"], "year": "2023", "source": "a.pdf"},
         {"title": "文B", "authors": [], "year": "", "source": "b.pdf"}]


def test_format_references_order_preserved(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "references": ["李四. 文A[J]. 某刊, 2023.",
                       "佚名. 文B[EB/OL]. «请补全»."]})
    refs = synthesizer.format_references(CARDS)
    assert len(refs) == 2
    assert refs[0].startswith("李四")


def test_format_references_count_mismatch_falls_back(monkeypatch, capsys):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "references": ["只有一条"]})     # 数量对不上 -> 不可信，整体降级
    refs = synthesizer.format_references(CARDS)
    assert refs == ["文A. «请补全著录信息»（来源文件：a.pdf）",
                    "文B. «请补全著录信息»（来源文件：b.pdf）"]
    assert "参考文献" in capsys.readouterr().out


def test_format_references_llm_failure_falls_back(monkeypatch):
    def boom(system, user):
        raise RuntimeError("挂了")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    refs = synthesizer.format_references(CARDS)
    assert len(refs) == 2 and "a.pdf" in refs[0]


def test_format_references_empty_cards():
    assert synthesizer.format_references([]) == []
