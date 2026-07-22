# -*- coding: utf-8 -*-
from src import synthesizer

TOPIC = {"title": "题目X", "author": "张三", "background": "背景"}
CARDS = [{"title": "文A", "topic": "tA", "source": "a.pdf"}]
CHS = [{"title": "系统设计", "kind": "core", "cards": [0]},
       {"title": "总结与展望", "kind": "conclusion", "cards": []}]


def test_write_points_batch_maps_titles(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "系统设计": ["先画总体架构图", "说明模块划分依据，可参考[1]"],
        "总结与展望": ["概括三项工作"]})
    points = synthesizer.write_points(CHS, TOPIC, CARDS)
    assert points["系统设计"][0] == "先画总体架构图"
    assert len(points["总结与展望"]) == 1


def test_write_points_failure_returns_empty(monkeypatch, capsys):
    def boom(system, user):
        raise RuntimeError("挂了")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    assert synthesizer.write_points(CHS, TOPIC, CARDS) == {}
    assert "写作要点" in capsys.readouterr().out


def test_write_points_skips_when_no_chapters():
    # 不该发起任何调用（无 monkeypatch 也不能炸）
    assert synthesizer.write_points([], TOPIC, CARDS) == {}
