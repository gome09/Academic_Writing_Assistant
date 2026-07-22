# -*- coding: utf-8 -*-
import json

from src import llm_enhancer
from src.organizer import PLACEHOLDER


def _thesis():
    def ch(title, para):
        return {"title": title, "level": 1, "paras": [para], "subs": []}
    return {"title": "题目", "author": "作者",
            "chapters": [ch("绪论", "研究背景是重要的现实问题。"),
                         ch("系统设计", "系统分为三个模块，各司其职。"),
                         ch("总结与展望", "本文完成了预期目标。")]}


def test_rebuild_deck_makes_exactly_two_llm_calls(monkeypatch):
    calls = []

    def fake_chat_json(system, user):
        calls.append(system)
        if len(calls) == 1:          # 第一次：章节分类
            return {}
        return {"绪论": ["要点A", "要点B"]}   # 第二次：批量要点

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    deck = llm_enhancer.rebuild_deck(_thesis())

    assert len(calls) == 2            # 3 章也只有 2 次调用（分类 + 批量要点）
    content = next(s for s in deck["slides"]
                   if s["type"] == "content" and s["title"] == "绪论")
    assert content["bullets"] == ["要点A", "要点B"]


def test_rebuild_deck_batch_failure_falls_back_to_rules(monkeypatch):
    def fake_chat_json(system, user):
        raise RuntimeError("接口超时")

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    deck = llm_enhancer.rebuild_deck(_thesis())   # 不应抛异常
    contents = [s for s in deck["slides"] if s["type"] == "content"]
    assert contents and all(s["bullets"] for s in contents)


def test_llm_bullets_batch_normalizes_and_caps(monkeypatch):
    monkeypatch.setattr(
        llm_enhancer, "_chat_json",
        lambda s, u: {"绪 论": ["x" * 60] + [f"要点{i}" for i in range(8)]})
    out = llm_enhancer._llm_bullets_batch({"绪论": ["段落。"]})
    key = llm_enhancer._norm_title("绪论")
    assert key in out
    assert len(out[key]) <= 6                 # 上限 6 条
    assert all(len(b) <= 40 for b in out[key])  # 单条截 40 字


def test_llm_bullets_batch_filters_placeholder_chapters(monkeypatch):
    sent = {}

    def fake_chat_json(system, user):
        sent["user"] = user
        return {"绪论": ["要点A"]}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    llm_enhancer._llm_bullets_batch({
        "绪论": ["研究背景是重要的现实问题。"],
        "相关理论与技术": [PLACEHOLDER],       # 骨架章：仅占位符
    })
    payload = json.loads(sent["user"])
    assert "绪论" in payload                    # 真实章节在 payload 中
    assert "相关理论与技术" not in payload      # 占位符章节不发给 LLM
    assert PLACEHOLDER not in sent["user"]     # 占位符文本不得出现
