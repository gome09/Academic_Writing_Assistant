# -*- coding: utf-8 -*-
from src.organizer import organize, _to_bullets
from tests.factories import h, p, doc


def test_bullets_overflow_creates_continuation_slide():
    sentences = "".join(f"第{i}个要点内容够长可以入选。" for i in range(1, 11))
    docs = [doc([h(1, "系统实现"), p(sentences)])]
    _, deck = organize(docs)
    contents = [s for s in deck["slides"]
                if s["type"] == "content" and s["title"].startswith("系统实现")]
    assert len(contents) == 2
    assert len(contents[0]["bullets"]) == 6
    assert contents[1]["title"] == "系统实现（续）"
    assert len(contents[1]["bullets"]) == 4
    # 无静默丢弃：10 个要点全部保留
    assert sum(len(s["bullets"]) for s in contents) == 10


def test_no_continuation_when_few_bullets():
    docs = [doc([h(1, "系统实现"), p("只有一个要点。")])]
    _, deck = organize(docs)
    contents = [s for s in deck["slides"]
                if s["type"] == "content" and s["title"].startswith("系统实现")]
    assert len(contents) == 1


def test_truncation_cuts_at_comma():
    long = "本系统采用了轻量化设计方案与迁移学习方法，通过知识蒸馏进一步压缩模型体积并保持精度水平"
    bullets = _to_bullets([long + "。"])
    # 在逗号处截断而不是硬切 40 字
    assert bullets[0] == "本系统采用了轻量化设计方案与迁移学习方法…"
