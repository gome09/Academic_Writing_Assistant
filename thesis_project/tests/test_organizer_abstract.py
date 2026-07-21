# -*- coding: utf-8 -*-
from src import organizer
from src.readers import _block


def _doc(blocks):
    return {"source": "t.txt", "type": "txt", "blocks": blocks, "meta": {}}


def test_abstraction_paragraph_not_treated_as_abstract():
    d = _doc([_block("paragraph",
                     "Abstraction is a key concept in computer science.")])
    meta = organizer._extract_meta([d])
    assert meta["abstract"] == organizer.PLACEHOLDER


def test_abstract_label_with_body_same_block():
    d = _doc([_block("paragraph", "摘要：本文研究了基于深度学习的图像识别方法。")])
    meta = organizer._extract_meta([d])
    assert meta["abstract"].startswith("本文研究了")


def test_abstract_next_block_skips_keyword_line():
    d = _doc([
        _block("paragraph", "摘要"),
        _block("paragraph", "关键词：深度学习 图像识别 卷积网络"),
        _block("paragraph", "本文研究了基于深度学习的图像识别方法，并完成了系统实现。"),
    ])
    meta = organizer._extract_meta([d])
    assert meta["abstract"].startswith("本文研究了")
    assert "关键词" not in meta["abstract"]


def test_abstract_next_block_skips_heading():
    d = _doc([
        _block("paragraph", "摘要"),
        _block("heading", "第一章 绪论", level=1),
        _block("paragraph", "正文段落。"),
    ])
    meta = organizer._extract_meta([d])
    # 标题是节边界，摘要缺失时回退占位符
    assert meta["abstract"] == organizer.PLACEHOLDER
