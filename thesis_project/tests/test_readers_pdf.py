# -*- coding: utf-8 -*-
from src.readers import _join_lines, _pdf_lines_to_blocks


def test_join_lines_chinese_no_space():
    assert _join_lines(["深度学习在图像", "识别领域取得突破。"]) == \
        "深度学习在图像识别领域取得突破。"


def test_join_lines_english_with_space():
    assert _join_lines(["deep learning", "model"]) == "deep learning model"


def test_heading_detected():
    blocks = []
    _pdf_lines_to_blocks("第一章 绪论\n本文研究了垃圾分类\n问题的解决方案。", blocks)
    assert blocks[0]["kind"] == "heading"
    assert blocks[0]["level"] == 1
    assert blocks[0]["text"] == "第一章 绪论"
    assert blocks[1]["kind"] == "paragraph"
    assert blocks[1]["text"] == "本文研究了垃圾分类问题的解决方案。"


def test_numbered_subheading_level():
    blocks = []
    _pdf_lines_to_blocks("1.1 研究背景\n内容一。", blocks)
    assert blocks[0]["kind"] == "heading"
    assert blocks[0]["level"] == 2


def test_sentence_end_splits_paragraphs():
    blocks = []
    _pdf_lines_to_blocks("第一句话结束。\n第二句话开始，\n然后结束。", blocks)
    paras = [b["text"] for b in blocks if b["kind"] == "paragraph"]
    assert paras == ["第一句话结束。", "第二句话开始，然后结束。"]


def test_long_unpunctuated_line_not_heading():
    blocks = []
    _pdf_lines_to_blocks("2023 年以来国内外研究者在垃圾分类领域开展了大量卓有成效的工作\n并取得进展。", blocks)
    assert blocks[0]["kind"] == "paragraph"
