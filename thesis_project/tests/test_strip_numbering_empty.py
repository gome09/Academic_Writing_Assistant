# -*- coding: utf-8 -*-
from src.organizer import _strip_numbering, organize, PLACEHOLDER
from tests.factories import h, p, doc


def test_strip_numbering_pure_whitespace_returns_placeholder():
    assert _strip_numbering(" ") == PLACEHOLDER
    assert _strip_numbering("   ") == PLACEHOLDER
    assert _strip_numbering(chr(0x3000)) == PLACEHOLDER  # 全角空格


def test_strip_numbering_real_blank_in_doc_loses_chapter_title():
    docs = [doc([h(1, "绪论"), p("内容。"),
                 h(1, "   "), p("更多内容。")])]
    thesis, _ = organize(docs)
    titles = [c["title"] for c in thesis["chapters"]]
    # 第二章标题不应是空白；退化为 PLACEHOLDER
    assert "   " not in titles
    assert PLACEHOLDER in titles or len(titles) == 1
