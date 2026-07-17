# -*- coding: utf-8 -*-
import pytest
from src.organizer import organize, _strip_numbering
from tests.factories import h, p, doc


@pytest.mark.parametrize("raw,expect", [
    ("第一章 绪论", "绪论"),
    ("第1章　绪论", "绪论"),
    ("一、绪论", "绪论"),
    ("1.1 研究背景", "研究背景"),
    ("1.1卷积网络", "卷积网络"),
    ("3 系统实现", "系统实现"),
    ("绪论", "绪论"),               # 无前缀不变
    ("2023年发展综述", "2023年发展综述"),  # 年份不是编号
    ("第一章", "第一章"),            # 剥完为空则保留原文
])
def test_strip_numbering(raw, expect):
    assert _strip_numbering(raw) == expect


def test_chapter_and_sub_titles_stripped():
    docs = [doc([h(1, "第一章 绪论"), h(2, "1.1 研究背景"), p("内容。")])]
    thesis, _ = organize(docs)
    ch = thesis["chapters"][0]
    assert ch["title"] == "绪论"
    assert ch["subs"][0]["title"] == "研究背景"
