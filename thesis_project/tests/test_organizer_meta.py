# -*- coding: utf-8 -*-
from src.organizer import organize, PLACEHOLDER
from tests.factories import h, p, doc


def test_generic_chapter_heading_not_title():
    """# 绪论 这类章节名不能被当成论文题目。"""
    docs = [doc([h(1, "绪论"), p("背景内容。")])]
    thesis, _ = organize(docs)
    assert thesis["title"] == PLACEHOLDER


def test_numbered_chapter_heading_not_title():
    docs = [doc([h(1, "第一章 绪论"), p("背景内容。")])]
    thesis, _ = organize(docs)
    assert thesis["title"] == PLACEHOLDER


def test_real_title_heading_recognized():
    docs = [doc([h(1, "基于深度学习的垃圾分类系统设计"), h(1, "绪论"), p("内容。")])]
    thesis, _ = organize(docs)
    assert thesis["title"] == "基于深度学习的垃圾分类系统设计"


def test_meta_title_wins_over_heading():
    docs = [doc([h(1, "基于深度学习的垃圾分类系统设计")],
                meta={"title": "来自frontmatter的题目"})]
    thesis, _ = organize(docs)
    assert thesis["title"] == "来自frontmatter的题目"
