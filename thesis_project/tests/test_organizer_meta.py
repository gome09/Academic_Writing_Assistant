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


def test_english_abstract_inline_body_not_truncated():
    """Bug A：字符类误写 [：: \\s] 导致英文摘要正文开头的 's' 被吞。"""
    docs = [doc([p("Abstract: significant findings about X")])]
    thesis, _ = organize(docs)
    assert thesis["abstract"].startswith("significant")
    assert thesis["abstract"] == "significant findings about X"


def test_abstract_heading_with_single_space_recognized():
    """Bug B：'摘 要'（单空格）标题应能识别，摘要取下一段。"""
    docs = [doc([h(1, "摘 要"), p("本文研究了基于深度学习的垃圾分类方法。")])]
    thesis, _ = organize(docs)
    assert thesis["abstract"] == "本文研究了基于深度学习的垃圾分类方法。"


def test_abstract_heading_with_double_space_recognized():
    """Bug B：'摘  要'（双空格，本项目 docx 输出样式）同样可识别。"""
    docs = [doc([h(1, "摘  要"), p("本文研究了基于深度学习的垃圾分类方法。")])]
    thesis, _ = organize(docs)
    assert thesis["abstract"] == "本文研究了基于深度学习的垃圾分类方法。"
