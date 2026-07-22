# -*- coding: utf-8 -*-
from src import synthesizer
from src.organizer import PLACEHOLDER


def _doc(blocks, meta=None, source="input/topic.md"):
    return {"source": source, "type": "md", "blocks": blocks, "meta": meta or {}}


def _h(text, level=1):
    return {"kind": "heading", "level": level, "text": text, "rows": None}


def _p(text):
    return {"kind": "paragraph", "level": 0, "text": text, "rows": None}


def test_parse_topic_title_from_first_heading():
    doc = _doc([_h("基于YOLOv8的课堂行为识别系统"), _p("研究内容：检测举手、低头等行为。")])
    t = synthesizer.parse_topic(doc)
    assert t["title"] == "基于YOLOv8的课堂行为识别系统"
    assert "举手" in t["background"]


def test_parse_topic_title_from_first_paragraph_when_no_heading():
    doc = _doc([_p("基于知识图谱的推荐系统研究"), _p("拟采用图神经网络。")])
    t = synthesizer.parse_topic(doc)
    assert t["title"] == "基于知识图谱的推荐系统研究"


def test_parse_topic_author_from_frontmatter():
    doc = _doc([_h("题目X")], meta={"author": "张三"})
    assert synthesizer.parse_topic(doc)["author"] == "张三"


def test_parse_topic_defaults_placeholder():
    doc = _doc([])
    t = synthesizer.parse_topic(doc)
    assert t["title"] == PLACEHOLDER
    assert t["author"] == PLACEHOLDER
    assert t["background"] == ""
