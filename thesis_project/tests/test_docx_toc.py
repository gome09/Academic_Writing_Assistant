# -*- coding: utf-8 -*-
"""TOC \\o 字段 levels 与 headings 实际层级联动。"""
import zipfile
import re

from src import docx_builder
from config.format_spec import WORD_SPEC as W


def _build_with_levels(tmp_path):
    """构造一份带三级标题的论文（docx_builder 应读到 heading 3 存在）。"""
    thesis = {
        "title": "T", "author": "A",
        "abstract": "AB", "abstract_en": "PLACE",
        "keywords": ["k"], "keywords_en": ["PLACE"],
        "chapters": [
            {"title": "章1", "level": 1, "paras": ["P"], "subs": [
                {"title": "1.1", "level": 2, "paras": ["P"], "subs": [
                    {"title": "1.1.1", "level": 3, "paras": ["P"], "subs": []},
                ]},
            ]},
        ],
        "auto_skeleton": False,
        "references": ["示例."],
    }
    deck = {"title": "T", "slides": [{"type": "cover", "title": "T", "subtitle": ""}]}
    return thesis, deck, str(tmp_path / "t.docx")


def test_toc_levels_follow_headings(tmp_path):
    thesis, deck, out = _build_with_levels(tmp_path)
    docx_builder.build(thesis, out)
    # 默认级别至少 2；若 headings 有 3 级，TOC 应该 >= 3
    expected = min(max(2, len(W["headings"])), 3)
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    m = re.search(r'TOC \\o "1-(\d+)"', xml)
    assert m, "TOC field not found"
    assert int(m.group(1)) == expected
