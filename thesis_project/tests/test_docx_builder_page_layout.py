# -*- coding: utf-8 -*-
"""docx_builder 页面与页眉页脚数值落进结果文档。"""
import zipfile
import re

from src import docx_builder
from config.format_spec import WORD_SPEC as W


def _thesis_minimal():
    return {
        "title": "T", "author": "A",
        "abstract": "AB", "abstract_en": "PLACE",
        "keywords": ["k"], "keywords_en": ["PLACE"],
        "chapters": [{"title": "章1", "level": 1, "paras": ["P"], "subs": []}],
        "auto_skeleton": False, "references": ["示例."],
    }


def test_docx_page_margins_present(tmp_path):
    thesis = _thesis_minimal()
    out = str(tmp_path / "p.docx")
    docx_builder.build(thesis, out)
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    p = W["page"]
    # 1cm = 567 twips。检查 gutter (装订线) 存在
    assert 'w:gutter' in xml
    # 至少包含一份 sectPr 块
    assert '<w:sectPr' in xml


def test_docx_page_number_field_present(tmp_path):
    thesis = _thesis_minimal()
    out = str(tmp_path / "p.docx")
    docx_builder.build(thesis, out)
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        has_footer = any(n.startswith("word/footer") for n in names)
        assert has_footer
        footer_xml = z.read([n for n in names if n.startswith("word/footer")][0]).decode("utf-8")
    assert "PAGE" in footer_xml
