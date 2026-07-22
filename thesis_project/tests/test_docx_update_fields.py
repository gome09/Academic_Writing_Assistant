# -*- coding: utf-8 -*-
import docx
from docx.oxml.ns import qn

from src import docx_builder


def _min_thesis():
    return {"title": "题目", "author": "作者",
            "abstract": "本文摘要。", "abstract_en": "EN abstract.",
            "keywords": ["a", "b", "c"], "keywords_en": ["a", "b", "c"],
            "chapters": [{"title": "绪论", "level": 1, "paras": ["正文。"],
                          "subs": []}],
            "auto_skeleton": False, "references": ["某文献[J]. 2024."]}


def test_update_fields_flag_present(tmp_path):
    out = str(tmp_path / "o.docx")
    docx_builder.build(_min_thesis(), out)
    d = docx.Document(out)
    el = d.settings.element.find(qn("w:updateFields"))
    assert el is not None
    assert el.get(qn("w:val")) == "true"
