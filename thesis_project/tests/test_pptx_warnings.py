# -*- coding: utf-8 -*-
"""PPT 总页数 < 规范下限时输出告警（不影响文件生成）。"""
from src import pptx_builder
from src.organizer import PPT_SPEC


def _too_short_deck():
    """构造一份总页数 < PPT_SPEC.principle.total_slides_min 的 deck。"""
    slides = []
    while len(slides) < PPT_SPEC["principle"]["total_slides_min"] - 2:
        slides.append({"type": "cover" if not slides else "content",
                       "title": f"占位 {len(slides)}",
                       "bullets": ["占位"]} if slides else
                      {"type": "cover", "title": "短 PPT", "subtitle": ""})
    return {"title": "短 PPT", "slides": slides}


def test_pptx_under_min_emits_warning(tmp_path, capsys):
    deck = _too_short_deck()
    out = tmp_path / "short.pptx"
    pptx_builder.build(deck, str(out))
    out_xml = capsys.readouterr().out
    assert "[警告]" in out_xml
    assert str(len(deck["slides"])) in out_xml


def test_pptx_builds_file_even_when_too_short(tmp_path):
    deck = _too_short_deck()
    out = tmp_path / "short.pptx"
    pptx_builder.build(deck, str(out))
    import os
    assert os.path.getsize(str(out)) > 0
