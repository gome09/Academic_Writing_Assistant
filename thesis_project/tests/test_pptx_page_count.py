# -*- coding: utf-8 -*-
"""PPT 总页数检查：不少于下限 15。"""
from src import pptx_builder
from src.organizer import PPT_SPEC


def _build_deck(n_slides):
    """生成 n_slides 的 deck。"""
    slides = []
    for i in range(n_slides):
        kind = "content"
        if i == 0:
            kind = "cover"
        elif i == n_slides - 1:
            kind = "thanks"
        slides.append({"type": kind, "title": f"X{i}",
                       "bullets": ["bullet"]} if kind == "content"
                      else {"type": kind, "title": f"X{i}",
                            "subtitle": "S"})
    return {"title": "T", "slides": slides}


def test_pptx_under_min_warns(tmp_path, capsys):
    n = PPT_SPEC["principle"]["total_slides_min"] - 5
    deck = _build_deck(n)
    pptx_builder.build(deck, str(tmp_path / "p.pptx"))
    out = capsys.readouterr().out
    assert "[警告]" in out


def test_pptx_within_range_no_warn(tmp_path, capsys):
    n = PPT_SPEC["principle"]["total_slides_min"] + 2
    deck = _build_deck(n)
    pptx_builder.build(deck, str(tmp_path / "p.pptx"))
    out = capsys.readouterr().out
    assert "[警告]" not in out
