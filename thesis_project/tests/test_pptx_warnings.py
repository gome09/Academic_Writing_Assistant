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


# ---------------------------------------------------------------------------
#  空 deck：不崩溃 + 文件生成 + 低于下限警告（回归 UnboundLocalError bug）
# ---------------------------------------------------------------------------
def _in_range_deck(cover_count=1):
    """构造总页数在规范区间内的 deck，cover 数量可调（其余为 outline/content/thanks）。"""
    n_total = PPT_SPEC["principle"]["total_slides_min"] + 1
    slides = []
    for _ in range(cover_count):
        slides.append({"type": "cover", "title": "题目", "subtitle": ""})
    slides.append({"type": "outline", "title": "目录", "items": ["一", "二"]})
    while len(slides) < n_total - 1:
        slides.append({"type": "content",
                       "title": f"内容 {len(slides)}",
                       "bullets": ["要点"]})
    slides.append({"type": "thanks", "title": "谢谢", "subtitle": ""})
    return {"title": "T", "slides": slides}


def test_pptx_empty_deck_builds_and_warns(tmp_path, capsys):
    """slides 为空时 build 不应抛异常（曾因循环内赋值 n 触发 UnboundLocalError）。"""
    import os
    deck = {"title": "空", "slides": []}
    out = tmp_path / "empty.pptx"
    pptx_builder.build(deck, str(out))
    assert os.path.getsize(str(out)) > 0
    captured = capsys.readouterr().out
    assert "[警告]" in captured
    assert "低于规范下限" in captured


def test_pptx_missing_cover_warns(tmp_path, capsys):
    deck = _in_range_deck(cover_count=0)
    pptx_builder.build(deck, str(tmp_path / "p.pptx"))
    out = capsys.readouterr().out
    assert "[警告]" in out
    assert "封面" in out
    assert "低于规范下限" in out


def test_pptx_duplicate_cover_warns(tmp_path, capsys):
    deck = _in_range_deck(cover_count=2)
    pptx_builder.build(deck, str(tmp_path / "p.pptx"))
    out = capsys.readouterr().out
    assert "[警告]" in out
    assert "封面" in out
    assert "高于规范上限" in out


def test_pptx_structure_ok_no_warn(tmp_path, capsys):
    deck = _in_range_deck(cover_count=1)
    pptx_builder.build(deck, str(tmp_path / "p.pptx"))
    out = capsys.readouterr().out
    assert "[警告]" not in out
