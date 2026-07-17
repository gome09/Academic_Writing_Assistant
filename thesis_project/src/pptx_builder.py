# -*- coding: utf-8 -*-
"""
PPT 生成器 —— 按 PPT_SPEC 生成答辩演示草案 (.pptx)。

已落实的规范：
  - 16:9 画布 (13.333 x 7.5 in)
  - 字号层级：封面 40 / 章节 32 / 页标题 30 / 正文 24
  - 白底 + 主色标题条；左对齐；每页要点 <=6
  - 结构：封面 -> 目录 -> 4 个分节 + 内容页 -> 致谢
  - 页脚含页码与论文题目
  - 中文字体：微软雅黑（含 eastAsia 设置）
"""
from __future__ import annotations
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

from config.format_spec import PPT_SPEC as P

_C = P["theme"]
PRIMARY = RGBColor(*_C["primary_rgb"])
ACCENT = RGBColor(*_C["accent_rgb"])
TEXT = RGBColor(*_C["text_rgb"])
MUTED = RGBColor(*_C["muted_rgb"])
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

CN_FONT = P["font"]["family_cn"]
EN_FONT = P["font"]["family_en"]
SZ = P["sizes"]
W_IN = P["slide"]["width_inch"]
H_IN = P["slide"]["height_inch"]


def _set_font(run, size_pt, color=TEXT, bold=False, cn=CN_FONT):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = EN_FONT
    # eastAsia 中文字体
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", cn)


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # 全空白版式


def _fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def _textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top),
                                  Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    return tf


def _rect(slide, left, top, width, height, color):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top),
                                 Inches(width), Inches(height))
    _fill(shp, color)
    shp.shadow.inherit = False
    return shp


# ---------------------------------------------------------------------------
#  页脚（页码 + 题目）
# ---------------------------------------------------------------------------
def _footer(slide, idx, total, title):
    tf = _textbox(slide, 0.5, H_IN - 0.45, W_IN - 1.0, 0.35)
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title[:30]
    _set_font(r, 10, MUTED)
    tf2 = _textbox(slide, W_IN - 1.5, H_IN - 0.45, 1.0, 0.35)
    pp = tf2.paragraphs[0]; pp.alignment = PP_ALIGN.RIGHT
    rr = pp.add_run(); rr.text = f"{idx} / {total}"
    _set_font(rr, 10, MUTED)


# ---------------------------------------------------------------------------
#  各类幻灯片
# ---------------------------------------------------------------------------
def _slide_cover(prs, s):
    slide = _blank(prs)
    _rect(slide, 0, 0, W_IN, H_IN, WHITE)
    _rect(slide, 0, H_IN * 0.38, W_IN, 0.06, PRIMARY)   # 装饰细条
    tf = _textbox(slide, 1.0, 2.2, W_IN - 2.0, 2.0, MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = s["title"]
    _set_font(r, SZ["cover_title_pt"], PRIMARY, bold=True)
    tf2 = _textbox(slide, 1.0, 4.6, W_IN - 2.0, 1.5)
    p = tf2.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = s.get("subtitle", "")
    _set_font(r, SZ["body_pt"], TEXT)
    p2 = tf2.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r = p2.add_run(); r.text = "本科毕业论文答辩"
    _set_font(r, SZ["caption_pt"], MUTED)
    return slide


def _slide_outline(prs, s):
    slide = _blank(prs)
    _title_bar(slide, s["title"])
    tf = _textbox(slide, 1.2, 1.8, W_IN - 2.4, H_IN - 2.5)
    for i, item in enumerate(s["items"], 1):
        p = tf.paragraphs[0] if i == 1 else tf.add_paragraph()
        p.space_after = Pt(14)
        r = p.add_run(); r.text = f"0{i}   {item}"
        _set_font(r, SZ["section_title_pt"] - 4, TEXT)
    return slide


def _slide_section(prs, s):
    slide = _blank(prs)
    _rect(slide, 0, 0, W_IN, H_IN, PRIMARY)
    tf = _textbox(slide, 1.0, H_IN / 2 - 0.8, W_IN - 2.0, 1.6, MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = s["title"]
    _set_font(r, SZ["section_title_pt"], WHITE, bold=True)
    return slide


def _slide_content(prs, s):
    slide = _blank(prs)
    _title_bar(slide, s["title"])
    tf = _textbox(slide, 1.0, 1.8, W_IN - 2.0, H_IN - 2.6)
    bullets = s["bullets"][:P["layout"]["max_bullets_per_slide"]]
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(12)
        r = p.add_run(); r.text = "▪  " + b
        _set_font(r, SZ["body_pt"], TEXT)
    return slide


def _slide_thanks(prs, s):
    slide = _blank(prs)
    _rect(slide, 0, 0, W_IN, H_IN, WHITE)
    tf = _textbox(slide, 1.0, H_IN / 2 - 1.0, W_IN - 2.0, 2.0, MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = s["title"]
    _set_font(r, SZ["cover_title_pt"], PRIMARY, bold=True)
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r = p2.add_run(); r.text = s.get("subtitle", "")
    _set_font(r, SZ["body_pt"], TEXT)
    return slide


def _title_bar(slide, title):
    """页标题：左侧主色竖条 + 标题文字。"""
    _rect(slide, 0.5, 0.55, 0.12, 0.7, ACCENT)
    tf = _textbox(slide, 0.8, 0.5, W_IN - 1.6, 0.9, MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title
    _set_font(r, SZ["slide_title_pt"], PRIMARY, bold=True)
    # 标题下细分隔线
    _rect(slide, 0.5, 1.45, W_IN - 1.0, 0.02, RGBColor(0xDD, 0xDD, 0xDD))


# ---------------------------------------------------------------------------
#  主构建
# ---------------------------------------------------------------------------
_DISPATCH = {
    "cover": _slide_cover,
    "outline": _slide_outline,
    "section": _slide_section,
    "content": _slide_content,
    "thanks": _slide_thanks,
}


def build(deck, out_path):
    prs = Presentation()
    prs.slide_width = Inches(W_IN)
    prs.slide_height = Inches(H_IN)

    slides = deck["slides"]
    total = len(slides)
    for idx, s in enumerate(slides, 1):
        fn = _DISPATCH.get(s["type"], _slide_content)
        slide = fn(prs, s)
        # 封面/分节/致谢不加页脚
        if s["type"] in ("outline", "content"):
            _footer(slide, idx, total, deck["title"])

    prs.save(out_path)
    return out_path
