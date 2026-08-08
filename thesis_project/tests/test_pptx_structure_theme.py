# -*- coding: utf-8 -*-
"""T2-1/T2-2/T2-3: PPT 结构配置驱动 + 主题包 + 布局参数化测试。"""
from __future__ import annotations

import copy

from config.format_spec import PPT_SPEC, WORD_SPEC, THEME_PRESETS, _PPT_DEFAULTS
from config.template import apply_template
from src.organizer import _build_deck


# ---------------------------------------------------------------------------
#  T2-3: 布局参数化——配置字段存在且被消费
# ---------------------------------------------------------------------------
def test_layout_params_exist():
    """T2-3：布局参数字段存在于配置中。"""
    layout = PPT_SPEC["layout"]
    assert "table_max_rows" in layout
    assert "table_max_cols" in layout
    assert "chars_per_line_text" in layout
    assert "chars_per_line_media" in layout
    assert layout["table_max_rows"] == 8
    assert layout["table_max_cols"] == 6
    assert layout["chars_per_line_text"] == 44
    assert layout["chars_per_line_media"] == 26


def test_word_layout_params_exist():
    """T2-3：Word 表格字号和图片宽度字段存在。"""
    assert WORD_SPEC["figure"]["width_cm"] == 14
    assert WORD_SPEC["table"]["font_size_pt"] == 10.5


def test_layout_params_yaml_overridable(tmp_path):
    """T2-3：布局参数可通过 YAML 覆盖。"""
    path = tmp_path / "layout.yml"
    path.write_text("ppt:\n  layout:\n    table_max_rows: 5\n    table_max_cols: 4\n",
                    encoding="utf-8")
    apply_template(str(path))
    assert PPT_SPEC["layout"]["table_max_rows"] == 5
    assert PPT_SPEC["layout"]["table_max_cols"] == 4
    # 还原
    PPT_SPEC.clear()
    PPT_SPEC.update(copy.deepcopy(_PPT_DEFAULTS))


def test_principle_talk_minutes_implemented():
    """T2-5：talk_minutes 已落实（不再是'暂未落实'）。"""
    assert PPT_SPEC["principle"]["talk_minutes"] == 10


def _meta(title="测试论文", author="张三"):
    return {"title": title, "author": author, "abstract": "",
            "abstract_en": "", "keywords": [], "keywords_en": []}


def _ch(title, paras=None):
    return {"title": title, "level": 1, "paras": paras or ["测试段落。"],
            "subs": [], "tables": [], "images": [], "blocks": [],
            "section_role": "body"}


# ---------------------------------------------------------------------------
#  T2-1: 结构配置驱动生成
# ---------------------------------------------------------------------------
def test_default_structure_produces_seven_segments():
    """默认结构仍为 7 段（封面+目录+4内容段+致谢），向后兼容。"""
    deck = _build_deck(_meta(), [_ch("绪论"), _ch("系统设计"), _ch("实验结果")])
    types = [s["type"] for s in deck["slides"]]
    assert types[0] == "cover"
    assert types[1] == "outline"
    assert types[-1] == "thanks"
    # 4 个内容段各有至少 1 个 section 页
    section_slides = [s for s in deck["slides"] if s["type"] == "section"]
    assert len(section_slides) == 4


def test_outline_items_match_configured_segments():
    """目录页的条目取自 PPT_SPEC['structure'] 的内容段标题。"""
    deck = _build_deck(_meta(), [])
    outline = deck["slides"][1]
    assert outline["type"] == "outline"
    content_segs = [seg["title"] for seg in PPT_SPEC["structure"]
                    if seg["key"] not in ("cover", "outline", "thanks")]
    assert outline["items"] == content_segs


def test_custom_structure_changes_deck(tmp_path):
    """YAML 自定义结构后产物随之变化（T2-1 核心）。"""
    from config.format_spec import PPT_SPEC
    old_structure = copy.deepcopy(PPT_SPEC["structure"])
    try:
        PPT_SPEC["structure"] = [
            {"key": "cover", "title": "封面", "min": 1, "max": 1},
            {"key": "outline", "title": "目录", "min": 1, "max": 1},
            {"key": "background", "title": "研究背景", "min": 1, "max": 2},
            {"key": "method", "title": "方法", "min": 1, "max": 4},
            {"key": "thanks", "title": "致谢", "min": 1, "max": 1},
        ]
        deck = _build_deck(_meta(), [_ch("绪论"), _ch("方法设计")])
        section_titles = [s["title"] for s in deck["slides"]
                          if s["type"] == "section"]
        assert "研究背景" in section_titles
        assert "方法" in section_titles
        # 移除了 result 和 conclusion，不应出现
        assert "研究成果" not in section_titles
        assert "结论与展望" not in section_titles
        # 目录条目也应反映新结构
        outline = deck["slides"][1]
        assert "研究背景" in outline["items"]
        assert "方法" in outline["items"]
        assert len(outline["items"]) == 2
    finally:
        PPT_SPEC["structure"] = old_structure


def test_unknown_bucket_falls_back_to_first_segment():
    """classify 返回未知 bucket 时归到首个内容段，章节不丢失。"""
    deck = _build_deck(_meta(), [_ch("未知分类的章")],
                       classify_fn=lambda title: "nonexistent")
    # 仍有内容页
    content_slides = [s for s in deck["slides"] if s["type"] == "content"]
    assert len(content_slides) > 0


def test_empty_structure_falls_back_to_default():
    """配置异常（无内容段）时回退默认四段。"""
    from config.format_spec import PPT_SPEC
    old_structure = copy.deepcopy(PPT_SPEC["structure"])
    try:
        PPT_SPEC["structure"] = [
            {"key": "cover", "title": "封面", "min": 1, "max": 1},
            {"key": "thanks", "title": "致谢", "min": 1, "max": 1},
        ]
        deck = _build_deck(_meta(), [])
        section_slides = [s for s in deck["slides"] if s["type"] == "section"]
        assert len(section_slides) == 4  # 回退默认四段
    finally:
        PPT_SPEC["structure"] = old_structure


# ---------------------------------------------------------------------------
#  T2-2: 主题包系统
# ---------------------------------------------------------------------------
def test_theme_presets_has_five_themes():
    """内置至少 5 套主题。"""
    assert len(THEME_PRESETS) >= 5
    for name, preset in THEME_PRESETS.items():
        assert "primary_rgb" in preset
        assert "accent_rgb" in preset
        assert "text_rgb" in preset
        assert "muted_rgb" in preset


def test_default_preset_is_academic_blue():
    """默认主题为 academic_blue。"""
    assert PPT_SPEC["theme"]["preset"] == "academic_blue"
    assert PPT_SPEC["theme"]["primary_rgb"] == (0x1F, 0x4E, 0x79)


def test_switching_preset_changes_colors(tmp_path):
    """YAML 切换 preset 后配色变化。"""
    path = tmp_path / "theme.yml"
    path.write_text("ppt:\n  theme:\n    preset: minimal_gray\n",
                    encoding="utf-8")
    apply_template(str(path))
    assert PPT_SPEC["theme"]["preset"] == "minimal_gray"
    assert tuple(PPT_SPEC["theme"]["primary_rgb"]) == (0x40, 0x40, 0x40)
    # 还原
    apply_template_write_empty(tmp_path)


def apply_template_write_empty(tmp_path):
    """还原 PPT_SPEC 到默认值。"""
    PPT_SPEC.clear()
    PPT_SPEC.update(copy.deepcopy(_PPT_DEFAULTS))


def test_invalid_preset_keeps_existing_colors(tmp_path):
    """未知 preset 名 -> 保持现有配色（优雅降级）。"""
    path = tmp_path / "bad.yml"
    path.write_text("ppt:\n  theme:\n    preset: nonexistent_theme\n",
                    encoding="utf-8")
    apply_template(str(path))
    # preset 字段被设为 nonexistent 但颜色保持默认
    assert tuple(PPT_SPEC["theme"]["primary_rgb"]) == (0x1F, 0x4E, 0x79)
    apply_template_write_empty(tmp_path)


def test_individual_color_override_still_works(tmp_path):
    """直接覆盖颜色字段仍有效（与 preset 共存）。"""
    path = tmp_path / "custom.yml"
    path.write_text(
        "ppt:\n  theme:\n    preset: campus_red\n    primary_rgb: [1, 2, 3]\n",
        encoding="utf-8")
    apply_template(str(path))
    # preset 加载后，显式覆盖的 primary_rgb 应生效
    assert PPT_SPEC["theme"]["preset"] == "campus_red"
    assert tuple(PPT_SPEC["theme"]["primary_rgb"]) == (1, 2, 3)
    # accent 仍来自 preset
    assert tuple(PPT_SPEC["theme"]["accent_rgb"]) == (0xD4, 0xA0, 0x17)
    apply_template_write_empty(tmp_path)


def test_dark_theme_has_light_text():
    """深色主题文字为浅色（可读性保证）。"""
    dark = THEME_PRESETS["dark"]
    # 浅色文字：R+G+B > 384 (即平均 > 128)
    text_sum = sum(dark["text_rgb"])
    assert text_sum > 384, "深色主题文字应为浅色"
