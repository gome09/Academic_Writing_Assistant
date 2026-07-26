# -*- coding: utf-8 -*-
"""
格式规范配置 —— 本次项目的"硬标准"。

来源（网络整理，通用主流方案，具体以本校教务处模板为准）：
  - 本科毕业论文 Word 规范：页边距/字体/字号/行距/标题层级/参考文献 GB/T 7714
  - 答辩 PPT 规范：16:9、字号层级、每页要点数、10/20/30 原则

单位约定：
  - 字号用「磅(pt)」；页边距用「厘米(cm)」；行距固定值用「磅(pt)」。
  - python-docx 用 Pt / Cm；python-pptx 用 Pt / Inches。
"""

# =============================================================================
#  一、论文级别 WORD 规范（本科毕业论文）
# =============================================================================

WORD_SPEC = {
    # ---- 页面 ----
    "page": {
        "size": "A4",              # 当前仅支持 A4；其他值在生成前报错
        "orientation": "portrait", # portrait | landscape
        "margin_top_cm": 2.5,
        "margin_bottom_cm": 2.5,
        "margin_left_cm": 2.5,     # 含装订线，靠左
        "margin_right_cm": 2.5,
        "gutter_cm": 1.0,          # 装订线
        "header_distance_cm": 1.5,
        "footer_distance_cm": 1.75,
    },

    # ---- 正文 ----
    "body": {
        "font_cn": "宋体",
        "font_en": "Times New Roman",
        "size_pt": 12,             # 小四号 = 12pt
        "line_spacing_pt": 20,     # 固定值 20 磅
        "first_line_indent_chars": 2,
        "alignment": "justify",    # 两端对齐
        "space_before_pt": 0,
        "space_after_pt": 0,
    },

    # ---- 各级标题（套用样式 -> 可自动生成目录）----
    # size_pt: 字号；level: Word 大纲级别
    "headings": {
        1: {  # 章标题
            "font_cn": "黑体", "font_en": "Times New Roman",
            "size_pt": 16,          # 三号
            "bold": True, "alignment": "center",
            "space_before_pt": 24, "space_after_pt": 18,
            "line_spacing": 1.0, "outline_level": 0,   # outline_level 暂未落实
        },
        2: {  # 节标题
            "font_cn": "黑体", "font_en": "Times New Roman",
            "size_pt": 14,          # 四号
            "bold": True, "alignment": "left",
            "space_before_pt": 24, "space_after_pt": 6,
            "line_spacing_pt": 20, "outline_level": 1, # outline_level 暂未落实
        },
        3: {  # 条标题
            "font_cn": "黑体", "font_en": "Times New Roman",
            "size_pt": 13,
            "bold": True, "alignment": "left",
            "space_before_pt": 12, "space_after_pt": 6,
            "line_spacing_pt": 20, "outline_level": 2, # outline_level 暂未落实
        },
    },

    # ---- 摘要 / 关键词 ----
    "abstract": {
        "title_cn": "摘  要",
        "title_size_pt": 18,        # 小二号
        "title_bold": True,
        "title_alignment": "center",
        "title_en": "Abstract",
        "title_en_size_pt": 16,     # 三号
        "keywords_label_cn": "关键词：",
        "keywords_label_en": "Key words: ",
        "keywords_sep": "；",
        "first_line_indent_chars": 0,
        "keywords_min": 3,
        "keywords_max": 5,
    },

    # ---- 目录 ----
    "toc": {
        "title": "目  录",
        "levels": 3,               # 列至三级标题
        "leader": "dot",           # 暂未落实（圆点前导符）
        "page_number_align": "right",  # 暂未落实
    },

    # ---- 图 / 表 ----
    "figure": {
        "caption_position": "below",   # 图题在下
        "caption_align": "center",
        "number_by_chapter": True,     # 图2-1
        "prefix": "图",
    },
    "table": {
        "caption_position": "above",   # 表题在上
        "caption_align": "center",
        "number_by_chapter": True,     # 表2-1
        "prefix": "表",
        "style": "three_line",         # 三线表
    },

    # ---- 页码 ----
    "page_number": {
        "front_matter": "roman",       # 摘要/目录用 罗马数字
        "body": "arabic",              # 正文从引言起阿拉伯数字
        "alignment": "center",
    },

    # ---- 参考文献 ----
    "reference": {
        "standard": "GB/T 7714",   # 暂未落实
        "title": "参考文献",
    },
}


# =============================================================================
#  二、答辩级别 PPT 规范
# =============================================================================

PPT_SPEC = {
    # ---- 画布 ----
    "slide": {
        "ratio": "16:9",               # 暂未落实（实际尺寸取自 width_inch/height_inch）
        "width_inch": 13.333,
        "height_inch": 7.5,
    },

    # ---- 字体 ----
    "font": {
        "family_cn": "微软雅黑",       # 更正式、易把握
        "family_en": "Arial",
        "max_font_kinds": 3,           # 暂未落实
    },

    # ---- 字号层级 ----
    "sizes": {
        "cover_title_pt": 40,          # 封面主标题 36-44
        "section_title_pt": 32,        # 章/节标题 28-32
        "slide_title_pt": 30,          # 页标题
        "body_pt": 24,                 # 正文 24-28
        "caption_pt": 16,
    },

    # ---- 版面约束 ----
    "layout": {
        "max_bullets_per_slide": 6,    # 每页要点 <=6~8
        "max_lines_per_bullet": 2,     # 每条要点最多行数（超出会产生告警）
        "content_max_ratio": 0.66,     # 内容 <= 2/3 页面
        "text_align": "left",          # left | center | right
        "background": "white",         # 暂未落实（白底利于表格/公式/数据）
    },

    # ---- 结构（15~20 页；问题->方法->结果->启示）----
    "structure": [
        {"key": "cover",       "title": "封面",          "min": 1, "max": 1},
        {"key": "outline",     "title": "目录",          "min": 1, "max": 1},
        {"key": "background",  "title": "研究背景与意义",  "min": 1, "max": 2},
        {"key": "method",      "title": "研究方法与过程",  "min": 3, "max": 6},
        {"key": "result",      "title": "研究成果",       "min": 3, "max": 6},
        {"key": "conclusion",  "title": "结论与展望",     "min": 1, "max": 2},
        {"key": "thanks",      "title": "致谢",          "min": 1, "max": 1},
    ],

    # ---- 篇幅与原则 ----
    "principle": {
        "total_slides_min": 15,
        "total_slides_max": 20,
        "talk_minutes": 10,            # 暂未落实
        "rule": "10/20/30",            # 暂未落实（10页/20分钟/字号>=30）
        "narrative": ["问题", "方法", "结果", "启示"],  # 暂未落实
    },

    # ---- 主题配色（简洁、白底 + 一个主色）----
    "theme": {
        "primary_rgb": (0x1F, 0x4E, 0x79),   # 深蓝
        "accent_rgb": (0xC0, 0x50, 0x4D),    # 砖红点缀
        "text_rgb": (0x26, 0x26, 0x26),      # 近黑
        "muted_rgb": (0x80, 0x80, 0x80),     # 灰
    },
}


# =============================================================================
#  规范来源（用于 README / 文档标注）
# =============================================================================

SPEC_SOURCES = [
    ("PaperRight 本科毕业论文格式规范", "http://www.paperright.com/20170720/1387.html"),
    ("南方科技大学图书馆 Word 编排学位论文",
     "https://lib.sustech.edu.cn/_upload/article/files/69/a1/688e975647f6bc3c62b8a6994696/b6a493fa-e68e-49e8-96c2-b82982aad9c4.pdf"),
    ("复旦大学图书馆 学位论文排版",
     "https://library.fudan.edu.cn/_upload/article/files/93/d0/3004ecf649519314a294611bc34c/18df71a5-14ef-4fdc-ab58-d34ad41518e4.pdf"),
    ("少数派 毕业论文格式调整指南", "https://sspai.com/post/66140"),
    ("知乎 本科毕业论文答辩PPT该怎么做", "https://zhuanlan.zhihu.com/p/92131000"),
    ("知乎 如何制作优秀的毕业答辩PPT", "https://zhuanlan.zhihu.com/p/625502268"),
    ("艾思科蓝 论文答辩ppt应该注意什么", "https://www.ais.cn/news/featured/33657"),
    ("清新电源 毕业论文答辩PPT制作技巧", "http://www.sztspi.com/archives/25749.html"),
]


# =============================================================================
#  三、参考资料模式（研究写作辅助）
# =============================================================================

REFS_SPEC = {
    # input/ 下存在这些文件名之一（不区分大小写）即触发参考资料模式
    "topic_filenames": ["topic.md", "topic.txt", "题目.md", "题目.txt"],

    # 大纲生成失败时的降级骨架；kind: intro|review|core|conclusion
    # review 章由 LLM 撰写综述正文；其余章只给写作要点 + 占位符
    "default_outline": [
        {"title": "绪论", "kind": "intro"},
        {"title": "相关理论与技术", "kind": "review"},
        {"title": "国内外研究现状", "kind": "review"},
        {"title": "研究内容与方案设计", "kind": "core"},
        {"title": "系统实现与结果分析", "kind": "core"},
        {"title": "总结与展望", "kind": "conclusion"},
    ],

    # 插在每个综述章开头的醒目提示
    "review_notice": "本章由 AI 基于所给文献生成，请逐条核对原文后改写为自己的表述。",

    # 无法按语义匹配到章节的表格/截图统一挂到这一章
    "materials_chapter": "素材附录（整理用，定稿前删除）",
}
