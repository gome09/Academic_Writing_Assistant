# -*- coding: utf-8 -*-
"""
内容整理器 —— 把读取到的若干 Document 整理成两份结构化"草案数据"：

  1) ThesisModel  给 Word 生成器用
  2) SlideDeck     给 PPT 生成器用

策略（草案级别，可人工再润色）：
  - 汇总所有 Document 的 blocks，按标题层级重建章节树。
  - 若源文件几乎没有标题（如纯 txt），则套用本科论文的标准章节骨架，
    并把段落按顺序填入"研究内容"章节，保证产出结构完整。
  - 自动抽取：题目、作者、摘要、关键词（识别不到就留占位符 <请填写>）。
  - PPT 大纲从章节树映射到规范里的 7 段结构，正文要点自动提炼。
"""
from __future__ import annotations
import re

PLACEHOLDER = "<请填写>"

# 本科论文标准章节骨架（识别不到标题时使用）
DEFAULT_CHAPTERS = [
    "绪论",
    "相关理论与技术",
    "需求分析与系统设计",
    "系统实现",
    "测试与结果分析",
    "总结与展望",
]

# 章节标题 -> PPT 结构段 的映射关键词
_SECTION_HINT = {
    "background": ["绪论", "引言", "背景", "意义", "概述", "introduction"],
    "method":     ["方法", "设计", "理论", "技术", "模型", "算法", "需求", "方案", "过程"],
    "result":     ["实现", "实验", "测试", "结果", "分析", "评估", "应用", "成果"],
    "conclusion": ["总结", "结论", "展望", "结束", "未来", "conclusion"],
}

# 常见章节名 / 结构性标题 —— 不能当论文题目
_GENERIC_HEADING = re.compile(
    r"^(第\s*[一二三四五六七八九十百\d]+\s*[章节部分]"
    r"|\d+(\.\d+)*[\s、.．]"
    r"|[一二三四五六七八九十]+[、.．]"
    r"|绪论|引言|前言|导论|概述|摘\s*要|abstract|目\s*录|结论|总结|展望"
    r"|参考文献|致\s*谢|附\s*录|关键词"
    r"|相关(理论|技术|工作)|文献综述|研究(背景|现状|意义|内容|方法)"
    r"|国内外研究现状|需求分析|系统设计|系统实现|实验|测试)",
    re.IGNORECASE,
)


def _looks_generic(title: str) -> bool:
    return bool(_GENERIC_HEADING.match(title.strip()))


# 标题自带的编号前缀（生成时会统一重新编号，避免"第一章 第一章 绪论"）
_NUM_PREFIX = re.compile(
    r"^(第\s*[一二三四五六七八九十百\d]+\s*[章节]\s*[、.．:：]?"
    r"|[一二三四五六七八九十]+\s*[、.．]"
    r"|\d+(\.\d+)*\s*[、.．]?\s+"
    r"|\d+(\.\d+)+\s*)"
)


def _strip_numbering(title: str) -> str:
    stripped = _NUM_PREFIX.sub("", title).strip()
    return stripped or title.strip()


# ---------------------------------------------------------------------------
#  抽取标题 / 作者 / 摘要 / 关键词
# ---------------------------------------------------------------------------
def _extract_meta(docs):
    title = author = abstract = None
    keywords = []

    # 先看各 Document 自带 meta
    for d in docs:
        m = d.get("meta") or {}
        title = title or m.get("title")
        author = author or m.get("author")
        if m.get("keywords"):
            kw = m["keywords"]
            keywords = kw if isinstance(kw, list) else re.split(r"[,，;；]\s*", str(kw))

    # 从正文块启发式识别
    all_blocks = [b for d in docs for b in d["blocks"]]
    for i, b in enumerate(all_blocks):
        t = b["text"]
        if (not title and b["kind"] == "heading" and b["level"] <= 1
                and t and not _looks_generic(t)):
            title = t
        # 摘要
        if abstract is None and re.match(r"^(摘\s*要|abstract)", t, re.I):
            # 摘要正文：同块去掉标签，或取下一段
            body = re.sub(r"^(摘\s*要|abstract)[：:\s]*", "", t, flags=re.I).strip()
            if len(body) < 10 and i + 1 < len(all_blocks):
                body = all_blocks[i + 1]["text"]
            abstract = body
        # 关键词
        if not keywords:
            mkw = re.match(r"^(关键词|关键字|key\s*words)[：:\s]*(.+)$", t, re.I)
            if mkw:
                keywords = [k for k in re.split(r"[,，;；\s]+", mkw.group(2)) if k]

    return {
        "title": title or PLACEHOLDER,
        "author": author or PLACEHOLDER,
        "abstract": abstract or PLACEHOLDER,
        "keywords": keywords or [PLACEHOLDER],
    }


# ---------------------------------------------------------------------------
#  重建章节树
# ---------------------------------------------------------------------------
def _build_chapters(docs):
    """返回 [{title, level, paras:[...], subs:[...]}]，两层足够草案用。"""
    blocks = [b for d in docs for b in d["blocks"]]
    has_heading = any(b["kind"] == "heading" for b in blocks)

    chapters = []
    if has_heading:
        current_ch = None
        current_sub = None
        for b in blocks:
            if b["kind"] == "heading" and b["level"] <= 1:
                current_ch = {"title": _strip_numbering(b["text"]), "level": 1,
                              "paras": [], "subs": []}
                chapters.append(current_ch)
                current_sub = None
            elif b["kind"] == "heading" and b["level"] == 2:
                if current_ch is None:
                    current_ch = {"title": PLACEHOLDER, "level": 1, "paras": [], "subs": []}
                    chapters.append(current_ch)
                current_sub = {"title": _strip_numbering(b["text"]), "level": 2,
                               "paras": []}
                current_ch["subs"].append(current_sub)
            elif b["kind"] in ("heading",):  # level>=3 归入当前段落文本
                target = current_sub or current_ch
                if target is not None:
                    target["paras"].append(b["text"])
            else:  # 正文/列表/代码/表格文本
                text = b["text"] if b["kind"] != "table" else _table_to_text(b)
                if not text:
                    continue
                if current_sub is not None:
                    current_sub["paras"].append(text)
                elif current_ch is not None:
                    current_ch["paras"].append(text)
                else:
                    # 出现在任何标题之前的段落 -> 前言缓冲
                    if not chapters or chapters[0]["title"] != "前言":
                        chapters.insert(0, {"title": "前言", "level": 1,
                                            "paras": [], "subs": []})
                    chapters[0]["paras"].append(text)
    else:
        # 无标题：套标准骨架并留占位符；全部段落按原文顺序放入
        # "研究内容"一章，保持叙述连贯（LLM 增强层可再做语义分章）。
        for name in DEFAULT_CHAPTERS:
            chapters.append({"title": name, "level": 1,
                             "paras": [PLACEHOLDER], "subs": []})
        paras = [b["text"] for b in blocks if b["text"]]
        chapters.insert(3, {"title": "研究内容", "level": 1,
                            "paras": paras, "subs": []})

    return chapters, not has_heading


def _table_to_text(block):
    rows = block.get("rows") or []
    return " | ".join(" ".join(r) for r in rows[:2])


# ---------------------------------------------------------------------------
#  提炼要点（给 PPT）
# ---------------------------------------------------------------------------
def _to_bullets(paras, max_bullets=6, max_len=40):
    bullets = []
    for p in paras:
        # 按句号/分号切，取较短句作要点
        for seg in re.split(r"[。；;\n]", p):
            seg = seg.strip()
            if len(seg) < 4:
                continue
            bullets.append(seg[:max_len] + ("…" if len(seg) > max_len else ""))
            if len(bullets) >= max_bullets:
                return bullets
    return bullets or ["<待补充要点>"]


def _classify(title):
    t = title.lower()
    for key, hints in _SECTION_HINT.items():
        if any(h in t for h in hints):
            return key
    return "method"  # 默认归到方法/过程


# ---------------------------------------------------------------------------
#  对外主函数
# ---------------------------------------------------------------------------
def organize(docs):
    meta = _extract_meta(docs)
    chapters, auto_skeleton = _build_chapters(docs)

    thesis = {
        "title": meta["title"],
        "author": meta["author"],
        "abstract": meta["abstract"],
        "abstract_en": PLACEHOLDER,
        "keywords": meta["keywords"],
        "keywords_en": [PLACEHOLDER],
        "chapters": chapters,
        "auto_skeleton": auto_skeleton,
        "references": [
            "示例. GB/T 7714 著录格式. 出版地: 出版者, 年份.  <请替换为真实文献>",
        ],
    }

    deck = _build_deck(meta, chapters)
    return thesis, deck


def _build_deck(meta, chapters):
    """把章节映射到 PPT 7 段结构。"""
    buckets = {"background": [], "method": [], "result": [], "conclusion": []}
    for ch in chapters:
        key = _classify(ch["title"])
        paras = list(ch["paras"])
        for sub in ch.get("subs", []):
            paras.extend(sub["paras"])
        buckets[key].append({"title": ch["title"], "bullets": _to_bullets(paras)})

    slides = []
    slides.append({"type": "cover", "title": meta["title"],
                   "subtitle": f"答辩人：{meta['author']}"})
    slides.append({"type": "outline", "title": "目录",
                   "items": ["研究背景与意义", "研究方法与过程",
                             "研究成果", "结论与展望"]})

    label = {"background": "研究背景与意义", "method": "研究方法与过程",
             "result": "研究成果", "conclusion": "结论与展望"}
    for key in ["background", "method", "result", "conclusion"]:
        group = buckets[key]
        if not group:
            slides.append({"type": "section", "title": label[key]})
            slides.append({"type": "content", "title": label[key],
                           "bullets": ["<待补充要点>"]})
            continue
        slides.append({"type": "section", "title": label[key]})
        for item in group:
            slides.append({"type": "content",
                           "title": item["title"],
                           "bullets": item["bullets"]})

    slides.append({"type": "thanks", "title": "致  谢",
                   "subtitle": "恳请各位老师批评指正"})
    return {"title": meta["title"], "slides": slides}
