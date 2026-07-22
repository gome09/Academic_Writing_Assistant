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
import ast
import re

from config.format_spec import PPT_SPEC

PLACEHOLDER = "<请填写>"

_MAX_BULLETS_PER_SLIDE = PPT_SPEC["layout"]["max_bullets_per_slide"]

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
    """剥除自带编号前缀；纯空白标题退化为 PLACEHOLDER。"""
    stripped = _NUM_PREFIX.sub("", title or "").strip()
    if stripped:
        return stripped
    if not (title or "").strip():
        return PLACEHOLDER
    return title.strip()



_REF_TITLE = re.compile(r"^(参\s*考\s*文\s*献|references?)\s*$", re.IGNORECASE)
_ABSTRACT_LABEL = re.compile(r"^(摘[，\s]*要|abstract\b)[：:\s]*", re.I)
_DROP_TITLE = re.compile(r"^(摘\s*要|abstract|目\s*录|致\s*谢|谢\s*辞|附\s*录)",
                         re.IGNORECASE)


def _split_special_chapters(chapters):
    """参考文献章 -> 抽成条目列表；摘要/目录/致谢/附录章 -> 移出正文。

    摘要正文已由 _extract_meta 抽取，这里剔除只是防止正文重复成章。
    """
    body, references = [], []
    for ch in chapters:
        t = ch["title"].strip()
        if _REF_TITLE.match(t):
            paras = list(ch["paras"])
            tables = list(ch.get("tables", []))
            for sub in ch.get("subs", []):
                paras.extend(sub["paras"])
                tables.extend(sub.get("tables", []))
            # 条目也可能排在表格里：展平成行后按同样规则切分
            for rows in tables:
                text = _table_to_text({"rows": rows})
                if text:
                    paras.append(text)
            for para in paras:
                for line in para.splitlines():
                    line = re.sub(r"^\[?\d+\]?[\.、]?\s*", "", line.strip())
                    if line:
                        references.append(line)
        elif _DROP_TITLE.match(t):
            n = len(ch.get("paras", [])) + sum(
                len(s.get("paras", []))
                + sum(len(s3.get("paras", [])) for s3 in s.get("subs", []))
                for s in ch.get("subs", []))
            n_media = _count_media(ch) + sum(
                _count_media(s)
                + sum(_count_media(s3) for s3 in s.get("subs", []))
                for s in ch.get("subs", []))
            if n or n_media:
                extra = f"、{n_media} 个表格/图片" if n_media else ""
                print(f"  [提示] 已从正文剔除章节《{ch['title']}》"
                      f"（{n} 段{extra}）——摘要/目录/致谢/附录不进入正文")
            continue
        else:
            body.append(ch)
    return body, references


def _count_media(node):
    return len(node.get("tables", [])) + len(node.get("images", []))


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
            if isinstance(kw, str):
                s = kw.strip()
                if s.startswith("[") and s.endswith("]"):
                    try:
                        parsed = ast.literal_eval(s)
                        kw = parsed if isinstance(parsed, list) else s.split(",")
                    except (ValueError, SyntaxError):
                        kw = [s]
                else:
                    parts = re.split(r"[,;，;； \s]+", s)
                    kw = [x for x in parts if x]
            keywords = list(kw)

    # 从正文块启发式识别
    all_blocks = [b for d in docs for b in d["blocks"]]
    for i, b in enumerate(all_blocks):
        t = b["text"]
        if (not title and b["kind"] == "heading" and b["level"] <= 1
                and t and not _looks_generic(t)):
            title = t
        # 摘要
        if abstract is None and _ABSTRACT_LABEL.match(t):
            # 摘要正文：同块去掉标签；太短则向后找第一个普通段落
            body = _ABSTRACT_LABEL.sub("", t).strip()
            if len(body) < 10:
                for nb in all_blocks[i + 1:i + 4]:
                    if nb["kind"] == "heading":
                        break  # 进入下一节，摘要缺失
                    if nb["kind"] != "paragraph":
                        continue
                    if re.match(r"^(关键词|关键字|key\s*words)", nb["text"], re.I):
                        continue
                    body = nb["text"]
                    break
            abstract = body or None
        # 关键词
        if not keywords:
            mkw = re.match(r"^(关键词|关键字|key\s*words)[：:\s]*(.+)$", t, re.I)
            if mkw:
                keywords = [k for k in re.split(r"[,;，;； \s]+", mkw.group(2)) if k]

    return {
        "title": title or PLACEHOLDER,
        "author": author or PLACEHOLDER,
        "abstract": abstract or PLACEHOLDER,
        "keywords": keywords or [PLACEHOLDER],
    }


# ---------------------------------------------------------------------------
#  重建章节树
# ---------------------------------------------------------------------------
def _node(title, level):
    """章节树节点：章/节/条统一结构。"""
    return {"title": title, "level": level, "paras": [], "subs": [],
            "tables": [], "images": []}


def _table_has_content(b):
    """表格块至少有一个非空单元格才值得挂载。"""
    rows = b.get("rows")
    return bool(rows) and any(c.strip() for r in rows for c in r)


def _build_chapters(docs):
    """返回 [{title, level, paras:[...], subs:[...], tables:[...], images:[...]}]。"""
    blocks = [b for d in docs for b in d["blocks"]]
    has_heading = any(b["kind"] == "heading" for b in blocks)

    chapters = []
    if has_heading:
        current_ch = None
        current_sub = None
        current_sub3 = None

        def target_node():
            t = current_sub3 or current_sub or current_ch
            if t is not None:
                return t
            # 出现在任何标题之前的内容 -> 前言缓冲
            if not chapters or chapters[0]["title"] != "前言":
                chapters.insert(0, _node("前言", 1))
            return chapters[0]

        for b in blocks:
            if b["kind"] == "heading" and b["level"] <= 1:
                current_ch = _node(_strip_numbering(b["text"]), 1)
                chapters.append(current_ch)
                current_sub = None
                current_sub3 = None
            elif b["kind"] == "heading" and b["level"] == 2:
                if current_ch is None:
                    current_ch = _node(PLACEHOLDER, 1)
                    chapters.append(current_ch)
                current_sub = _node(_strip_numbering(b["text"]), 2)
                current_ch["subs"].append(current_sub)
                current_sub3 = None
            elif b["kind"] == "heading":  # level >= 3
                if current_sub is None:
                    # 无上级小节：提升为二级
                    if current_ch is None:
                        current_ch = _node(PLACEHOLDER, 1)
                        chapters.append(current_ch)
                    current_sub = _node(_strip_numbering(b["text"]), 2)
                    current_ch["subs"].append(current_sub)
                    current_sub3 = None
                else:
                    current_sub3 = _node(_strip_numbering(b["text"]), 3)
                    current_sub["subs"].append(current_sub3)
            elif b["kind"] == "table":
                if _table_has_content(b):
                    target_node()["tables"].append(b["rows"])
            elif b["kind"] == "image":
                target_node()["images"].append(
                    {"data": b.get("data"), "ext": b.get("ext", ".png")})
            else:  # 正文/列表/代码
                text = b["text"]
                if not text:
                    continue
                target_node()["paras"].append(text)
    else:
        # 无标题：套标准骨架并留占位符；全部内容按原文顺序放入
        # "研究内容"一章，保持叙述连贯（LLM 增强层可再做语义分章）。
        for name in DEFAULT_CHAPTERS:
            ch = _node(name, 1)
            ch["paras"].append(PLACEHOLDER)
            chapters.append(ch)
        body = _node("研究内容", 1)
        for b in blocks:
            if b["kind"] == "table" and _table_has_content(b):
                body["tables"].append(b["rows"])
            elif b["kind"] == "image":
                body["images"].append(
                    {"data": b.get("data"), "ext": b.get("ext", ".png")})
            elif b["text"]:
                body["paras"].append(b["text"])
        chapters.insert(3, body)

    return chapters, not has_heading


def _table_to_text(block):
    """表格 -> 多行文本：每行一条、竖线分隔、跳过空单元格与空行。"""
    rows = block.get("rows") or []
    lines = []
    for r in rows:
        cells = [c.strip() for c in r if c and c.strip()]
        if cells:
            lines.append(" | ".join(cells))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
#  提炼要点（给 PPT）
# ---------------------------------------------------------------------------
def _to_bullets(paras, max_bullets=12, max_len=40, title=None):
    bullets = []
    for p in paras:
        # 按句号/分号切，取较短句作要点
        for seg in re.split(r"[。；;\n]", p):
            seg = seg.strip()
            if len(seg) < 4:
                continue
            if len(seg) > max_len:
                cut = max(seg.rfind(c, 0, max_len + 1) for c in "，、,")
                seg = seg[:cut] if cut >= max_len // 2 else seg[:max_len]
                seg += "…"
            bullets.append(seg)
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
    chapters, references = _split_special_chapters(chapters)

    thesis = {
        "title": meta["title"],
        "author": meta["author"],
        "abstract": meta["abstract"],
        "abstract_en": PLACEHOLDER,
        "keywords": meta["keywords"],
        "keywords_en": [PLACEHOLDER],
        "chapters": chapters,
        "auto_skeleton": auto_skeleton,
        "references": references or [
            "示例. GB/T 7714 著录格式. 出版地: 出版者, 年份.  <请替换为真实文献>",
        ],
    }

    deck = _build_deck(meta, chapters)
    return thesis, deck


def _build_deck(meta, chapters, classify_fn=None, bullets_fn=None):
    """把章节映射到 PPT 7 段结构。

    classify_fn(title)->bucket、bullets_fn(paras)->list[str] 为可选注入点，
    默认用规则实现（_classify / _to_bullets），供 LLM 增强层替换。
    """
    classify = classify_fn or _classify
    to_bullets = bullets_fn or _to_bullets
    buckets = {"background": [], "method": [], "result": [], "conclusion": []}
    for ch in chapters:
        key = classify(ch["title"])
        paras = list(ch["paras"])
        for sub in ch.get("subs", []):
            paras.extend(sub["paras"])
            for sub3 in sub.get("subs", []):
                paras.extend(sub3["paras"])
        try:
            bl = to_bullets(paras, title=ch["title"])
        except TypeError:      # 注入的 bullets_fn 不接受 title 时退回旧签名
            bl = to_bullets(paras)
        buckets[key].append({"title": ch["title"], "bullets": bl})

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
            slides.append({"type": "section", "title": label[key],
                           "bucket": key})
            slides.append({"type": "content", "title": label[key],
                           "bullets": ["<待补充要点>"], "bucket": key})
            continue
        slides.append({"type": "section", "title": label[key], "bucket": key})
        for item in group:
            bullets = item["bullets"]
            chunks = [bullets[i:i + _MAX_BULLETS_PER_SLIDE]
                      for i in range(0, len(bullets), _MAX_BULLETS_PER_SLIDE)]
            for pi, chunk in enumerate(chunks):
                title = item["title"] if pi == 0 else f"{item['title']}（续）"
                slides.append({"type": "content", "title": title,
                               "bullets": chunk, "bucket": key})

    slides.append({"type": "thanks", "title": "致  谢",
                   "subtitle": "恳请各位老师批评指正"})
    return {"title": meta["title"], "slides": slides}
