# -*- coding: utf-8 -*-
"""
参考资料综合器 —— 参考资料模式下 organizer 的平级替代品。

输入：题目 Document + 参考资料 Document 列表
输出：与 organizer.organize() 完全同构的 thesis dict（docx_builder 直接消费）

管道（总体架构见 docs/ARCHITECTURE.md）：
  ① make_cards       逐篇文献 -> 摘要卡（逐篇容错）
  ② build_outline    题目 + 摘要卡 -> 章节大纲（失败退 REFS_SPEC 默认骨架）
  ③ write_review     综述章正文（带 [n] 引用与 AI 标记；失败留素材+占位）
  ④ write_points     核心章写作要点（批量；失败留占位）
  ⑤ references.py     本地元数据 -> 确定性 GB/T 7714 条目
  ⑥ attach_media     xlsx/csv/截图挂载（纯规则）

原则：
  - LLM 生成综述与要点，不代写核心研究章节正文（学术诚信边界）。
  - 综述正文和视觉摘要带 AI_MARK；写作要点不统一附加标记，仍需人工核对。
  - 任何步骤降级都记入 _degraded 并汇总打印。
  - 网络出口复用 llm_enhancer._chat_json / llm_vision._chat_vision，
    测试打桩 synthesizer._chat_json / llm_vision._chat_vision 即可。
"""
from __future__ import annotations
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.format_spec import REFS_SPEC
from src.llm_enhancer import AI_MARK, _chat_json
from src.organizer import PLACEHOLDER, _node
from src.references import (entry_from_card, extract_local_metadata,
                            format_gbt, lookup_crossref, validate_citations)

_VALID_KINDS = {"intro", "review", "core", "conclusion"}
_MEDIA_TYPES = ("xlsx", "csv", "image")

# 本次 synthesize() 调用中发生的降级记录（步骤名列表）
_degraded = []


def _note_degrade(step: str, err) -> None:
    _degraded.append(step)
    print(f"  [LLM告警] {step}失败，已降级：{err}")


def _doc_text(doc: dict, limit: int = 6000) -> str:
    """Document -> 纯文本（标题与段落顺序拼接，表格拍平），截断到 limit。"""
    parts = []
    for b in doc["blocks"]:
        if b.get("text"):
            parts.append(b["text"])
        elif b.get("kind") == "table" and b.get("rows"):
            parts.append("\n".join(
                " | ".join(c for c in r if c and c.strip())
                for r in b["rows"] if any(c.strip() for c in r if c)))
    return "\n".join(parts)[:limit]


# ---------------------------------------------------------------------------
#  题目解析
# ---------------------------------------------------------------------------
def parse_topic(topic_doc: dict) -> dict:
    """题目文件 -> {"title", "author", "background"}。

    title：meta.title > 首个一级标题 > 首个非空段落首行 > 占位符。
    background：全文文本（喂给大纲/综述做背景），截断 4000 字符。
    """
    meta = topic_doc.get("meta") or {}
    title = (meta.get("title") or "").strip()
    if not title:
        for b in topic_doc["blocks"]:
            if b["kind"] == "heading" and b["level"] <= 1 and b["text"]:
                title = b["text"]
                break
    if not title:
        for b in topic_doc["blocks"]:
            if b["kind"] == "paragraph" and b["text"]:
                title = b["text"].splitlines()[0].strip()
                break
    return {"title": title or PLACEHOLDER,
            "author": (meta.get("author") or "").strip() or PLACEHOLDER,
            "background": _doc_text(topic_doc, limit=4000)}


# ---------------------------------------------------------------------------
#  ① 文献摘要卡
# ---------------------------------------------------------------------------
_CARD_SYS = ("你是文献调研助手。阅读给定的单篇文献片段，抽取结构化信息。"
             "只能从原文归纳，禁止编造；无法确定的字段输出空字符串或空列表。")


def make_cards(ref_docs: list) -> list:
    """逐篇 -> 摘要卡；单篇失败退化为文本卡，不影响其它篇。

    卡片字段：title/authors/year/topic/method/conclusion/quotes/source；
    退化卡额外带 fallback_text（原文截断），title 用文件名。
    """
    cards = []
    for d in ref_docs:
        name = os.path.basename(d["source"])
        local = extract_local_metadata(d)
        try:
            data = _chat_json(
                _CARD_SYS,
                '请以 JSON 输出 {"title": "...", "authors": ["..."], '
                '"year": "...", "topic": "一句话主题", "method": "...", '
                '"conclusion": "...", "quotes": ["可直接引用的关键观点"]}：'
                "\n\n" + _doc_text(d))
            if not isinstance(data, dict):
                raise ValueError("摘要卡结果不是 JSON 对象")
            cards.append({
                "title": str(local.get("title") or data.get("title") or name).strip(),
                "authors": local.get("authors") or
                           [str(a).strip() for a in data.get("authors") or []
                            if str(a).strip()],
                "year": str(local.get("year") or data.get("year") or "").strip(),
                "topic": str(data.get("topic") or "").strip(),
                "method": str(data.get("method") or "").strip(),
                "conclusion": str(data.get("conclusion") or "").strip(),
                "quotes": [str(q).strip() for q in data.get("quotes") or []
                           if str(q).strip()],
                "source": name,
                "_local_meta": local,
            })
        except Exception as e:  # noqa: BLE001
            _note_degrade(f"《{name}》摘要卡", e)
            cards.append({"title": name, "authors": [], "year": "",
                          "topic": "", "method": "", "conclusion": "",
                          "quotes": [], "source": name, "_local_meta": local,
                          "fallback_text": _doc_text(d, limit=500)})
    return cards


# ---------------------------------------------------------------------------
#  ② 论文大纲
# ---------------------------------------------------------------------------
_OUTLINE_SYS = (
    "你是论文结构顾问。根据论文题目/研究方向与文献摘要卡，为一篇本科毕业论文"
    "设计章节大纲（4~10章）。每章标注 kind："
    "intro（绪论）、review（文献综述类，可由AI基于文献撰写）、"
    "core（作者必须自己完成的研究/设计/实现/实验章）、conclusion（总结展望）。"
    "cards 列出与该章相关的文献编号（从0开始）。")


def _default_outline(n_cards: int) -> list:
    """降级骨架：全部卡关联到每个综述章，保证素材不丢。"""
    out = []
    for spec_ch in REFS_SPEC["default_outline"]:
        ch = dict(spec_ch)
        ch["cards"] = list(range(n_cards)) if ch["kind"] == "review" else []
        out.append(ch)
    return out


def build_outline(topic: dict, cards: list, img_notes: list) -> list:
    """返回 [{"title", "kind", "cards": [卡编号]}]；失败/不合格退默认骨架。"""
    card_lines = [f"[{i}] {c['title']}：{c.get('topic', '')}"
                  for i, c in enumerate(cards)]
    img_lines = [f"（图片素材）{n['caption']}：{n['summary']}"
                 for n in img_notes if n.get("summary")]
    try:
        data = _chat_json(
            _OUTLINE_SYS,
            '请以 JSON 输出 {"chapters": [{"title": "...", '
            '"kind": "intro|review|core|conclusion", "cards": [0]}]}。\n'
            f"论文题目与研究方向：\n{topic['background'] or topic['title']}\n\n"
            "文献摘要卡：\n" + "\n".join(card_lines + img_lines))
        raw = data.get("chapters") if isinstance(data, dict) else None
        if not isinstance(raw, list):
            raise ValueError("大纲结果缺少 chapters 列表")
        outline = []
        for ch in raw:
            if not isinstance(ch, dict):
                continue
            title = str(ch.get("title") or "").strip()
            kind = str(ch.get("kind") or "").strip()
            if not title or kind not in _VALID_KINDS:
                continue
            ids = sorted({i for i in (ch.get("cards") or [])
                          if isinstance(i, int) and 0 <= i < len(cards)})
            outline.append({"title": title, "kind": kind, "cards": ids})
        if not 4 <= len(outline) <= 10:
            raise ValueError(f"大纲章数不合理：{len(outline)}")
        return outline
    except Exception as e:  # noqa: BLE001
        _note_degrade("大纲生成", e)
        return _default_outline(len(cards))


# ---------------------------------------------------------------------------
#  ③ 综述章撰写
# ---------------------------------------------------------------------------
_REVIEW_SYS = (
    "你是学术写作助手，为本科毕业论文撰写文献综述章节的初稿段落。"
    "只能综合给定摘要卡的信息，禁止编造数据与文献；"
    "引用某张卡片的内容时，在句末用其方括号编号标注（如[1]）。"
    "输出 3~6 个中文段落，每段 100~200 字。")


def _card_brief(idx: int, card: dict) -> str:
    """卡片 -> 提示词行；idx 为全局 1-based 引用编号。"""
    bits = [f"[{idx}] {card['title']}"]
    for k, label in (("topic", "主题"), ("method", "方法"),
                     ("conclusion", "结论")):
        if card.get(k):
            bits.append(f"{label}：{card[k]}")
    if card.get("quotes"):
        bits.append("观点：" + "；".join(card["quotes"][:3]))
    if card.get("fallback_text"):
        bits.append("原文片段：" + card["fallback_text"][:300])
    return "。".join(bits)


def _material_paras(ch: dict, cards: list) -> list:
    """综述降级产物：素材摘录 + 占位符。"""
    paras = ["素材摘录（LLM 综述失败，以下为原始卡片信息）："]
    for i in ch["cards"]:
        paras.append(_card_brief(i + 1, cards[i]))
    paras.append(PLACEHOLDER)
    return paras


def write_review(ch: dict, topic: dict, cards: list, validate=False) -> list:
    """综述章 -> 正文段落列表（每段带 AI_MARK）；失败退素材摘录。"""
    briefs = [_card_brief(i + 1, cards[i]) for i in ch["cards"]]
    try:
        prompt = ('请以 JSON 输出 {"paras": ["段落1", "段落2"]}。\n'
                  f"论文题目：{topic['title']}\n章节标题:{ch['title']}\n\n"
                  "文献摘要卡：\n" + "\n".join(briefs))
        data = _chat_json(_REVIEW_SYS, prompt)
        paras = [str(p).strip() for p in (data.get("paras") or [])
                 if str(p).strip()] if isinstance(data, dict) else []
        if not paras:
            raise ValueError("综述结果没有段落")
        if validate:
            issues = validate_citations(paras, {i + 1 for i in ch["cards"]},
                                        len(cards))
            if issues:
                retry = _chat_json(
                    _REVIEW_SYS,
                    prompt + "\n\n上次输出存在以下引用错误，请修正后重新输出："
                    + "；".join(issues))
                paras = [str(p).strip() for p in (retry.get("paras") or [])
                         if str(p).strip()] if isinstance(retry, dict) else []
                issues = validate_citations(
                    paras, {i + 1 for i in ch["cards"]}, len(cards))
                if not paras or issues:
                    raise ValueError("；".join(issues) or "纠正后没有综述段落")
        return [f"{p} {AI_MARK}" for p in paras]
    except Exception as e:  # noqa: BLE001
        _note_degrade(f"《{ch['title']}》综述", e)
        return _material_paras(ch, cards)


# ---------------------------------------------------------------------------
#  ④ 核心章写作要点（一次批量调用）
# ---------------------------------------------------------------------------
_POINTS_SYS = (
    "你是论文写作教练。为每个章节生成写作要点：作者应当写什么内容、"
    "按什么顺序展开、可参考哪些文献编号（如[1]）。"
    "每章 3~6 条，每条不超过 60 字。只给指引，不代写正文。"
    '只输出 JSON：{"章节标题": ["要点", ...], ...}')


def write_points(chapters: list, topic: dict, cards: list) -> dict:
    """非 review 章 -> {章节标题: [要点...]}；失败返回 {}（调用方留占位）。"""
    if not chapters:
        return {}
    briefs = [_card_brief(i + 1, c) for i, c in enumerate(cards)]
    payload = {"论文题目": topic["title"],
               "章节": [ch["title"] for ch in chapters],
               "文献摘要卡": briefs}
    try:
        data = _chat_json(_POINTS_SYS, json.dumps(payload, ensure_ascii=False))
        if not isinstance(data, dict):
            raise ValueError("写作要点结果不是 JSON 对象")
        out = {}
        for ch in chapters:
            v = data.get(ch["title"])
            if isinstance(v, list):
                pts = [str(x).strip()[:60] for x in v if str(x).strip()]
                if pts:
                    out[ch["title"]] = pts[:6]
        return out
    except Exception as e:  # noqa: BLE001
        _note_degrade("写作要点", e)
        return {}


# ---------------------------------------------------------------------------
#  ⑥ 视觉理解 + 媒体挂载（纯规则挂载）
# ---------------------------------------------------------------------------
def describe_images(media_docs: list) -> list:
    """截图 -> [{"source", "caption", "summary"}]；未配置视觉模型返回 []。"""
    from src import llm_vision
    if not llm_vision.is_vision_available():
        return []
    notes = []
    for d in media_docs:
        if d["type"] != "image":
            continue
        name = os.path.basename(d["source"])
        b = d["blocks"][0]
        try:
            r = llm_vision.describe_image(b["data"], b.get("ext", ".png"))
            notes.append({"source": name, "caption": r["caption"],
                          "summary": r["summary"]})
        except Exception as e:  # noqa: BLE001
            _note_degrade(f"《{name}》视觉理解", e)
    return notes


def _match_chapter(chapters: list, filename: str):
    """文件名（去扩展名）与章标题做双向子串匹配；命中最先出现的章。"""
    stem = os.path.splitext(filename)[0]
    for ch in chapters:
        if ch.get("kind") == "review":
            continue          # 综述章是文献内容，不挂作者自己的素材
        if stem and (stem in ch["title"] or ch["title"] in stem or any(
                len(seg) >= 2 and seg in ch["title"]
                for seg in re.split(r"[\s_\-（）()]+", stem))):
            return ch
    return None


def attach_media(chapters: list, media_docs: list, img_notes: list) -> None:
    """xlsx/csv 表格与截图就地挂到语义匹配章；无匹配挂素材附录章。"""
    notes_by_source = {n["source"]: n for n in img_notes}
    materials = None

    def target_for(filename):
        nonlocal materials
        ch = _match_chapter(chapters, filename)
        if ch is not None:
            return ch
        if materials is None:
            materials = _node(REFS_SPEC["materials_chapter"], 1)
            materials["kind"] = "core"
            chapters.append(materials)
        return materials

    for d in media_docs:
        name = os.path.basename(d["source"])
        ch = target_for(name)
        for b in d["blocks"]:
            if b["kind"] == "table":
                ch["tables"].append(b["rows"])
                ch["paras"].append(f"（下表数据来自：{name}，表题请补全）")
            elif b["kind"] == "image":
                ch["images"].append({"data": b.get("data"),
                                     "ext": b.get("ext", ".png")})
                note = notes_by_source.get(name)
                if note:
                    ch["paras"].append(
                        f"（插图来自：{name}；图题建议：{note['caption']}；"
                        f"内容摘要：{note['summary']} {AI_MARK}）")
                else:
                    ch["paras"].append(f"（插图来自：{name}，图题请补全）")


# ---------------------------------------------------------------------------
#  总入口
# ---------------------------------------------------------------------------
def synthesize(topic_doc: dict, ref_docs: list, lookup_metadata=False,
               cache_path=None) -> dict:
    """参考资料 -> thesis dict（与 organizer.organize 同构，仅 Word 用）。"""
    del _degraded[:]
    topic = parse_topic(topic_doc)
    text_docs = [d for d in ref_docs if d["type"] not in _MEDIA_TYPES]
    media_docs = [d for d in ref_docs if d["type"] in _MEDIA_TYPES]

    print(f"  文献 {len(text_docs)} 篇，数据/截图 {len(media_docs)} 个")
    cards = make_cards(text_docs)
    if lookup_metadata:
        for card in cards:
            try:
                hit = lookup_crossref(card.get("title", ""),
                                      card.get("_local_meta", {}).get("doi", ""),
                                      cache_path)
                if not hit:
                    continue
                if hit.get("author") and not card.get("authors"):
                    card["authors"] = [" ".join(filter(None, (a.get("family"),
                                                               a.get("given"))))
                                       for a in hit["author"]]
                card["journal"] = (hit.get("container-title") or [""])[0]
                card["volume"] = hit.get("volume", "")
                card["issue"] = hit.get("issue", "")
                card["pages"] = hit.get("page", "")
                card["doi"] = hit.get("DOI", "")
                local = card.setdefault("_local_meta", {})
                local.setdefault("title", (hit.get("title") or [""])[0])
                local.setdefault("doi", hit.get("DOI", ""))
                if not local.get("year"):
                    parts = hit.get("published", {}).get("date-parts", [[]])
                    local["year"] = str(parts[0][0]) if parts and parts[0] else ""
            except Exception as e:  # noqa: BLE001
                _note_degrade(f"《{card.get('title', '')}》Crossref 元数据", e)
    img_notes = describe_images(media_docs)
    outline = build_outline(topic, cards, img_notes)

    chapters = []
    for spec_ch in outline:
        ch = _node(spec_ch["title"], 1)
        ch["kind"] = spec_ch["kind"]
        if spec_ch["kind"] == "review":
            ch["paras"].append("【提示】" + REFS_SPEC["review_notice"])
            ch["paras"].extend(write_review(spec_ch, topic, cards,
                                             validate=True))
        chapters.append(ch)

    plain = [spec_ch for spec_ch in outline if spec_ch["kind"] != "review"]
    points = write_points(plain, topic, cards)
    for spec_ch, ch in zip(outline, chapters):
        if spec_ch["kind"] == "review":
            continue
        pts = points.get(ch["title"])
        if pts:
            ch["paras"].append("【写作要点】")
            ch["paras"].extend(f"· {p}" for p in pts)
        quotes = [f"[{i + 1}] {q}" for i in spec_ch["cards"]
                  for q in cards[i].get("quotes", [])[:2]]
        if quotes:
            ch["paras"].append("素材摘录：")
            ch["paras"].extend(quotes)
        ch["paras"].append(PLACEHOLDER)

    attach_media(chapters, media_docs, img_notes)
    reference_entries = [entry_from_card(c, c.get("_local_meta")) for c in cards]
    references = [format_gbt(entry) for entry in reference_entries]

    if _degraded:
        print(f"  [提示] 本次共 {len(_degraded)} 步降级，请检查上方告警。")
    return {
        "title": topic["title"],
        "author": topic["author"],
        "abstract": PLACEHOLDER,
        "abstract_en": PLACEHOLDER,
        "keywords": [PLACEHOLDER],
        "keywords_en": [PLACEHOLDER],
        "chapters": chapters,
        "auto_skeleton": False,
        "references": references,
        "reference_entries": reference_entries,
        "degraded_steps": list(_degraded),
    }
