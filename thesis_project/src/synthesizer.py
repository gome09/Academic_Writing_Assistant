# -*- coding: utf-8 -*-
"""
参考资料综合器 —— 参考资料模式下 organizer 的平级替代品。

输入：题目 Document + 参考资料 Document 列表
输出：与 organizer.organize() 完全同构的 thesis dict（docx_builder 直接消费）

管道（见 docs/superpowers/specs/2026-07-22-refs-to-draft-design.md）：
  ① make_cards       逐篇文献 -> 摘要卡（逐篇容错）
  ② build_outline    题目 + 摘要卡 -> 章节大纲（失败退 REFS_SPEC 默认骨架）
  ③ write_review     综述章正文（带 [n] 引用与 AI 标记；失败留素材+占位）
  ④ write_points     核心章写作要点（批量；失败留占位）
  ⑤ format_references 摘要卡 -> GB/T 7714（失败罗列原始标题）
  ⑥ attach_media     xlsx/csv/截图挂载（纯规则）

原则：
  - LLM 生成综述与要点，不代写核心研究章节正文（学术诚信边界）。
  - 所有 AI 正文带 AI_MARK；任何步骤降级都记入 _degraded 并汇总打印。
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
