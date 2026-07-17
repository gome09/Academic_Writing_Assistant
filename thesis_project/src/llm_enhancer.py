# -*- coding: utf-8 -*-
"""
LLM 增强层 —— 通过 OpenAI 兼容接口提升草案质量（可选，失败自动回退）。

配置（环境变量）：
    LLM_API_KEY    必填；未设置则增强层整体跳过
    LLM_BASE_URL   选填；默认官方 OpenAI；可指向 DeepSeek/通义/Kimi/Ollama 等兼容端点
    LLM_MODEL      选填；默认 gpt-4o-mini

原则：
  - LLM 只做「整理 / 提炼 / 翻译 / 分类」，不扩写正文、不编造事实。
  - AI 生成的摘要类内容带标记 <AI生成，请核对>，与 <请填写> 同一哲学。
  - 任意一步失败只打印告警并保留规则结果，绝不让主流程崩溃。
  - 测试中不真实调用 API：_chat 是唯一网络出口，打桩即可。
"""
from __future__ import annotations
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

AI_MARK = "<AI生成，请核对>"
_VALID_BUCKETS = {"background", "method", "result", "conclusion"}


def is_available() -> bool:
    return bool(os.environ.get("LLM_API_KEY"))


def _client():
    from openai import OpenAI
    return OpenAI(api_key=os.environ["LLM_API_KEY"],
                  base_url=os.environ.get("LLM_BASE_URL") or None)


def _chat(system: str, user: str) -> str:
    """单轮对话，返回原始文本。唯一网络出口，便于测试打桩。"""
    resp = _client().chat.completions.create(
        model=os.environ.get("LLM_MODEL", "gpt-4o-mini"),
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": user}],
        temperature=0.2,
    )
    return resp.choices[0].message.content or ""


def _chat_json(system: str, user: str):
    """要求模型输出 JSON 并解析；容忍代码块包裹与前后废话。"""
    text = _chat(system + "\n只输出 JSON，不要任何其它文字。", user)
    m = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL)
    if m:
        text = m.group(1)
    starts = [i for i in (text.find("{"), text.find("[")) if i >= 0]
    if not starts:
        raise ValueError(f"LLM 未返回 JSON：{text[:80]!r}")
    obj, _ = json.JSONDecoder().raw_decode(text[min(starts):].strip())
    return obj


# ---------------------------------------------------------------------------
#  元信息抽取 / 英文摘要翻译
# ---------------------------------------------------------------------------
_META_SYS = ("你是论文排版助手。根据用户提供的论文草稿开头片段抽取元信息。"
             "只能从原文归纳，禁止编造；无法确定的字段输出空字符串或空列表。")

_EN_SYS = ("你是学术翻译。把给定的中文论文摘要与关键词翻译成规范的学术英文。"
           "忠实原文，不添加原文没有的信息。")


def refine_meta(thesis: dict, docs: list) -> None:
    """就地补全 title/author/abstract/keywords 中仍为占位符的字段。"""
    from src.organizer import PLACEHOLDER
    need = [k for k in ("title", "author", "abstract")
            if thesis[k] == PLACEHOLDER]
    if thesis["keywords"] == [PLACEHOLDER]:
        need.append("keywords")
    if not need:
        return
    head = "\n".join(b["text"] for d in docs
                     for b in d["blocks"][:40] if b["text"])[:3000]
    data = _chat_json(
        _META_SYS,
        '请以 JSON 输出 {"title": "...", "author": "...", '
        '"abstract": "...", "keywords": ["...", "..."]}：\n\n' + head)
    for k in need:
        v = data.get(k)
        if not v:
            continue
        if k == "keywords":
            thesis[k] = [str(x).strip() for x in v if str(x).strip()][:5]
        elif k == "abstract":
            thesis[k] = f"{str(v).strip()} {AI_MARK}"
        else:
            thesis[k] = str(v).strip()


def translate_abstract(thesis: dict) -> None:
    """中文摘要/关键词 -> 英文（就地写入 abstract_en / keywords_en）。"""
    from src.organizer import PLACEHOLDER
    ab = thesis["abstract"].replace(AI_MARK, "").strip()
    if not ab or ab == PLACEHOLDER:
        return
    kws = [k for k in thesis["keywords"] if k != PLACEHOLDER]
    data = _chat_json(
        _EN_SYS,
        '请以 JSON 输出 {"abstract_en": "...", "keywords_en": ["..."]}：\n\n'
        f"摘要：{ab}\n关键词：{'；'.join(kws)}")
    if data.get("abstract_en"):
        thesis["abstract_en"] = f"{str(data['abstract_en']).strip()} {AI_MARK}"
    if data.get("keywords_en"):
        thesis["keywords_en"] = [str(x).strip()
                                 for x in data["keywords_en"]
                                 if str(x).strip()][:5]
