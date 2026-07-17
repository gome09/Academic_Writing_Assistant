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
