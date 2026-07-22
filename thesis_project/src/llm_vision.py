# -*- coding: utf-8 -*-
"""
视觉理解模块 —— 用 OpenAI 兼容多模态端点理解截图（可选）。

配置（环境变量）：
    LLM_API_KEY       复用文本模型的密钥（必填）
    LLM_VISION_MODEL  多模态模型名（如 gpt-4o / qwen-vl-plus / glm-4v）；
                      未设置则视觉理解整体跳过，截图仅作插图。
    LLM_BASE_URL / LLM_TIMEOUT  与文本调用共用。

原则：
  - _chat_vision 是本模块唯一网络出口，测试打桩即可。
  - 失败抛异常，由调用方（synthesizer）降级为"仅插图"。
"""
from __future__ import annotations
import base64
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.llm_enhancer import _parse_json

_MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
         ".bmp": "image/bmp", ".webp": "image/webp"}

_VISION_PROMPT = (
    "这是论文写作素材截图。请以 JSON 输出 "
    '{"caption": "适合作论文图题的一句话", "summary": "图中信息摘要，100字以内"}。'
    "只描述图中真实可见的内容，不得编造。只输出 JSON。")


def is_vision_available() -> bool:
    return bool(os.environ.get("LLM_API_KEY")) and \
        bool(os.environ.get("LLM_VISION_MODEL"))


def _chat_vision(prompt: str, image_b64: str, mime: str) -> str:
    """唯一网络出口：单图 + 文本 -> 原始回复文本。"""
    from openai import OpenAI
    timeout = float(os.environ.get("LLM_TIMEOUT", "60"))
    client = OpenAI(api_key=os.environ["LLM_API_KEY"],
                    base_url=os.environ.get("LLM_BASE_URL") or None,
                    timeout=timeout, max_retries=1)
    resp = client.chat.completions.create(
        model=os.environ["LLM_VISION_MODEL"],
        messages=[{"role": "user", "content": [
            {"type": "text", "text": prompt},
            {"type": "image_url",
             "image_url": {"url": f"data:{mime};base64,{image_b64}"}},
        ]}],
        temperature=0.2)
    return resp.choices[0].message.content or ""


def describe_image(data: bytes, ext: str) -> dict:
    """截图 -> {"caption": 图题建议, "summary": 内容摘要}；失败抛异常。"""
    b64 = base64.b64encode(data).decode("ascii")
    text = _chat_vision(_VISION_PROMPT, b64, _MIME.get(ext, "image/png"))
    obj = _parse_json(text)
    if not isinstance(obj, dict):
        raise ValueError("视觉理解结果不是 JSON 对象")
    return {"caption": str(obj.get("caption", "")).strip(),
            "summary": str(obj.get("summary", "")).strip()}
