# -*- coding: utf-8 -*-
"""
LLM 增强层 —— 通过 OpenAI 兼容接口提升草案质量（可选，失败自动回退）。

配置（环境变量）：
    LLM_API_KEY    必填；未设置则增强层整体跳过
    LLM_BASE_URL   选填；默认官方 OpenAI；可指向 DeepSeek/通义/Kimi/Ollama 等兼容端点
    LLM_MODEL      选填；默认 gpt-4o-mini
    LLM_TIMEOUT    选填；单步超时秒数，默认 60

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
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

AI_MARK = "<AI生成，请核对>"
_VALID_BUCKETS = {"background", "method", "result", "conclusion"}


def is_available() -> bool:
    return bool(os.environ.get("LLM_API_KEY"))


def _client():
    from openai import OpenAI
    timeout = float(os.environ.get("LLM_TIMEOUT", "60"))
    return OpenAI(api_key=os.environ["LLM_API_KEY"],
                  base_url=os.environ.get("LLM_BASE_URL") or None,
                  timeout=timeout, max_retries=1)


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


# ---------------------------------------------------------------------------
#  章节 -> PPT 分区 语义分类
# ---------------------------------------------------------------------------
_CLS_SYS = ("你是答辩PPT助手。把论文章节标题分类到四个部分之一："
            "background（背景/绪论/综述）、method（方法/设计/理论/需求）、"
            "result（实现/实验/测试/结果）、conclusion（总结/结论/展望）。")


def classify_chapters(titles) -> dict:
    """返回 {标题: bucket}，仅保留合法 bucket；标题字典与查询都做归一化。

    归一化规则：NFKC + 去前后空白 + 去全/半角空格。
    "未命中的标题由调用方退化为规则分类。
    """
    raw = _chat_json(
        _CLS_SYS,
        'Please output JSON {"章节标题": "background|method|result|conclusion", ...}:' + chr(10)
        + json.dumps(list(titles), ensure_ascii=False))
    if not isinstance(raw, dict):
        raise ValueError("分类结果不是 JSON 对象")
    out = {}
    for t, b in raw.items():
        if b not in _VALID_BUCKETS:
            continue
        key = _norm_title(t)
        out[key] = b
    return out


def _norm_title(s) -> str:
    """NFKC + 去空白 + 去全角空格，用于查表键归一化。"""
    if not isinstance(s, str):
        return str(s)
    s = unicodedata.normalize("NFKC", s).strip()
    return s.replace(" ", "").replace(chr(0x3000), "")



# ---------------------------------------------------------------------------
#  PPT 要点提炼 + deck 重建
# ---------------------------------------------------------------------------
_BULLET_SYS = ("你是答辩PPT助手。把给定原文提炼成要点列表：最多6条、每条不超过40字、"
               "用名词短语或短句。只依据原文归纳，不得引入原文没有的内容。")


def _llm_bullets(paras) -> list:
    from src.organizer import PLACEHOLDER
    src = "\n".join(p for p in paras if p and p != PLACEHOLDER)[:2000]
    if not src.strip():
        return ["<待补充要点>"]
    data = _chat_json(_BULLET_SYS,
                      '请以 JSON 输出 {"bullets": ["...", "..."]}：\n\n' + src)
    bullets = [str(b).strip()[:40] for b in data.get("bullets", [])
               if str(b).strip()]
    return bullets[:6] or ["<待补充要点>"]


def _safe_bullets(paras) -> list:
    """LLM 提炼失败时回退规则截取。"""
    from src.organizer import _to_bullets
    try:
        return _llm_bullets(paras)
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] 要点提炼失败，改用规则截取：{e}")
        return _to_bullets(paras)


def rebuild_deck(thesis: dict) -> dict:
    """用 LLM 分类 + LLM 要点重建 PPT 大纲；分类失败整体回退规则分类。"""
    from src.organizer import _build_deck, _classify
    try:
        mapping = classify_chapters([c["title"] for c in thesis["chapters"]])
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] 章节分类失败，改用关键词规则：{e}")
        mapping = {}
    meta = {"title": thesis["title"], "author": thesis["author"]}
    return _build_deck(meta, thesis["chapters"],
                       classify_fn=lambda t: mapping.get(_norm_title(t)) or _classify(t),
                       bullets_fn=_safe_bullets)


# ---------------------------------------------------------------------------
#  无标题文档的语义分章
# ---------------------------------------------------------------------------
_CHAP_SYS = ("你是论文结构助手。把编号段落分配到给定的章节骨架中，"
             "依据语义就近归类；不确定的段落可以不分配。"
             "不得改写段落内容，每个编号至多出现一次。")


def rechapter(thesis: dict) -> None:
    """无标题文档：把"研究内容"中的段落按语义分配到骨架各章。

    章内保持原文顺序；LLM 未分配的段落留在"研究内容"；全部失败则不动。
    """
    from src.organizer import PLACEHOLDER, DEFAULT_CHAPTERS
    if not thesis.get("auto_skeleton"):
        return
    src = next((c for c in thesis["chapters"] if c["title"] == "研究内容"), None)
    if src is None or not src["paras"]:
        return
    paras = src["paras"]
    numbered = "\n".join(f"[{i}] {p[:200]}" for i, p in enumerate(paras))
    data = _chat_json(
        _CHAP_SYS,
        "章节骨架：" + json.dumps(DEFAULT_CHAPTERS, ensure_ascii=False)
        + '\n请以 JSON 输出 {"章节名": [段落编号, ...], ...}：\n\n' + numbered)
    if not isinstance(data, dict):
        return
    used = set()
    assign = {}
    for name in DEFAULT_CHAPTERS:
        idxs = sorted(i for i in data.get(name, [])
                      if isinstance(i, int) and 0 <= i < len(paras)
                      and i not in used)
        used.update(idxs)
        assign[name] = idxs
    if not used:
        return
    new_chapters = []
    for name in DEFAULT_CHAPTERS:
        ps = [paras[i] for i in assign[name]] or [PLACEHOLDER]
        new_chapters.append({"title": name, "level": 1,
                             "paras": ps, "subs": []})
    leftovers = [p for i, p in enumerate(paras) if i not in used]
    if leftovers:
        new_chapters.insert(3, {"title": "研究内容", "level": 1,
                                "paras": leftovers, "subs": []})
    thesis["chapters"] = new_chapters


# ---------------------------------------------------------------------------
#  总入口
# ---------------------------------------------------------------------------
def enhance(thesis: dict, deck: dict, docs: list):
    """依次执行各增强步骤，每步独立容错；返回 (thesis, deck)。"""
    if not is_available():
        print("  [LLM] 未设置 LLM_API_KEY，跳过增强（使用纯规则结果）。")
        return thesis, deck
    steps = [
        ("元信息抽取", lambda: refine_meta(thesis, docs)),
        ("语义分章", lambda: rechapter(thesis)),
        ("英文摘要翻译", lambda: translate_abstract(thesis)),
    ]
    for name, fn in steps:
        try:
            fn()
            print(f"  [LLM] {name} 完成")
        except Exception as e:  # noqa: BLE001
            print(f"  [LLM告警] {name} 失败，保留规则结果：{e}")
    try:
        deck = rebuild_deck(thesis)
        print("  [LLM] PPT 大纲重建完成")
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] PPT 大纲重建失败，保留规则结果：{e}")
    return thesis, deck
