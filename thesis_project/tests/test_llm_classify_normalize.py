# -*- coding: utf-8 -*-
"""classify_chapters 标题归一化：NFKC + 去空白。"""
import unicodedata

from src import llm_enhancer


def _norm(s):
    s = unicodedata.normalize("NFKC", s).strip()
    return s.replace(" ", "").replace("\u3000", "")


def test_classify_chapters_normalizes_keys(monkeypatch):
    """LLM 返回带前后空白/全角空格的标题，classify_chapters 内部归一化后依然命中。"""

    def fake_chat_json(system, user):
        return {"  实验环境  ": "result", "结论与展望": "conclusion"}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    titles = ["实验环境", "结论与展望"]
    result = llm_enhancer.classify_chapters(titles)
    # 归一化键查询
    assert result.get(_norm("实验环境")) == "result"
    assert result.get(_norm("结论与展望")) == "conclusion"


def test_classify_chapters_invalid_buckets_filtered(monkeypatch):
    """非法 bucket 被过滤。"""

    def fake_chat_json(system, user):
        return {"实验环境": "noise", "结论与展望": "conclusion"}

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    result = llm_enhancer.classify_chapters(["实验环境", "结论与展望"])
    # 「noise」非法桶，整条不进入
    assert "实验环境" not in result and _norm("实验环境") not in result
    # 若归一化键查询包含应仍可命中
    assert result.get(_norm("结论与展望")) == "conclusion"


# ---------------------------------------------------------------------------
#  rebuild_deck：查询侧同样需要归一化（bug 回归测试）
# ---------------------------------------------------------------------------
_SECTION_LABEL = {"background": "研究背景与意义", "method": "研究方法与过程",
                  "result": "研究成果", "conclusion": "结论与展望"}


def _make_thesis(title):
    return {
        "title": "测试论文",
        "author": "张三",
        "chapters": [{"title": title, "level": 1,
                      "paras": ["这是一段用于测试的正文内容。"], "subs": []}],
    }


def _stub_chat_json(mapping):
    """分类请求返回 mapping；要点提炼请求返回固定 bullets。"""

    def fake_chat_json(system, user):
        if "分类" in system:
            return mapping
        return {"bullets": ["要点一"]}

    return fake_chat_json


def _bucket_of(deck, chapter_title):
    """定位 content 页所属分区：返回其前最近一个 section 页的标题。"""
    current = None
    for slide in deck["slides"]:
        if slide["type"] == "section":
            current = slide["title"]
        elif slide["type"] == "content" and slide["title"] == chapter_title:
            return current
    raise AssertionError(f"deck 中找不到章节 {chapter_title!r} 的 content 页")


def test_rebuild_deck_fullwidth_space_title_hits_llm_mapping(monkeypatch):
    """标题含全角空格时，LLM 分类结果仍应命中，而非静默回退规则分类。

    「拓展　应用」含"应用"，规则会分到 result；桩定 LLM 分到 conclusion。
    若查询未归一化则 mapping.get 落空、回退规则 -> result，测试失败。
    """
    title = "拓展　应用"  # 含全角空格 U+3000
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        _stub_chat_json({title: "conclusion"}))
    deck = llm_enhancer.rebuild_deck(_make_thesis(title))
    assert _bucket_of(deck, title) == _SECTION_LABEL["conclusion"]


def test_rebuild_deck_normalized_query_hit(monkeypatch):
    """归一化命中路径：「系统　实现」桩定为 background（规则会给 result）。"""
    title = "系统　实现"
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        _stub_chat_json({title: "background"}))
    deck = llm_enhancer.rebuild_deck(_make_thesis(title))
    assert _bucket_of(deck, title) == _SECTION_LABEL["background"]


def test_rebuild_deck_miss_falls_back_to_rule(monkeypatch):
    """未命中回退路径：LLM 映射不含该标题，退化为规则分类（实验分析 -> result）。"""
    title = "实验分析"
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        _stub_chat_json({"其它章节": "conclusion"}))
    deck = llm_enhancer.rebuild_deck(_make_thesis(title))
    assert _bucket_of(deck, title) == _SECTION_LABEL["result"]
