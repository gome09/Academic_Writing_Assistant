# -*- coding: utf-8 -*-
from src import llm_enhancer
from src.organizer import organize, PLACEHOLDER, DEFAULT_CHAPTERS
from tests.factories import p, doc


def _skeleton_thesis():
    paras = ["A段背景。", "B段方法。", "C段其它。", "D段展望。"]
    thesis, _ = organize([doc([p(t) for t in paras], type_="txt")])
    return thesis


def test_rechapter_assigns_in_order(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "绪论": [1, 0],            # 故意乱序，实现应按原文顺序排回
        "总结与展望": [3],
    })
    t = _skeleton_thesis()
    llm_enhancer.rechapter(t)
    by_title = {c["title"]: c for c in t["chapters"]}
    assert by_title["绪论"]["paras"] == ["A段背景。", "B段方法。"]
    assert by_title["总结与展望"]["paras"] == ["D段展望。"]
    # 未分配段落留在"研究内容"
    assert by_title["研究内容"]["paras"] == ["C段其它。"]
    # 空章保留占位符
    assert by_title["系统实现"]["paras"] == [PLACEHOLDER]


def test_rechapter_invalid_indices_ignored(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "绪论": [0, 99, -1, "x"],
        "系统实现": [0],           # 与绪论重复 -> 后者忽略
    })
    t = _skeleton_thesis()
    llm_enhancer.rechapter(t)
    by_title = {c["title"]: c for c in t["chapters"]}
    assert by_title["绪论"]["paras"] == ["A段背景。"]
    assert by_title["系统实现"]["paras"] == [PLACEHOLDER]


def test_rechapter_noop_without_flag(monkeypatch):
    def boom(s, u):
        raise AssertionError("非骨架文档不应调用 LLM")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    t = _skeleton_thesis()
    t["auto_skeleton"] = False
    before = [c["title"] for c in t["chapters"]]
    llm_enhancer.rechapter(t)
    assert [c["title"] for c in t["chapters"]] == before


def test_rechapter_noop_when_llm_assigns_nothing(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {})
    t = _skeleton_thesis()
    llm_enhancer.rechapter(t)
    by_title = {c["title"]: c for c in t["chapters"]}
    assert by_title["研究内容"]["paras"] == \
        ["A段背景。", "B段方法。", "C段其它。", "D段展望。"]


def test_rechapter_normalizes_returned_keys(monkeypatch):
    # LLM 返回的键带全角空格与编号前缀，仍应命中骨架章节
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        lambda s, u: {"1. 绪论　": [0], "总结与展望 ": [1]})
    thesis = {"auto_skeleton": True,
              "chapters": [{"title": "研究内容", "level": 1,
                            "paras": ["第一段。", "第二段。"], "subs": []}]}
    llm_enhancer.rechapter(thesis)
    by_title = {c["title"]: c["paras"] for c in thesis["chapters"]}
    assert by_title["绪论"] == ["第一段。"]
    assert by_title["总结与展望"] == ["第二段。"]
