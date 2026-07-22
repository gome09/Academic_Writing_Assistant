# -*- coding: utf-8 -*-
from src import synthesizer
from src.organizer import PLACEHOLDER
from src.llm_enhancer import AI_MARK
from config.format_spec import REFS_SPEC


def _topic_doc():
    return {"source": "input/topic.md", "type": "md",
            "meta": {"author": "张三"},
            "blocks": [{"kind": "heading", "level": 1,
                        "text": "基于X的Y系统", "rows": None}]}


def _ref_doc(name, text):
    return {"source": f"input/{name}", "type": "pdf", "meta": {},
            "blocks": [{"kind": "paragraph", "level": 0, "text": text,
                        "rows": None}]}


def _fake_chat_json_factory():
    """按 system 提示词分发的全流程假 LLM。"""
    def fake(system, user):
        if "文献调研助手" in system:
            return {"title": "文A", "authors": ["李四"], "year": "2023",
                    "topic": "tA", "method": "mA", "conclusion": "cA",
                    "quotes": ["观点A"]}
        if "论文结构顾问" in system:
            return {"chapters": [
                {"title": "绪论", "kind": "intro", "cards": []},
                {"title": "相关技术综述", "kind": "review", "cards": [0]},
                {"title": "系统设计", "kind": "core", "cards": [0]},
                {"title": "总结与展望", "kind": "conclusion", "cards": []}]}
        if "学术写作助手" in system:
            return {"paras": ["综述正文一[1]。", "综述正文二[1]。"]}
        if "写作教练" in system:
            return {"绪论": ["交代背景"], "系统设计": ["先画架构图，参考[1]"],
                    "总结与展望": ["概括工作"]}
        if "参考文献格式化" in system:
            return {"references": ["李四. 文A[J]. 某刊, 2023."]}
        raise AssertionError("未知提示词: " + system[:20])
    return fake


def test_synthesize_full_pipeline(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", _fake_chat_json_factory())
    monkeypatch.delenv("LLM_VISION_MODEL", raising=False)
    thesis = synthesizer.synthesize(_topic_doc(), [_ref_doc("a.pdf", "正文")])

    # thesis dict 与 organizer 输出同构
    assert set(thesis) >= {"title", "author", "abstract", "abstract_en",
                           "keywords", "keywords_en", "chapters",
                           "auto_skeleton", "references"}
    assert thesis["title"] == "基于X的Y系统"
    assert thesis["author"] == "张三"
    assert thesis["abstract"] == PLACEHOLDER      # 不编造摘要
    assert thesis["auto_skeleton"] is False
    assert thesis["references"] == ["李四. 文A[J]. 某刊, 2023."]

    titles = [c["title"] for c in thesis["chapters"]]
    assert titles == ["绪论", "相关技术综述", "系统设计", "总结与展望"]

    review = thesis["chapters"][1]
    assert REFS_SPEC["review_notice"] in review["paras"][0]   # 醒目提示在首段
    assert review["paras"][1].endswith(AI_MARK)

    core = thesis["chapters"][2]
    assert core["paras"][0] == "【写作要点】"
    assert any("架构图" in p for p in core["paras"])
    assert core["paras"][-1] == PLACEHOLDER       # 核心章正文留给作者
    # 章节节点结构完整（docx_builder 兼容）
    for ch in thesis["chapters"]:
        assert set(ch) >= {"title", "level", "paras", "subs", "tables",
                           "images"}


def test_synthesize_prints_degrade_summary(monkeypatch, capsys):
    def boom(system, user):
        raise RuntimeError("全挂")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    monkeypatch.delenv("LLM_VISION_MODEL", raising=False)
    thesis = synthesizer.synthesize(_topic_doc(), [_ref_doc("a.pdf", "正文")])
    # 全部降级仍产出完整骨架
    assert [c["title"] for c in thesis["chapters"]] == \
        [c["title"] for c in REFS_SPEC["default_outline"]]
    assert len(thesis["references"]) == 1
    out = capsys.readouterr().out
    assert "步降级" in out
