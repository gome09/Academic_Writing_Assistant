# -*- coding: utf-8 -*-
from src import synthesizer, llm_vision
from config.format_spec import REFS_SPEC


def _img_doc(name):
    return {"source": f"input/{name}", "type": "image", "meta": {},
            "blocks": [{"kind": "image", "level": 0, "text": "", "rows": None,
                        "data": b"png-bytes", "ext": ".png"}]}


def _xlsx_doc(name):
    return {"source": f"input/{name}", "type": "xlsx", "meta": {},
            "blocks": [{"kind": "table", "level": 0, "text": "",
                        "rows": [["a", "1"]]}]}


def _ch(title, kind="core"):
    n = synthesizer._node(title, 1)
    n["kind"] = kind
    return n


def test_describe_images_skipped_when_unavailable(monkeypatch):
    monkeypatch.delenv("LLM_VISION_MODEL", raising=False)
    assert synthesizer.describe_images([_img_doc("架构图.png")]) == []


def test_describe_images_collects_notes(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("LLM_VISION_MODEL", "qwen-vl-plus")
    monkeypatch.setattr(llm_vision, "_chat_vision",
                        lambda p, b, m: '{"caption": "架构图", "summary": "三层"}')
    notes = synthesizer.describe_images([_img_doc("架构图.png")])
    assert notes == [{"source": "架构图.png", "caption": "架构图",
                      "summary": "三层"}]


def test_describe_images_failure_degrades(monkeypatch, capsys):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("LLM_VISION_MODEL", "qwen-vl-plus")

    def boom(p, b, m):
        raise RuntimeError("端点不支持")
    monkeypatch.setattr(llm_vision, "_chat_vision", boom)
    assert synthesizer.describe_images([_img_doc("图.png")]) == []
    assert "视觉理解" in capsys.readouterr().out


def test_attach_media_matches_by_filename_keyword():
    chapters = [_ch("系统架构设计"), _ch("总结", "conclusion")]
    synthesizer.attach_media(chapters, [_img_doc("架构.png")], [])
    assert len(chapters[0]["images"]) == 1        # 文件名"架构"命中章标题
    assert len(chapters) == 2                     # 命中即不建素材附录章


def test_attach_media_unmatched_goes_to_materials_chapter():
    chapters = [_ch("系统设计")]
    synthesizer.attach_media(chapters, [_xlsx_doc("问卷统计.xlsx")], [])
    assert chapters[-1]["title"] == REFS_SPEC["materials_chapter"]
    assert chapters[-1]["tables"] == [[["a", "1"]]]


def test_attach_media_vision_note_becomes_para():
    chapters = [_ch("系统架构设计")]
    notes = [{"source": "架构.png", "caption": "总体架构",
              "summary": "三层结构"}]
    synthesizer.attach_media(chapters, [_img_doc("架构.png")], notes)
    joined = "\n".join(chapters[0]["paras"])
    assert "总体架构" in joined and "三层结构" in joined


def test_attach_media_no_media_no_extra_chapter():
    chapters = [_ch("系统设计")]
    synthesizer.attach_media(chapters, [], [])
    assert len(chapters) == 1
