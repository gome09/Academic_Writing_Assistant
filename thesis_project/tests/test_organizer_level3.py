# -*- coding: utf-8 -*-
from src.organizer import organize
from tests.factories import h, p, doc


def _docs():
    return [doc([h(1, "系统设计"),
                 h(2, "总体架构"),
                 h(3, "前端模块"), p("前端负责采集。"),
                 h(3, "后端模块"), p("后端负责推理。")])]


def test_level3_headings_build_tree():
    thesis, _ = organize(_docs())
    sub = thesis["chapters"][0]["subs"][0]
    assert sub["title"] == "总体架构"
    assert [s3["title"] for s3 in sub["subs"]] == ["前端模块", "后端模块"]
    assert sub["subs"][0]["paras"] == ["前端负责采集。"]
    assert sub["subs"][1]["paras"] == ["后端负责推理。"]


def test_level3_without_parent_sub_promoted():
    """没有二级小节时，三级标题提升为二级。"""
    docs = [doc([h(1, "系统设计"), h(3, "某模块"), p("内容。")])]
    thesis, _ = organize(docs)
    sub = thesis["chapters"][0]["subs"][0]
    assert sub["title"] == "某模块"
    assert sub["paras"] == ["内容。"]


def test_level3_paras_reach_ppt_bullets():
    _, deck = organize(_docs())
    all_bullets = [b for s in deck["slides"] if s["type"] == "content"
                   for b in s["bullets"]]
    assert any("前端负责采集" in b for b in all_bullets)


def test_docx_renders_level3(tmp_path):
    import docx as docx_lib
    from src import docx_builder
    thesis, _ = organize(_docs())
    out = docx_builder.build(thesis, str(tmp_path / "t.docx"))
    d = docx_lib.Document(out)
    h3 = [q.text for q in d.paragraphs if q.style.name == "Heading 3"]
    assert any("前端模块" in t for t in h3)
