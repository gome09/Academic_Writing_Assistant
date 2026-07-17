# -*- coding: utf-8 -*-
from src.organizer import organize, PLACEHOLDER, DEFAULT_CHAPTERS
from tests.factories import h, p, doc


def _headingless_docs(n=6):
    return [doc([p(f"第{i}段。") for i in range(1, n + 1)], type_="txt")]


def test_paragraph_order_preserved():
    thesis, _ = organize(_headingless_docs())
    content = [c for c in thesis["chapters"] if c["title"] == "研究内容"]
    assert len(content) == 1
    assert content[0]["paras"] == [f"第{i}段。" for i in range(1, 7)]


def test_skeleton_chapters_have_placeholder():
    thesis, _ = organize(_headingless_docs())
    for ch in thesis["chapters"]:
        if ch["title"] in DEFAULT_CHAPTERS:
            assert ch["paras"] == [PLACEHOLDER]


def test_auto_skeleton_flag_true_for_headingless():
    thesis, _ = organize(_headingless_docs())
    assert thesis["auto_skeleton"] is True


def test_auto_skeleton_flag_false_with_headings():
    thesis, _ = organize([doc([h(1, "绪论"), p("内容。")])])
    assert thesis["auto_skeleton"] is False
