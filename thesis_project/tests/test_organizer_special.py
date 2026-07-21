# -*- coding: utf-8 -*-
from src import organizer
from src.organizer import organize
from tests.factories import h, p, doc


def test_reference_chapter_extracted_to_references():
    docs = [doc([h(1, "总结"), p("总结内容。"),
                 h(1, "参考文献"),
                 p("[1] 张三. 某研究[J]. 某期刊, 2024."),
                 p("[2] 李四. 某系统[D]. 某大学, 2023.")])]
    thesis, _ = organize(docs)
    titles = [c["title"] for c in thesis["chapters"]]
    assert "参考文献" not in titles
    assert thesis["references"] == [
        "张三. 某研究[J]. 某期刊, 2024.",
        "李四. 某系统[D]. 某大学, 2023.",
    ]


def test_abstract_and_thanks_chapters_dropped():
    docs = [doc([h(1, "摘要"), p("摘要正文内容，足够长的一段。"),
                 h(1, "绪论"), p("绪论内容。"),
                 h(1, "致谢"), p("感谢导师。")])]
    thesis, _ = organize(docs)
    titles = [c["title"] for c in thesis["chapters"]]
    assert titles == ["绪论"]
    # 摘要仍被 meta 抽取，没有丢
    assert thesis["abstract"] == "摘要正文内容，足够长的一段。"


def test_default_reference_kept_when_no_ref_chapter():
    docs = [doc([h(1, "绪论"), p("内容。")])]
    thesis, _ = organize(docs)
    assert len(thesis["references"]) == 1
    assert "GB/T 7714" in thesis["references"][0]


def test_dropped_chapter_with_content_prints_notice(capsys):
    chapters = [
        {"title": "绪论", "level": 1, "paras": ["正文。"], "subs": []},
        {"title": "附录", "level": 1, "paras": ["问卷原文。", "代码清单。"], "subs": []},
    ]
    kept, refs = organizer._split_special_chapters(chapters)
    out = capsys.readouterr().out
    assert "附录" in out and "2 段" in out
    assert all(c["title"] != "附录" for c in kept)


def test_dropped_empty_chapter_silent(capsys):
    chapters = [{"title": "目录", "level": 1, "paras": [], "subs": []}]
    organizer._split_special_chapters(chapters)
    assert capsys.readouterr().out == ""


def test_dropped_chapter_counts_nested_sub_paras(capsys):
    chapters = [{"title": "附录", "level": 1, "paras": [], "subs": [
        {"title": "附录A", "level": 2, "paras": ["x"], "subs": [
            {"title": "A.1", "level": 3, "paras": ["y"], "subs": []}]}]}]
    organizer._split_special_chapters(chapters)
    assert "2 段" in capsys.readouterr().out
