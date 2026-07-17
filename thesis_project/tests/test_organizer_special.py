# -*- coding: utf-8 -*-
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
