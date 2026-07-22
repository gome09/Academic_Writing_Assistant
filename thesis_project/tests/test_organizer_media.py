# -*- coding: utf-8 -*-
from src import organizer, llm_enhancer
from src.readers import _block


def _doc(blocks):
    return {"source": "t.txt", "type": "txt", "blocks": blocks, "meta": {}}


def _img_block():
    b = _block("image")
    b["data"] = b"\x89PNGfake"
    b["ext"] = ".png"
    return b


def test_table_attached_to_chapter_not_paras():
    blocks = [_block("heading", "第一章 绪论", level=1),
              _block("paragraph", "正文段落。"),
              _block("table", "", rows=[["指标", "数值"], ["准确率", "94%"]])]
    chapters, _ = organizer._build_chapters([_doc(blocks)])
    ch = chapters[0]
    assert ch["tables"] == [[["指标", "数值"], ["准确率", "94%"]]]
    assert ch["paras"] == ["正文段落。"]          # 表格文本不再混入段落


def test_image_attached_to_sub():
    blocks = [_block("heading", "第一章", level=1),
              _block("heading", "1.1 背景", level=2),
              _img_block()]
    chapters, _ = organizer._build_chapters([_doc(blocks)])
    sub = chapters[0]["subs"][0]
    assert len(sub["images"]) == 1
    assert sub["images"][0]["ext"] == ".png"


def test_media_before_any_heading_goes_to_preface():
    # 图片出现在首个标题之前 -> 挂到"前言"缓冲章
    blocks = [_img_block(), _block("heading", "第一章 绪论", level=1)]
    chapters, _ = organizer._build_chapters([_doc(blocks)])
    assert chapters[0]["title"] == "前言"
    assert len(chapters[0]["images"]) == 1


def test_skeleton_branch_collects_media():
    blocks = [_block("paragraph", "无标题文档的正文。"), _img_block(),
              _block("table", "", rows=[["a", "b"]])]
    chapters, auto = organizer._build_chapters([_doc(blocks)])
    assert auto is True
    body = next(c for c in chapters if c["title"] == "研究内容")
    assert len(body["images"]) == 1 and len(body["tables"]) == 1


def test_rechapter_carries_media(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        lambda s, u: {"绪论": [0]})
    src_ch = {"title": "研究内容", "level": 1,
              "paras": ["第一段。", "第二段。"], "subs": [],
              "tables": [[["a", "b"]]], "images": [{"data": b"x", "ext": ".png"}]}
    thesis = {"auto_skeleton": True, "chapters": [src_ch]}
    llm_enhancer.rechapter(thesis)
    all_tables = [t for c in thesis["chapters"] for t in c.get("tables", [])]
    all_images = [i for c in thesis["chapters"] for i in c.get("images", [])]
    assert len(all_tables) == 1 and len(all_images) == 1


def test_reference_table_chapter_extracted_to_references():
    # 参考文献条目放在表格里：行需展平并按条目切分
    blocks = [_block("heading", "参考文献", level=1),
              _block("table", "", rows=[["[1] 张三. 某研究[J]. 某期刊, 2024."],
                                        ["[2] 李四. 某系统[D]. 某大学, 2023."]])]
    chapters, _ = organizer._build_chapters([_doc(blocks)])
    body, refs = organizer._split_special_chapters(chapters)
    assert all(c["title"] != "参考文献" for c in body)
    assert refs == ["张三. 某研究[J]. 某期刊, 2024.",
                    "李四. 某系统[D]. 某大学, 2023."]


def test_all_empty_table_not_attached():
    blocks = [_block("heading", "第一章 绪论", level=1),
              _block("table", "", rows=[["", "  "], ["", ""]])]
    chapters, _ = organizer._build_chapters([_doc(blocks)])
    assert chapters[0]["tables"] == []


def test_all_empty_table_not_attached_in_skeleton():
    blocks = [_block("paragraph", "正文。"),
              _block("table", "", rows=[["", ""]])]
    chapters, _ = organizer._build_chapters([_doc(blocks)])
    body = next(c for c in chapters if c["title"] == "研究内容")
    assert body["tables"] == []


def test_dropped_chapter_notice_counts_media(capsys):
    chapters = [{"title": "附录", "level": 1, "paras": ["问卷原文。"], "subs": [],
                 "tables": [[["a", "b"]]],
                 "images": [{"data": b"x", "ext": ".png"}]}]
    organizer._split_special_chapters(chapters)
    out = capsys.readouterr().out
    assert "附录" in out and "1 段" in out and "2 个表格/图片" in out
