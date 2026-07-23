# -*- coding: utf-8 -*-
import docx

from src import docx_builder, organizer
from tests.factories import doc, h, p, table


def test_ordered_blocks_preserve_paragraph_table_paragraph(tmp_path):
    docs = [doc([h(1, "绪论"), p("表前"), table([["A"], ["1"]]), p("表后")])]
    thesis, _ = organizer.organize(docs)
    blocks = thesis["chapters"][0]["blocks"]
    assert [b["kind"] for b in blocks] == ["paragraph", "table", "paragraph"]

    out = docx_builder.build(thesis, str(tmp_path / "ordered.docx"))
    xml = docx.Document(out).element.body.xml
    assert xml.index("表前") < xml.index("w:tbl") < xml.index("表后")


def test_unheaded_second_file_does_not_attach_to_previous_chapter():
    docs = [doc([h(1, "绪论"), p("正文")]),
            {"source": "notes.txt", "type": "txt", "meta": {},
             "blocks": [p("独立素材")]}]
    chapters, _ = organizer._build_chapters(docs)
    assert chapters[-1]["title"] == "notes"
    assert chapters[-1]["paras"] == ["独立素材"]


def test_organize_retains_appendix_role():
    thesis, _ = organizer.organize([doc([h(1, "附录A 调查问卷"), p("问卷内容")])])
    assert thesis["chapters"][0]["section_role"] == "appendix"

