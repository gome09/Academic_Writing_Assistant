# -*- coding: utf-8 -*-
import docx
from src import readers


def _save(tmp_path, paragraphs, name="t.docx"):
    d = docx.Document()
    for p in paragraphs:
        d.add_paragraph(p)   # Normal 样式，模拟"手工排版"文档
    path = str(tmp_path / name)
    d.save(path)
    return path


def test_manual_numbered_headings_promoted(tmp_path):
    path = _save(tmp_path, ["1 绪论",
                            "这是一段足够长的正文内容，描述研究背景与意义。",
                            "1.1 研究背景",
                            "另一段正文。"])
    doc = readers.read_docx(path)
    kinds = [(b["kind"], b["level"]) for b in doc["blocks"]]
    assert ("heading", 1) in kinds
    assert ("heading", 2) in kinds


def test_chapter_word_heading_promoted(tmp_path):
    path = _save(tmp_path, ["第一章 绪论", "正文。"])
    doc = readers.read_docx(path)
    assert doc["blocks"][0]["kind"] == "heading"
    assert doc["blocks"][0]["level"] == 1


def test_year_line_stays_paragraph(tmp_path):
    path = _save(tmp_path, ["2023 年度研究进展"])
    doc = readers.read_docx(path)
    assert doc["blocks"][0]["kind"] == "paragraph"


def test_chapter_word_runon_stays_paragraph(tmp_path):
    path = _save(tmp_path, ["第一章正文内容如下"])
    doc = readers.read_docx(path)
    assert doc["blocks"][0]["kind"] == "paragraph"


def test_long_numbered_sentence_stays_paragraph(tmp_path):
    path = _save(tmp_path,
                 ["1. 本文首先分析了现有方法的不足，然后提出了改进方案，最后进行了实验验证。"])
    doc = readers.read_docx(path)
    assert doc["blocks"][0]["kind"] == "paragraph"


def test_pdf_heading_regex_rejects_year():
    assert readers._PDF_HEADING.match("2023 年国内研究综述") is None
    assert readers._PDF_HEADING.match("2.3 实验设计") is not None
