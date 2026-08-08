# -*- coding: utf-8 -*-
"""T1-5: 孤立文献检测测试。"""
from src.references import check_orphan_references


def test_all_cited_no_issues():
    """所有文献都被引用 -> 无告警。"""
    paras = ["综述文本[1]继续[2]结尾[3]"]
    issues = check_orphan_references(paras, 3)
    assert issues == []


def test_orphan_detected():
    """文献[3]未被引用 -> 告警。"""
    paras = ["综述文本[1]继续[2]"]
    issues = check_orphan_references(paras, 3)
    assert len(issues) == 1
    assert "[3]" in issues[0]
    assert "孤立" in issues[0]


def test_no_citations_all_orphans():
    """正文无任何引用 -> 全部为孤立文献。"""
    paras = ["综述文本无引用编号"]
    issues = check_orphan_references(paras, 3)
    assert len(issues) == 3
    assert any("[1]" in i for i in issues)
    assert any("[2]" in i for i in issues)
    assert any("[3]" in i for i in issues)


def test_zero_refs_no_issues():
    """无文献 -> 空列表。"""
    assert check_orphan_references(["text"], 0) == []
    assert check_orphan_references([], 0) == []


def test_multi_paragraph_citations():
    """多段落分散引用。"""
    paras = ["段落一[1]", "段落二[2,3]", "段落三[1]"]
    issues = check_orphan_references(paras, 4)
    assert len(issues) == 1
    assert "[4]" in issues[0]


def test_empty_paragraphs():
    """空段落列表 -> 全部孤立。"""
    issues = check_orphan_references([], 2)
    assert len(issues) == 2


def test_comma_separated_citations():
    """逗号分隔的多个引用编号。"""
    paras = ["综述[1,2,3]"]
    issues = check_orphan_references(paras, 3)
    assert issues == []
