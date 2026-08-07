# -*- coding: utf-8 -*-
"""T1-1/T1-2: 文献类型启发式识别 + 多样式著录格式测试。"""
from __future__ import annotations

from src.references import (
    _detect_type, entry_from_card, format_apa, format_chicago,
    format_gbt, format_mla, format_reference,
)


# ---------------------------------------------------------------------------
#  T1-1: 类型识别
# ---------------------------------------------------------------------------
def test_detect_type_journal_default():
    """无特征字段 -> J（期刊）。"""
    assert _detect_type({"title": "论文"}) == "J"
    assert _detect_type({}) == "J"


def test_detect_type_monograph_by_publisher():
    """有出版社无期刊 -> M（专著）。"""
    assert _detect_type({"publisher": "清华大学出版社"}) == "M"


def test_detect_type_conference_by_meeting():
    """有会议名 -> C（会议论文）。"""
    assert _detect_type({"conference": "ICML 2023"}) == "C"
    assert _detect_type({"proceedings": "CVPR"}) == "C"


def test_detect_type_dissertation_by_university():
    """有学位授予单位 -> D（学位论文）。"""
    assert _detect_type({"university": "清华大学"}) == "D"
    assert _detect_type({"degree_grantor": "北京大学"}) == "D"


def test_detect_type_explicit_card_type_wins():
    """card 显式指定 type 时优先使用。"""
    assert _detect_type({"type": "M", "publisher": "X"}) == "M"
    assert _detect_type({"type": "D", "university": "X"}) == "D"


def test_detect_type_publisher_with_journal_stays_journal():
    """有出版社但有期刊名 -> 仍为 J（期刊文章也有出版社）。"""
    assert _detect_type({"publisher": "Nature", "journal": "Nature"}) == "J"


def test_entry_from_card_uses_detection():
    """entry_from_card 用 _detect_type 而非恒等 J。"""
    entry = entry_from_card({"title": "专著", "publisher": "出版社"})
    assert entry["type"] == "M"
    entry2 = entry_from_card({"title": "学位论文", "university": "清华"})
    assert entry2["type"] == "D"


def test_entry_from_card_llm_type_passthrough():
    """LLM 返回的 type 经 card 传入 entry_from_card。"""
    entry = entry_from_card({"title": "x", "type": "C"})
    assert entry["type"] == "C"


# ---------------------------------------------------------------------------
#  T1-2: 多样式格式化
# ---------------------------------------------------------------------------
_ENTRY_J = {
    "title": "深度学习综述",
    "authors": ["张三", "李四"],
    "year": "2023",
    "type": "J",
    "journal": "计算机学报",
    "volume": "10",
    "issue": "2",
    "pages": "1-20",
    "doi": "10.1234/abc",
}

_ENTRY_M = {
    "title": "机器学习导论",
    "authors": ["王五"],
    "year": "2022",
    "type": "M",
    "publisher": "清华大学出版社",
}


def test_format_gbt_journal():
    out = format_gbt(_ENTRY_J)
    assert "[J]" in out
    assert "张三" in out
    assert "计算机学报" in out
    assert "2023" in out
    assert "10(2)" in out
    assert "1-20" in out
    assert "DOI:10.1234/abc" in out


def test_format_apa_journal():
    out = format_apa(_ENTRY_J)
    assert "(2023)" in out
    assert "深度学习综述" in out
    assert "计算机学报" in out
    assert "10(2)" in out
    assert "1-20" in out
    assert "doi.org/10.1234/abc" in out


def test_format_mla_journal():
    out = format_mla(_ENTRY_J)
    assert '"深度学习综述."' in out
    assert "计算机学报" in out
    assert "vol. 10" in out
    assert "no. 2" in out
    assert "2023" in out
    assert "pp. 1-20" in out


def test_format_chicago_journal():
    out = format_chicago(_ENTRY_J)
    assert "2023." in out
    assert '"深度学习综述."' in out
    assert "计算机学报" in out
    assert "10 (2)" in out
    assert "1-20" in out


def test_format_gbt_monograph():
    out = format_gbt(_ENTRY_M)
    assert "[M]" in out
    assert "清华大学出版社" in out


def test_format_apa_monograph():
    out = format_apa(_ENTRY_M)
    assert "(2022)" in out
    assert "机器学习导论" in out
    assert "清华大学出版社" in out


def test_format_reference_dispatcher_default():
    """默认使用 GB/T 7714。"""
    out = format_reference(_ENTRY_J)
    assert "[J]" in out


def test_format_reference_dispatcher_apa():
    out = format_reference(_ENTRY_J, "APA")
    assert "[J]" not in out
    assert "(2023)" in out


def test_format_reference_dispatcher_unknown_falls_back():
    """未知样式回退 GB/T 7714。"""
    out = format_reference(_ENTRY_J, "Unknown")
    assert "[J]" in out


def test_format_handles_missing_fields():
    """缺失字段用占位符。"""
    entry = {"title": "仅标题", "type": "J"}
    out = format_gbt(entry)
    assert "«请补全作者»" in out
    assert "«请补全期刊»" in out

    out_apa = format_apa(entry)
    assert "«请补全年份»" in out_apa

    out_mla = format_mla(entry)
    assert "«请补全作者»" in out_mla


def test_format_apa_multiple_authors():
    """多作者用 & 连接末位。"""
    entry = {"title": "T", "authors": ["Smith John", "Brown Alice", "Lee Bob"],
             "year": "2023", "type": "J", "journal": "J", "volume": "1"}
    out = format_apa(entry)
    assert "Smith, J." in out
    assert "Brown, A." in out
    assert "Lee, B." in out
    assert "&" in out
