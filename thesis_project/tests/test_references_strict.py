# -*- coding: utf-8 -*-
from src.references import format_gbt, validate_citations


def test_format_gbt_never_invents_missing_fields():
    text = format_gbt({"title": "文A", "authors": [], "year": "", "type": "J"})
    assert "«请补全作者»" in text
    assert "«请补全期刊»" in text


def test_validate_citations_rejects_out_of_range_and_unrelated():
    issues = validate_citations(["相关结论[2]，错误结论[4]。"], {1}, 3)
    assert any("不属于" in item for item in issues)
    assert any("超出" in item for item in issues)
