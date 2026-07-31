# -*- coding: utf-8 -*-
"""refs 参考文献表只使用本地元数据进行确定性格式化。"""
from src.references import entry_from_card, format_gbt


def test_reference_order_and_missing_metadata_are_preserved():
    cards = [
        {"title": "文A", "authors": ["李四"], "year": "2023", "source": "a.pdf"},
        {"title": "文B", "authors": [], "year": "", "source": "b.pdf"},
    ]

    refs = [format_gbt(entry_from_card(card)) for card in cards]

    assert refs[0].startswith("李四. 文A")
    assert "«请补全作者»" in refs[1]
    assert "«请补全期刊»" in refs[1]
