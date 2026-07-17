# -*- coding: utf-8 -*-
from src.organizer import _table_to_text
from tests.factories import table


def test_all_rows_kept():
    blk = table([["指标", "数值"],
                 ["准确率", "94.2%"],
                 ["延迟", "85ms"],
                 ["体积", "4MB"]])
    text = _table_to_text(blk)
    assert "准确率 | 94.2%" in text
    assert "延迟 | 85ms" in text
    assert "体积 | 4MB" in text


def test_empty_cells_skipped():
    blk = table([["a", "", "b"], ["", "", ""]])
    assert _table_to_text(blk) == "a | b"
