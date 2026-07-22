# -*- coding: utf-8 -*-
import openpyxl
import pytest
from src import readers


def _make_wb(path, sheets):
    """sheets: {表名: [[行], ...]}；写一个真实 xlsx 供读取。"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


def test_read_xlsx_each_nonempty_sheet_becomes_table(tmp_path):
    p = str(tmp_path / "数据.xlsx")
    _make_wb(p, {"实验": [["组别", "精度"], ["A", 0.91]], "空表": [], "问卷": [["Q1", 5]]})
    doc = readers.read_xlsx(p)
    assert doc["type"] == "xlsx"
    tables = [b for b in doc["blocks"] if b["kind"] == "table"]
    assert len(tables) == 2                      # 空表被跳过
    assert tables[0]["rows"][0] == ["组别", "精度"]
    assert tables[0]["rows"][1] == ["A", "0.91"]  # 数值转字符串


def test_read_xlsx_none_cells_become_empty_string(tmp_path):
    p = str(tmp_path / "n.xlsx")
    _make_wb(p, {"s": [["a", None, "c"]]})
    doc = readers.read_xlsx(p)
    assert doc["blocks"][0]["rows"][0] == ["a", "", "c"]


def test_read_xlsx_all_empty_raises(tmp_path):
    p = str(tmp_path / "e.xlsx")
    _make_wb(p, {"s1": [], "s2": [[None, None]]})
    with pytest.raises(RuntimeError):
        readers.read_xlsx(p)


def test_read_xlsx_registered_in_dispatch(tmp_path):
    p = str(tmp_path / "d.xlsx")
    _make_wb(p, {"s": [["x"]]})
    assert readers.read_file(p)["type"] == "xlsx"
