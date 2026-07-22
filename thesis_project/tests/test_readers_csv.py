# -*- coding: utf-8 -*-
import pytest
from src import readers


def test_read_csv_basic(tmp_path):
    p = tmp_path / "数据.csv"
    p.write_text("组别,精度\nA,0.91\n,\n", encoding="utf-8")
    doc = readers.read_csv(str(p))
    assert doc["type"] == "csv"
    assert doc["blocks"][0]["kind"] == "table"
    assert doc["blocks"][0]["rows"] == [["组别", "精度"], ["A", "0.91"]]  # 全空行跳过


def test_read_csv_gbk_fallback(tmp_path):
    p = tmp_path / "g.csv"
    p.write_bytes("名称,数值\n测试,1\n".encode("gb18030"))
    doc = readers.read_csv(str(p))
    assert doc["blocks"][0]["rows"][0] == ["名称", "数值"]


def test_read_csv_empty_raises(tmp_path):
    p = tmp_path / "e.csv"
    p.write_text("\n,\n", encoding="utf-8")
    with pytest.raises(RuntimeError):
        readers.read_csv(str(p))


def test_read_csv_registered(tmp_path):
    p = tmp_path / "d.csv"
    p.write_text("x\n", encoding="utf-8")
    assert readers.read_file(str(p))["type"] == "csv"
