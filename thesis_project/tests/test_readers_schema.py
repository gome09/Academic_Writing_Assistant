# -*- coding: utf-8 -*-
"""T4-5：Block / Document 契约测试。

验证各读取器输出符合 TypedDict 契约：Document 含 source/type/blocks/meta，
Block 含 kind/level/text。运行时仍为 dict，TypedDict 仅类型层契约。
"""
from __future__ import annotations

from src import readers
from src.readers import Block, Document


def _assert_document_shape(doc, path, type_):
    """断言 Document 四字段齐全且类型正确。"""
    assert isinstance(doc, dict)
    assert doc["source"] == path
    assert doc["type"] == type_
    assert isinstance(doc["blocks"], list)
    assert isinstance(doc["meta"], dict)


def _assert_block_shape(block):
    """断言 Block 必填字段齐全。"""
    assert "kind" in block and isinstance(block["kind"], str)
    assert "level" in block and isinstance(block["level"], int)
    assert "text" in block and isinstance(block["text"], str)


def test_block_and_document_are_typeddicts():
    """Block / Document 是 TypedDict（类型层契约存在，字段集完整）。

    注：模块启用 ``from __future__ import annotations`` 后，``NotRequired``
    在运行时不可内省（CPython 已知限制），但静态类型检查器（mypy/pyright）
    仍能正确识别可选字段。此处仅断言字段集存在，可选性由静态检查保障。
    """
    assert hasattr(Block, "__required_keys__")
    assert hasattr(Document, "__required_keys__")
    # 字段集完整（必填 + 可选均在 __annotations__ 中）
    assert set(Block.__annotations__) == {"kind", "level", "text",
                                          "rows", "data", "ext"}
    assert set(Document.__annotations__) == {"source", "type", "blocks", "meta"}
    # Document 必填：source/type/blocks/meta
    assert {"source", "type", "blocks", "meta"} <= set(Document.__required_keys__)
    # Block 必填核心：kind/level/text
    assert {"kind", "level", "text"} <= set(Block.__required_keys__)


def test_block_helper_returns_contract_shape():
    """_block 返回值满足 Block 契约。"""
    b = readers._block("heading", "标题", level=2)
    _assert_block_shape(b)
    assert b["kind"] == "heading"
    assert b["level"] == 2
    assert b["text"] == "标题"


def test_read_txt_contract(tmp_path):
    p = tmp_path / "a.txt"
    p.write_text("第一段。\n\n第二段。", encoding="utf-8")
    doc = readers.read_txt(str(p))
    _assert_document_shape(doc, str(p), "txt")
    for b in doc["blocks"]:
        _assert_block_shape(b)
    assert all(b["kind"] == "paragraph" for b in doc["blocks"])


def test_read_md_contract(tmp_path):
    p = tmp_path / "a.md"
    p.write_text("# 标题\n\n正文内容。", encoding="utf-8")
    doc = readers.read_md(str(p))
    _assert_document_shape(doc, str(p), "md")
    assert doc["blocks"][0]["kind"] == "heading"
    assert doc["blocks"][0]["level"] == 1


def test_read_json_contract(tmp_path):
    p = tmp_path / "a.json"
    p.write_text('{"title": "T", "content": "正文"}', encoding="utf-8")
    doc = readers.read_json(str(p))
    _assert_document_shape(doc, str(p), "json")
    assert doc["blocks"][0]["kind"] == "heading"


def test_read_csv_contract(tmp_path):
    p = tmp_path / "a.csv"
    p.write_text("a,b\n1,2\n", encoding="utf-8")
    doc = readers.read_csv(str(p))
    _assert_document_shape(doc, str(p), "csv")
    tbl = doc["blocks"][0]
    _assert_block_shape(tbl)
    assert tbl["kind"] == "table"
    assert tbl["rows"] == [["a", "b"], ["1", "2"]]


def test_read_image_contract(tmp_path):
    p = tmp_path / "a.png"
    p.write_bytes(b"\x89PNG\r\n\x1a\nfake")
    doc = readers.read_image(str(p))
    _assert_document_shape(doc, str(p), "image")
    b = doc["blocks"][0]
    _assert_block_shape(b)
    assert b["kind"] == "image"
    assert b["data"] == b"\x89PNG\r\n\x1a\nfake"
    assert b["ext"] == ".png"


def test_ensure_has_text_accepts_image_blocks():
    """T4-2：仅有 image 块时视为有内容，不触发扫描件报错。"""
    blocks = [readers._block("image")]
    blocks[0]["data"] = b"fake"
    blocks[0]["ext"] = ".png"
    # 不应抛异常
    readers._ensure_has_text(blocks, "img-only.pdf", ocr=False)
