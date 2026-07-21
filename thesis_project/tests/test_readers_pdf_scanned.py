# -*- coding: utf-8 -*-
import pytest
from src import readers


def test_ensure_has_text_raises_for_scanned_pdf():
    blocks = [readers._block("paragraph", "")]
    with pytest.raises(RuntimeError, match="扫描件"):
        readers._ensure_has_text(blocks, "a.pdf")


def test_ensure_has_text_raises_for_empty_blocks():
    with pytest.raises(RuntimeError):
        readers._ensure_has_text([], "a.pdf")


def test_ensure_has_text_passes_with_text():
    readers._ensure_has_text([readers._block("paragraph", "有内容")], "a.pdf")


def test_ensure_has_text_passes_with_table_only():
    readers._ensure_has_text(
        [readers._block("table", "", rows=[["a", "b"]])], "a.pdf")
