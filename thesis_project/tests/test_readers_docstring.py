# -*- coding: utf-8 -*-
"""固化 _read_text 三步回退顺序的语义，防止后续误改。"""
import inspect

from src.readers import _read_text


def test_read_text_docstring_lists_three_steps():
    src = inspect.getsource(_read_text)
    # 关键串必须直接出现在 docstring 中
    assert "utf-8-sig" in src
    assert "gb18030" in src
    # 兜底 errors=replace 也必须出现
    assert "errors=\"replace\"" in src


def test_read_text_explicit_docstring_in_function():
    doc = inspect.getdoc(_read_text) or ""
    # 三步编码回退顺序在 docstring 中显式列出
    assert "utf-8-sig" in doc
    assert "gb18030" in doc
    assert "utf-8" in doc and "errors" in doc
