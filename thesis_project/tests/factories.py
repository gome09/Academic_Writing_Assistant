# -*- coding: utf-8 -*-
"""构造 readers 中间结构（Block / Document）的测试工厂。"""


def h(level, text):
    """heading 块"""
    return {"kind": "heading", "level": level, "text": text, "rows": None}


def p(text):
    """paragraph 块"""
    return {"kind": "paragraph", "level": 0, "text": text, "rows": None}


def table(rows):
    """table 块"""
    return {"kind": "table", "level": 0, "text": "", "rows": rows}


def doc(blocks, meta=None, type_="md"):
    """Document"""
    return {"source": "test", "type": type_, "blocks": blocks, "meta": meta or {}}
