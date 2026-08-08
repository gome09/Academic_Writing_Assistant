# -*- coding: utf-8 -*-
"""构建器包（T4-4）：注册内置 word/ppt 构建器。

第三方可在导入本包后用 register_builder 注册新构建器（如 HTML/Markdown），
main.py 的 --only choices 与 draft 生成循环自动纳入。
"""
from __future__ import annotations

from src.builders.base import (
    BUILDERS, Builder, builder_names, register_builder,
)


@register_builder("word", ".docx", label="Word")
class WordBuilder(Builder):
    """Word 草案构建器适配：委托 docx_builder.build。"""

    def build(self, data, out_path):
        from src import docx_builder
        return docx_builder.build(data, out_path)


@register_builder("ppt", ".pptx", label="PPT")
class PptBuilder(Builder):
    """PPT 草案构建器适配：委托 pptx_builder.build（返回 BuildResult）。"""

    def build(self, data, out_path):
        from src import pptx_builder
        return pptx_builder.build(data, out_path)


__all__ = ["BUILDERS", "Builder", "builder_names", "register_builder",
           "WordBuilder", "PptBuilder"]
