# -*- coding: utf-8 -*-
"""T4-4：构建器插件化测试——Builder 基类、注册表、--only 动态 choices。"""
from __future__ import annotations

import pytest

from src.builders import BUILDERS, Builder, PptBuilder, WordBuilder, builder_names
from src.builders.base import register_builder


def test_default_builders_registered():
    """内置 word/ppt 构建器已注册。"""
    assert "word" in BUILDERS and "ppt" in BUILDERS
    assert isinstance(BUILDERS["word"], WordBuilder)
    assert isinstance(BUILDERS["ppt"], PptBuilder)
    assert BUILDERS["word"].ext == ".docx"
    assert BUILDERS["ppt"].ext == ".pptx"
    assert BUILDERS["word"].label == "Word"
    assert BUILDERS["ppt"].label == "PPT"


def test_builder_names_order():
    """word 先于 ppt（注册顺序保证 draft 日志可读）。"""
    names = builder_names()
    assert names.index("word") < names.index("ppt")
    assert set(names) >= {"word", "ppt"}


def test_builder_is_abstract():
    """Builder 是 ABC，不能直接实例化。"""
    with pytest.raises(TypeError):
        Builder()  # type: ignore[abstract]


def test_register_builder_custom():
    """第三方构建器可通过装饰器注册。"""

    @register_builder("demo", ".demo", label="Demo")
    class DemoBuilder(Builder):
        def build(self, data, out_path):
            return out_path

    try:
        assert "demo" in BUILDERS
        assert BUILDERS["demo"].ext == ".demo"
        assert BUILDERS["demo"].label == "Demo"
        assert "demo" in builder_names()
        # 注册实例可正常 build
        assert BUILDERS["demo"].build({}, "/tmp/out.demo") == "/tmp/out.demo"
    finally:
        # 清理：移除注册实例；builder_names() 自动跳过 _ORDER 中的残留项
        BUILDERS.pop("demo", None)
        assert "demo" not in builder_names()


def test_word_builder_delegates_to_docx_builder(monkeypatch, tmp_path):
    """WordBuilder.build 委托 docx_builder.build。"""
    from src import docx_builder

    calls = []

    def fake_build(data, out_path):
        calls.append((data, out_path))
        return out_path

    monkeypatch.setattr(docx_builder, "build", fake_build)
    out = str(tmp_path / "out.docx")
    result = BUILDERS["word"].build({"thesis": True}, out)
    assert result == out
    assert calls == [({"thesis": True}, out)]


def test_ppt_builder_delegates_to_pptx_builder(monkeypatch, tmp_path):
    """PptBuilder.build 委托 pptx_builder.build（返回 BuildResult）。"""
    from src import pptx_builder

    def fake_build(data, out_path):
        # 模拟 BuildResult：str 子类带 warnings
        class _R(str):
            warnings = ["w1"]
        return _R(out_path)

    monkeypatch.setattr(pptx_builder, "build", fake_build)
    out = str(tmp_path / "out.pptx")
    result = BUILDERS["ppt"].build({"deck": True}, out)
    assert str(result) == out
    assert getattr(result, "warnings", None) == ["w1"]


def test_post_build_default_returns_empty():
    """Builder.post_build 默认返回空告警列表。"""
    assert BUILDERS["word"].post_build("/tmp/x.docx", args=None) == []


def test_only_choices_derived_from_registry(monkeypatch):
    """--only choices 取自注册表（动态）。"""
    # 确保注册表中有 word/ppt
    names = set(builder_names())
    assert {"word", "ppt"} <= names
    # main 模块的 argparse choices 即 builder_names()
    # （直接验证 builder_names 与 --only 一致性，不解析 argv）
    import src.main as main_mod
    # builder_names 是 main 用来构造 choices 的来源
    assert set(main_mod.builder_names()) == set(builder_names())
