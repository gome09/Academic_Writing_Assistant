# -*- coding: utf-8 -*-
"""T4-3：读取器插件化测试——register_reader 装饰器注册新读取器。"""
from __future__ import annotations

from src import readers


def test_default_readers_registered():
    """内置扩展名均已注册到 _READERS。"""
    expected = {
        ".txt", ".text", ".md", ".markdown", ".json", ".docx", ".pdf",
        ".xlsx", ".csv", ".png", ".jpg", ".jpeg", ".bmp", ".webp",
    }
    assert expected <= set(readers._READERS)
    # .pdf 特殊：在 read_file 中显式分发，但仍在 _READERS 以便 read_dir_detailed 过滤
    assert readers._READERS[".pdf"] is readers.read_pdf
    assert readers._READERS[".md"] is readers.read_md


def test_register_reader_adds_extension(monkeypatch, tmp_path):
    """第三方读取器可通过装饰器注册并被 read_file 分发。"""
    calls = []

    @readers.register_reader(".xyz")
    def read_xyz(path):
        calls.append(path)
        return {"source": path, "type": "xyz",
                "blocks": [readers._block("paragraph", "XYZ 内容")],
                "meta": {}}

    try:
        # .xyz 现已注册
        assert readers._READERS[".xyz"] is read_xyz
        p = tmp_path / "demo.xyz"
        p.write_text("anything", encoding="utf-8")
        doc = readers.read_file(str(p))
        assert doc["type"] == "xyz"
        assert doc["blocks"][0]["text"] == "XYZ 内容"
        assert calls == [str(p)]
    finally:
        # 清理：移除测试注册的扩展名，避免污染其它测试
        readers._READERS.pop(".xyz", None)


def test_register_reader_multiple_exts(monkeypatch):
    """一个读取器可注册到多个扩展名。"""

    @readers.register_reader(".foo", ".bar")
    def read_foo(path):
        return {"source": path, "type": "foo", "blocks": [], "meta": {}}

    try:
        assert readers._READERS[".foo"] is read_foo
        assert readers._READERS[".bar"] is read_foo
    finally:
        readers._READERS.pop(".foo", None)
        readers._READERS.pop(".bar", None)


def test_read_dir_detailed_includes_registered_ext(monkeypatch, tmp_path):
    """read_dir_detailed 通过 _READERS 过滤，注册的扩展名会被读取。"""
    @readers.register_reader(".demo")
    def read_demo(path):
        return {"source": path, "type": "demo",
                "blocks": [readers._block("paragraph", "demo")], "meta": {}}

    try:
        p = tmp_path / "x.demo"
        p.write_text("x", encoding="utf-8")
        # 同时放一个不支持的扩展名，确保被跳过
        (tmp_path / "y.unsupported").write_text("x", encoding="utf-8")
        docs, errors = readers.read_dir_detailed(str(tmp_path))
        types = {d["type"] for d in docs}
        assert "demo" in types
        # .unsupported 不在 _READERS，被跳过且不出错
        assert all("y.unsupported" not in e[0] for e in errors)
    finally:
        readers._READERS.pop(".demo", None)
