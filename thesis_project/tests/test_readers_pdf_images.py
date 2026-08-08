# -*- coding: utf-8 -*-
"""T4-2：PDF 内嵌图片提取测试。

pypdf 为可选依赖，用 monkeypatch 注入假模块，不触真实网络/文件。
"""
from __future__ import annotations

import sys
import types

from src import readers


def _fake_pypdf(images_per_page):
    """构造假 pypdf 模块：PdfReader(path).pages[i].images -> list[SimpleNamespace]。"""
    pages = []
    for imgs in images_per_page:
        fake_images = [types.SimpleNamespace(data=d, name=n)
                       for d, n in imgs]
        pages.append(types.SimpleNamespace(images=fake_images))
    return types.SimpleNamespace(PdfReader=lambda path: types.SimpleNamespace(pages=pages))


def test_extract_pdf_images_returns_image_blocks(monkeypatch):
    """pypdf 可用时，提取内嵌图片为 image 块。"""
    monkeypatch.setitem(sys.modules, "pypdf",
                        _fake_pypdf([[(b"\x89PNGfake1", "im0.png"),
                                      (b"\x89PNGfake2", "im1.jpg")]]))
    blocks, count = readers._extract_pdf_images("fake.pdf")
    assert count == 2
    assert all(b["kind"] == "image" for b in blocks)
    assert blocks[0]["data"] == b"\x89PNGfake1"
    assert blocks[0]["ext"] == ".png"
    assert blocks[1]["data"] == b"\x89PNGfake2"
    assert blocks[1]["ext"] == ".jpg"


def test_extract_pdf_images_no_pypdf(monkeypatch):
    """pypdf 未安装时返回 ([], -1) 标识不可用，不抛异常。"""
    # sys.modules["pypdf"] = None 触发 ImportError
    monkeypatch.setitem(sys.modules, "pypdf", None)
    blocks, count = readers._extract_pdf_images("fake.pdf")
    assert blocks == []
    assert count == -1


def test_extract_pdf_images_skips_unreadable_image(monkeypatch):
    """单张图提取失败不影响其它图。"""
    good = types.SimpleNamespace(data=b"ok", name="im1.png")

    # 构造：第一页抛异常，第二页正常
    class _BadPage:
        @property
        def images(self):
            raise RuntimeError("decode error")

    fake = types.SimpleNamespace(
        PdfReader=lambda path: types.SimpleNamespace(
            pages=[_BadPage(), types.SimpleNamespace(images=[good])]))
    monkeypatch.setitem(sys.modules, "pypdf", fake)
    blocks, count = readers._extract_pdf_images("fake.pdf")
    # 第二页的 good 被提取，第一页错误被跳过
    assert count == 1
    assert blocks[0]["data"] == b"ok"


def test_extract_pdf_images_no_name_defaults_png(monkeypatch):
    """图片无 name 时扩展名默认 .png。"""
    img = types.SimpleNamespace(data=b"x", name="")
    fake = types.SimpleNamespace(
        PdfReader=lambda path: types.SimpleNamespace(
            pages=[types.SimpleNamespace(images=[img])]))
    monkeypatch.setitem(sys.modules, "pypdf", fake)
    blocks, count = readers._extract_pdf_images("fake.pdf")
    assert count == 1
    assert blocks[0]["ext"] == ".png"


def test_read_file_passes_extract_images_to_pdf(monkeypatch, tmp_path):
    """read_file(path, extract_images=True) 把 extract_images 传给 read_pdf。"""
    received = []

    def fake_read_pdf(path, ocr=False, extract_images=False):
        received.append(extract_images)
        return {"source": path, "type": "pdf",
                "blocks": [{"kind": "paragraph", "text": "x", "level": 0}],
                "meta": {}}

    monkeypatch.setattr(readers, "read_pdf", fake_read_pdf)
    fake_pdf = tmp_path / "t.pdf"
    fake_pdf.write_bytes(b"%PDF-1.4 fake")
    readers.read_file(str(fake_pdf), extract_images=True)
    assert received == [True]


def test_read_file_extract_images_default_false(monkeypatch, tmp_path):
    """默认 extract_images=False。"""
    received = []

    def fake_read_pdf(path, ocr=False, extract_images=False):
        received.append(extract_images)
        return {"source": path, "type": "pdf",
                "blocks": [{"kind": "paragraph", "text": "x", "level": 0}],
                "meta": {}}

    monkeypatch.setattr(readers, "read_pdf", fake_read_pdf)
    fake_pdf = tmp_path / "t.pdf"
    fake_pdf.write_bytes(b"%PDF-1.4 fake")
    readers.read_file(str(fake_pdf))
    assert received == [False]
