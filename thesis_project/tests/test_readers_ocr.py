# -*- coding: utf-8 -*-
"""T4-1: OCR 可选支持测试——无依赖时优雅报错，有依赖时识别成功。"""
from __future__ import annotations

import pytest

from src import readers


def test_ocr_available_returns_bool(monkeypatch):
    """_ocr_available 返回布尔值。"""
    # 无论结果如何，应返回 True/False 而非抛异常
    result = readers._ocr_available()
    assert isinstance(result, bool)


def test_ensure_has_text_passes_when_text_exists():
    """有文字时不触发 OCR。"""
    blocks = [{"kind": "paragraph", "text": "hello"}]
    # 不应抛异常
    readers._ensure_has_text(blocks, "fake.pdf", ocr=True)


def test_ensure_has_text_raises_without_ocr():
    """无文字、ocr=False -> 抛 RuntimeError。"""
    blocks = []
    with pytest.raises(RuntimeError, match="扫描件"):
        readers._ensure_has_text(blocks, "fake.pdf", ocr=False)


def test_ensure_has_text_raises_when_ocr_unavailable(monkeypatch):
    """无文字、ocr=True 但 OCR 依赖未安装 -> 抛 RuntimeError。"""
    monkeypatch.setattr(readers, "_ocr_available", lambda: False)
    monkeypatch.setattr(readers, "_ocr_pdf",
                        lambda p: (_ for _ in ()).throw(
                            RuntimeError("需要安装 pytesseract")))
    blocks = []
    with pytest.raises(RuntimeError, match="扫描件"):
        readers._ensure_has_text(blocks, "fake.pdf", ocr=True)


def test_ensure_has_text_ocr_success(monkeypatch, capsys):
    """无文字、ocr=True 且 OCR 成功 -> blocks 被填充。"""
    ocr_blocks = [{"kind": "paragraph", "text": "OCR 识别的文字"}]
    monkeypatch.setattr(readers, "_ocr_pdf", lambda p: ocr_blocks)
    blocks = []
    readers._ensure_has_text(blocks, "scanned.pdf", ocr=True)
    assert any(b.get("text") for b in blocks)
    assert "OCR 识别成功" in capsys.readouterr().out


def test_ensure_has_text_ocr_no_text(monkeypatch, capsys):
    """ocr=True 但 OCR 也没识别到文字 -> 抛 RuntimeError。"""
    monkeypatch.setattr(readers, "_ocr_pdf", lambda p: [])
    blocks = []
    with pytest.raises(RuntimeError, match="扫描件"):
        readers._ensure_has_text(blocks, "blank.pdf", ocr=True)
    assert "OCR 未识别到文字" in capsys.readouterr().out


def test_ensure_has_text_ocr_failure_falls_back(monkeypatch, capsys):
    """OCR 抛异常 -> 降级为原报错。"""
    def boom(path):
        raise ValueError("tesseract not found")

    monkeypatch.setattr(readers, "_ocr_pdf", boom)
    blocks = []
    with pytest.raises(RuntimeError, match="扫描件"):
        readers._ensure_has_text(blocks, "bad.pdf", ocr=True)
    assert "OCR 失败" in capsys.readouterr().out


def test_read_file_passes_ocr_to_pdf(monkeypatch, tmp_path):
    """read_file(path, ocr=True) 把 ocr 传给 read_pdf。"""
    received_ocr = []

    def fake_read_pdf(path, ocr=False, extract_images=False):
        received_ocr.append(ocr)
        return {"source": path, "type": "pdf", "blocks": [], "meta": {}}

    monkeypatch.setattr(readers, "read_pdf", fake_read_pdf)
    # 创建一个假 PDF 文件
    fake_pdf = tmp_path / "test.pdf"
    fake_pdf.write_bytes(b"%PDF-1.4 fake")
    readers.read_file(str(fake_pdf), ocr=True)
    assert received_ocr == [True]


def test_read_file_no_ocr_by_default(monkeypatch, tmp_path):
    """默认 ocr=False。"""
    received_ocr = []

    def fake_read_pdf(path, ocr=False, extract_images=False):
        received_ocr.append(ocr)
        return {"source": path, "type": "pdf", "blocks": [], "meta": {}}

    monkeypatch.setattr(readers, "read_pdf", fake_read_pdf)
    fake_pdf = tmp_path / "test.pdf"
    fake_pdf.write_bytes(b"%PDF-1.4 fake")
    readers.read_file(str(fake_pdf))
    assert received_ocr == [False]
