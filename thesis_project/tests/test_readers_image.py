# -*- coding: utf-8 -*-
import pytest
from src import readers

# 1x1 红点 PNG（最小合法 PNG 字节）
PNG_BYTES = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108020000009077"
    "3df80000000c4944415408d763f8cfc00000030101cf9e46a80000000049454e44ae426082")


def test_read_image_returns_single_image_block(tmp_path):
    p = tmp_path / "架构图.png"
    p.write_bytes(PNG_BYTES)
    doc = readers.read_image(str(p))
    assert doc["type"] == "image"
    assert len(doc["blocks"]) == 1
    b = doc["blocks"][0]
    assert b["kind"] == "image"
    assert b["data"] == PNG_BYTES
    assert b["ext"] == ".png"


def test_read_image_empty_file_raises(tmp_path):
    p = tmp_path / "空.jpg"
    p.write_bytes(b"")
    with pytest.raises(RuntimeError):
        readers.read_image(str(p))


@pytest.mark.parametrize("ext", [".png", ".jpg", ".jpeg", ".bmp", ".webp"])
def test_image_extensions_registered(tmp_path, ext):
    p = tmp_path / ("t" + ext)
    p.write_bytes(PNG_BYTES)
    assert readers.read_file(str(p))["type"] == "image"
