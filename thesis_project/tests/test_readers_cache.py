# -*- coding: utf-8 -*-
"""T5-1：read_file 与缓存的集成测试。

验证：
  - cache=None 时行为完全不变（向后兼容）。
  - 传入 cache 时未变文件命中缓存、跳过重读。
  - 文件内容变化后缓存失效、重新读取。
  - 含图片的文档不进缓存。
  - dry-run 语义由调用方保证（此处只验 cache=None 不落盘）。
"""
from src.readers import read_file
from src.cache import ReadCache


def _txt(tmp_path, name="a.txt", content="正文段落。"):
    f = tmp_path / name
    f.write_text(content, encoding="utf-8")
    return str(f)


def test_read_file_without_cache_unchanged(tmp_path):
    """cache=None 时 read_file 行为与历史一致。"""
    p = _txt(tmp_path, content="无缓存读取")
    doc = read_file(p)
    assert doc["type"] == "txt"
    assert doc["blocks"][0]["text"] == "无缓存读取"


def test_read_file_cache_miss_then_hit(tmp_path, monkeypatch):
    """首次读 → 写入缓存；二次读 → 命中缓存（跳过重读，不调用底层读取器）。"""
    p = _txt(tmp_path, content="首次内容")
    cache = ReadCache(str(tmp_path / "c.pkl"))

    doc1 = read_file(p, cache=cache)
    assert doc1["blocks"][0]["text"] == "首次内容"
    assert cache.hits == 0  # 首次未命中

    # 二次读：monkeypatch 底层读取器抛错，证明走缓存而未真正读盘
    from src import readers
    def _boom(path):
        raise AssertionError("不应调用底层读取器：应命中缓存")
    monkeypatch.setitem(readers._READERS, ".txt", _boom)

    doc2 = read_file(p, cache=cache)
    assert doc2["blocks"][0]["text"] == "首次内容"
    assert cache.hits == 1


def test_read_file_cache_invalidates_on_change(tmp_path):
    """文件内容变化后哈希变化，缓存未命中，重新读取。"""
    p = _txt(tmp_path, content="旧内容")
    cache = ReadCache(str(tmp_path / "c.pkl"))
    read_file(p, cache=cache)

    # 修改文件内容
    _txt(tmp_path, content="新内容已更新")
    doc = read_file(p, cache=cache)
    assert doc["blocks"][0]["text"] == "新内容已更新"
    assert cache.hits == 0  # 仍未命中（哈希变了）


def test_read_file_image_doc_not_cached(tmp_path):
    """含 image 块的文档不进缓存（避免大字节存储）。"""
    import struct
    import zlib
    from src.cache import file_hash

    def _chunk(typ, data):
        return (struct.pack(">I", len(data)) + typ + data
                + struct.pack(">I", zlib.crc32(typ + data) & 0xffffffff))

    # 构造一个完整的最小 1x1 PNG（签名 + IHDR + IDAT + IEND）
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = _chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 0, 0, 0, 0))
    idat = _chunk(b"IDAT", zlib.compress(b"\x00\x00"))
    iend = _chunk(b"IEND", b"")
    p = tmp_path / "img.png"
    p.write_bytes(sig + ihdr + idat + iend)

    cache = ReadCache(str(tmp_path / "c.pkl"))
    doc = read_file(str(p), cache=cache)
    assert doc["type"] == "image"
    assert any(b.get("kind") == "image" for b in doc["blocks"])
    # 含 image 块 → 未写入缓存
    assert len(cache) == 0
    assert cache.get(str(p), file_hash(str(p)), False, False) is None


def test_read_dir_detailed_with_cache(tmp_path):
    """read_dir_detailed 透传 cache，未变文件二次命中。"""
    from src.readers import read_dir_detailed
    _txt(tmp_path, "a.txt", "目录内文件")
    cache = ReadCache(str(tmp_path / "c.pkl"))

    docs1, errs1 = read_dir_detailed(str(tmp_path), cache=cache)
    assert len(docs1) == 1 and not errs1

    docs2, errs2 = read_dir_detailed(str(tmp_path), cache=cache)
    assert len(docs2) == 1 and not errs2
    assert cache.hits == 1  # 二次全部命中
