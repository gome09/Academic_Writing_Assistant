# -*- coding: utf-8 -*-
"""T5-1：增量读取缓存单测。

覆盖：文件哈希稳定性、缓存命中/未命中、含 image 块不缓存、版本失效、
持久化与重载、命中计数。不触真实网络，纯本地。
"""
from src.cache import ReadCache, file_hash, CACHE_VERSION


# ---------------------------------------------------------------------------
#  file_hash
# ---------------------------------------------------------------------------
def test_file_hash_stable(tmp_path):
    """同内容同哈希；内容变化则哈希变化。"""
    f = tmp_path / "a.txt"
    f.write_text("hello", encoding="utf-8")
    h1 = file_hash(str(f))
    h2 = file_hash(str(f))
    assert h1 == h2
    assert len(h1) == 64  # SHA256 hex


def test_file_hash_changes_on_content_change(tmp_path):
    f = tmp_path / "a.txt"
    f.write_text("hello", encoding="utf-8")
    h1 = file_hash(str(f))
    f.write_text("hello world", encoding="utf-8")
    h2 = file_hash(str(f))
    assert h1 != h2


# ---------------------------------------------------------------------------
#  命中 / 未命中
# ---------------------------------------------------------------------------
def _doc(text="段落", blocks=None):
    return {"source": "x", "type": "txt",
            "blocks": blocks if blocks is not None else [
                {"kind": "paragraph", "level": 0, "text": text}],
            "meta": {}}


def test_cache_miss_returns_none(tmp_path):
    cache = ReadCache(str(tmp_path / "c.pkl"))
    assert cache.get("a.txt", "h1", False, False) is None
    assert cache.hits == 0


def test_cache_put_then_get_hits(tmp_path):
    cache = ReadCache(str(tmp_path / "c.pkl"))
    doc = _doc("内容A")
    assert cache.put("a.txt", "h1", False, False, doc) is True
    got = cache.get("a.txt", "h1", False, False)
    assert got is not None
    assert got["blocks"][0]["text"] == "内容A"
    assert cache.hits == 1


def test_cache_get_returns_deep_copy(tmp_path):
    """命中返回深拷贝，下游修改不污染缓存。"""
    cache = ReadCache(str(tmp_path / "c.pkl"))
    doc = _doc("原始")
    cache.put("a.txt", "h1", False, False, doc)
    got = cache.get("a.txt", "h1", False, False)
    got["blocks"][0]["text"] = "被改了"
    got2 = cache.get("a.txt", "h1", False, False)
    assert got2["blocks"][0]["text"] == "原始"


def test_cache_different_options_separate(tmp_path):
    """同文件不同读取选项（ocr/extract_images）视为不同条目。"""
    cache = ReadCache(str(tmp_path / "c.pkl"))
    cache.put("a.pdf", "h1", False, False, _doc("普通"))
    cache.put("a.pdf", "h1", True, False, _doc("OCR"))
    assert cache.get("a.pdf", "h1", False, False)["blocks"][0]["text"] == "普通"
    assert cache.get("a.pdf", "h1", True, False)["blocks"][0]["text"] == "OCR"


def test_cache_different_hash_separate(tmp_path):
    """同路径不同哈希（文件内容变了）视为不同条目。"""
    cache = ReadCache(str(tmp_path / "c.pkl"))
    cache.put("a.txt", "h1", False, False, _doc("旧内容"))
    cache.put("a.txt", "h2", False, False, _doc("新内容"))
    assert cache.get("a.txt", "h1", False, False)["blocks"][0]["text"] == "旧内容"
    assert cache.get("a.txt", "h2", False, False)["blocks"][0]["text"] == "新内容"


# ---------------------------------------------------------------------------
#  含 image 块不缓存
# ---------------------------------------------------------------------------
def test_cache_put_image_doc_returns_false(tmp_path):
    cache = ReadCache(str(tmp_path / "c.pkl"))
    doc = _doc(blocks=[{"kind": "image", "level": 0, "text": "",
                        "data": b"\x89PNG", "ext": ".png"}])
    assert cache.put("a.png", "h1", False, False, doc) is False
    assert cache.get("a.png", "h1", False, False) is None


# ---------------------------------------------------------------------------
#  持久化与版本
# ---------------------------------------------------------------------------
def test_cache_save_and_reload(tmp_path):
    p = tmp_path / "c.pkl"
    c1 = ReadCache(str(p))
    c1.put("a.txt", "h1", False, False, _doc("持久化"))
    c1.save()
    assert p.exists()

    c2 = ReadCache(str(p))
    got = c2.get("a.txt", "h1", False, False)
    assert got is not None
    assert got["blocks"][0]["text"] == "持久化"


def test_cache_corrupt_file_ignored(tmp_path):
    """损坏的缓存文件被静默丢弃，不抛异常。"""
    p = tmp_path / "c.pkl"
    p.write_bytes(b"not a pickle !!!")
    cache = ReadCache(str(p))
    assert cache.get("a.txt", "h1", False, False) is None
    assert len(cache) == 0


def test_cache_version_mismatch_drops_old(tmp_path):
    """版本号不匹配时丢弃旧缓存数据。"""
    import pickle
    p = tmp_path / "c.pkl"
    # 写一个旧版本号的缓存
    with open(p, "wb") as f:
        pickle.dump({"version": CACHE_VERSION - 1,
                     "data": {("a.txt", "h1", False, False): _doc("旧版")}}, f)
    cache = ReadCache(str(p))
    assert cache.get("a.txt", "h1", False, False) is None


def test_cache_save_creates_parent_dir(tmp_path):
    """save 时父目录不存在则自动创建。"""
    p = tmp_path / "sub" / "deep" / "c.pkl"
    cache = ReadCache(str(p))
    cache.put("a.txt", "h1", False, False, _doc("x"))
    cache.save()
    assert p.exists()
