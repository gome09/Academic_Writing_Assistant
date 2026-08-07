# -*- coding: utf-8 -*-
"""T0-6: lookup_crossref 单测——缓存命中/miss/DOI 过滤/超时，全部 monkeypatch。"""
from __future__ import annotations

import io
import json
from urllib.error import URLError

import pytest

from src import references


# ---------------------------------------------------------------------------
#  辅助：模拟 urlopen 上下文
# ---------------------------------------------------------------------------
class _FakeResponse:
    """模拟 urllib.request.urlopen 返回的上下文管理器。"""

    def __init__(self, payload: dict | None):
        self._data = json.dumps(payload or {}).encode("utf-8")

    def __enter__(self):
        return io.BytesIO(self._data)

    def __exit__(self, *args):
        return False


def _make_crossref_item(title="Deep Learning", doi="10.1234/abc",
                        authors=None, year=2023, journal="Nature"):
    return {
        "DOI": doi,
        "title": [title],
        "author": authors or [{"family": "Zhang", "given": "San"}],
        "container-title": [journal],
        "volume": "10",
        "issue": "2",
        "page": "1-20",
        "published": {"date-parts": [[year]]},
    }


# ---------------------------------------------------------------------------
#  测试
# ---------------------------------------------------------------------------
def test_lookup_returns_empty_when_no_key():
    """无 title 且无 doi -> 空字典。"""
    assert references.lookup_crossref("") == {}
    assert references.lookup_crossref("", "") == {}


def test_cache_hit_skips_network(tmp_path, monkeypatch):
    """缓存命中时不发起网络请求。"""
    cache = tmp_path / "cache.json"
    item = _make_crossref_item(title="Cached Paper")
    cache.write_text(json.dumps({"Cached Paper": item}), encoding="utf-8")

    def fail(*a, **kw):
        pytest.fail("缓存命中不应发起网络请求")

    monkeypatch.setattr(references.urllib.request, "urlopen", fail)
    result = references.lookup_crossref("Cached Paper", cache_path=str(cache))
    assert result == item


def test_cache_miss_fetches_and_caches(tmp_path, monkeypatch):
    """缓存未命中 -> 发起网络请求 -> 写缓存。"""
    cache = tmp_path / "cache.json"
    item = _make_crossref_item(title="New Paper")

    calls = []

    def fake_urlopen(req, timeout=10):
        calls.append(req.full_url)
        return _FakeResponse({"message": {"items": [item]}})

    monkeypatch.setattr(references.urllib.request, "urlopen", fake_urlopen)
    result = references.lookup_crossref("New Paper", cache_path=str(cache))
    assert result == item
    assert len(calls) == 1
    # 二次调用命中缓存，不再请求
    references.lookup_crossref("New Paper", cache_path=str(cache))
    assert len(calls) == 1
    # 缓存文件已写入
    assert cache.exists()
    cached = json.loads(cache.read_text(encoding="utf-8"))
    assert "New Paper" in cached


def test_doi_query_uses_filter(tmp_path, monkeypatch):
    """传入 DOI 时使用 filter 参数而非 query.bibliographic。"""
    cache = tmp_path / "cache.json"
    item = _make_crossref_item(doi="10.1234/xyz")
    seen_urls = []

    def fake_urlopen(req, timeout=10):
        seen_urls.append(req.full_url)
        return _FakeResponse({"message": {"items": [item]}})

    monkeypatch.setattr(references.urllib.request, "urlopen", fake_urlopen)
    references.lookup_crossref("Some Title", doi="10.1234/xyz",
                               cache_path=str(cache))
    assert seen_urls
    assert "filter=doi" in seen_urls[0]
    assert "query.bibliographic" not in seen_urls[0]


def test_title_query_uses_bibliographic(tmp_path, monkeypatch):
    """仅传 title 时使用 query.bibliographic。"""
    cache = tmp_path / "cache.json"
    item = _make_crossref_item()
    seen_urls = []

    def fake_urlopen(req, timeout=10):
        seen_urls.append(req.full_url)
        return _FakeResponse({"message": {"items": [item]}})

    monkeypatch.setattr(references.urllib.request, "urlopen", fake_urlopen)
    references.lookup_crossref("My Title", cache_path=str(cache))
    assert seen_urls
    assert "query.bibliographic" in seen_urls[0]


def test_network_error_returns_empty(tmp_path, monkeypatch):
    """网络超时/URLError -> 不中断，返回空字典（当前实现未包 try/except，
    此测试验证期望行为——T1-3 将加 try/except 使其优雅降级）。"""
    cache = tmp_path / "cache.json"

    def fake_urlopen(req, timeout=10):
        raise URLError("timeout")

    monkeypatch.setattr(references.urllib.request, "urlopen", fake_urlopen)
    # 当前 lookup_crossref 不包 try/except，URLError 会抛出
    # T1-3 修复后此断言改为 == {}
    with pytest.raises(URLError):
        references.lookup_crossref("Fail", cache_path=str(cache))


def test_empty_items_returns_empty(tmp_path, monkeypatch):
    """Crossref 返回空 items -> 空字典。"""
    cache = tmp_path / "cache.json"

    def fake_urlopen(req, timeout=10):
        return _FakeResponse({"message": {"items": []}})

    monkeypatch.setattr(references.urllib.request, "urlopen", fake_urlopen)
    result = references.lookup_crossref("Nothing", cache_path=str(cache))
    assert result == {}


def test_corrupt_cache_falls_back_to_network(tmp_path, monkeypatch):
    """缓存文件损坏 -> 回退到网络请求。"""
    cache = tmp_path / "cache.json"
    cache.write_text("NOT JSON {{{", encoding="utf-8")
    item = _make_crossref_item(title="Corrupt Cache")

    def fake_urlopen(req, timeout=10):
        return _FakeResponse({"message": {"items": [item]}})

    monkeypatch.setattr(references.urllib.request, "urlopen", fake_urlopen)
    result = references.lookup_crossref("Corrupt Cache", cache_path=str(cache))
    assert result == item


def test_no_cache_path_still_works(monkeypatch):
    """不传 cache_path -> 不读/写缓存文件，仍可查询。"""
    item = _make_crossref_item()

    def fake_urlopen(req, timeout=10):
        return _FakeResponse({"message": {"items": [item]}})

    monkeypatch.setattr(references.urllib.request, "urlopen", fake_urlopen)
    result = references.lookup_crossref("No Cache")
    assert result == item
