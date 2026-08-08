# -*- coding: utf-8 -*-
"""T1-4: 可选语义文献检索测试——全 monkeypatch，不触真实网络。"""
from __future__ import annotations

import io
import json

from src import literature_search


class _FakeResponse:
    def __init__(self, payload):
        self._data = json.dumps(payload).encode("utf-8")

    def __enter__(self):
        return io.BytesIO(self._data)

    def __exit__(self, *a):
        return False


def _monkeypatch_urlopen(monkeypatch, payload):
    """Mock urlopen to return the given payload dict."""
    def fake_urlopen(req, timeout=15):
        return _FakeResponse(payload)
    monkeypatch.setattr(literature_search.urllib.request, "urlopen", fake_urlopen)


def _monkeypatch_urlopen_error(monkeypatch, exc):
    """Mock urlopen to raise."""
    def fake_urlopen(req, timeout=15):
        raise exc
    monkeypatch.setattr(literature_search.urllib.request, "urlopen", fake_urlopen)


# ---------------------------------------------------------------------------
#  OpenAlex
# ---------------------------------------------------------------------------
def test_search_openalex_parses_results(monkeypatch):
    payload = {
        "results": [
            {
                "title": "Deep Learning for Image Classification",
                "authorships": [
                    {"author": {"display_name": "Zhang San"}},
                    {"author": {"display_name": "Li Si"}},
                ],
                "publication_year": 2023,
                "doi": "https://doi.org/10.1234/abc",
                "abstract_inverted_index": {"deep": [0], "learning": [1],
                                            "is": [2], "great": [3]},
            },
        ]
    }
    _monkeypatch_urlopen(monkeypatch, payload)
    results = literature_search.search_openalex("deep learning")
    assert len(results) == 1
    r = results[0]
    assert r["title"] == "Deep Learning for Image Classification"
    assert "Zhang San" in r["authors"]
    assert r["year"] == "2023"
    assert r["doi"] == "10.1234/abc"
    assert "deep learning is great" in r["abstract"]


def test_search_openalex_empty_results(monkeypatch):
    _monkeypatch_urlopen(monkeypatch, {"results": []})
    results = literature_search.search_openalex("nothing")
    assert results == []


def test_search_openalex_network_error(monkeypatch):
    _monkeypatch_urlopen_error(monkeypatch, OSError("timeout"))
    results = literature_search.search_openalex("fail")
    assert results == []


def test_search_openalex_no_doi(monkeypatch):
    payload = {
        "results": [
            {"title": "No DOI Paper", "authorships": [],
             "publication_year": 2020, "doi": None,
             "abstract_inverted_index": {}},
        ]
    }
    _monkeypatch_urlopen(monkeypatch, payload)
    results = literature_search.search_openalex("test")
    assert len(results) == 1
    assert results[0]["doi"] == ""


# ---------------------------------------------------------------------------
#  Semantic Scholar
# ---------------------------------------------------------------------------
def test_search_s2_parses_results(monkeypatch):
    payload = {
        "data": [
            {
                "title": "Transformer Networks",
                "authors": [{"name": "Vaswani"}, {"name": "Shazeer"}],
                "year": 2017,
                "externalIds": {"DOI": "10.5555/3295222.3295349"},
                "abstract": "Attention is all you need.",
                "url": "https://semanticscholar.org/paper/123",
            },
        ]
    }
    _monkeypatch_urlopen(monkeypatch, payload)
    results = literature_search.search_semantic_scholar("transformer")
    assert len(results) == 1
    r = results[0]
    assert r["title"] == "Transformer Networks"
    assert "Vaswani" in r["authors"]
    assert r["year"] == "2017"
    assert r["doi"] == "10.5555/3295222.3295349"
    assert "Attention" in r["abstract"]


def test_search_s2_network_error(monkeypatch):
    _monkeypatch_urlopen_error(monkeypatch, OSError("timeout"))
    results = literature_search.search_semantic_scholar("fail")
    assert results == []


# ---------------------------------------------------------------------------
#  统一分发器
# ---------------------------------------------------------------------------
def test_search_literature_dispatches_openalex(monkeypatch):
    payload = {"results": [{"title": "A", "authorships": [],
                            "publication_year": 2020, "doi": None,
                            "abstract_inverted_index": {}}]}
    calls = []

    def fake_urlopen(req, timeout=15):
        calls.append(req.full_url)
        return _FakeResponse(payload)

    monkeypatch.setattr(literature_search.urllib.request, "urlopen", fake_urlopen)
    results = literature_search.search_literature("test", source="openalex")
    assert len(results) == 1
    assert "openalex.org" in calls[0]


def test_search_literature_dispatches_s2(monkeypatch):
    payload = {"data": [{"title": "B", "authors": [], "year": 2020,
                         "externalIds": {}, "abstract": "", "url": ""}]}
    calls = []

    def fake_urlopen(req, timeout=15):
        calls.append(req.full_url)
        return _FakeResponse(payload)

    monkeypatch.setattr(literature_search.urllib.request, "urlopen", fake_urlopen)
    results = literature_search.search_literature("test", source="s2")
    assert len(results) == 1
    assert "semanticscholar.org" in calls[0]


def test_search_literature_both_calls_both(monkeypatch):
    """source='both' -> 调用两个 API 并合并。"""
    payload1 = {"results": [{"title": "OA", "authorships": [],
                              "publication_year": 2020, "doi": None,
                              "abstract_inverted_index": {}}]}
    payload2 = {"data": [{"title": "S2", "authors": [], "year": 2020,
                           "externalIds": {}, "abstract": "", "url": ""}]}
    call_count = [0]

    def fake_urlopen(req, timeout=15):
        call_count[0] += 1
        return _FakeResponse(payload1 if call_count[0] == 1 else payload2)

    monkeypatch.setattr(literature_search.urllib.request, "urlopen", fake_urlopen)
    results = literature_search.search_literature("test", source="both")
    assert len(results) == 2
    assert call_count[0] == 2


# ---------------------------------------------------------------------------
#  results_to_cards
# ---------------------------------------------------------------------------
def test_results_to_cards():
    results = [
        {"title": "Paper A", "authors": ["X", "Y"], "year": "2023",
         "doi": "10.1/abc", "url": "https://doi.org/10.1/abc",
         "abstract": "Abstract text"},
    ]
    cards = literature_search.results_to_cards(results)
    assert len(cards) == 1
    c = cards[0]
    assert c["title"] == "Paper A"
    assert c["authors"] == ["X", "Y"]
    assert c["year"] == "2023"
    assert c["doi"] == "10.1/abc"
    assert c["_from_search"] is True
    assert "文献检索" in c["source"]


def test_results_to_cards_empty():
    cards = literature_search.results_to_cards([])
    assert cards == []
