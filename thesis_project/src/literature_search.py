# -*- coding: utf-8 -*-
"""T1-4：可选语义文献检索——接入 OpenAlex / Semantic Scholar 免费 API。

关键约束（红线）：
  - 默认关闭，需 --search-literature 显式启用
  - 启用需外发确认（同 LLM 机制）
  - dry-run 不触网
  - 结果带来源 URL
  - 无 API 密钥要求（OpenAlex / Semantic Scholar 均为免费公开 API）
"""
from __future__ import annotations

import json
import logging
import urllib.parse
import urllib.request

_logger = logging.getLogger("thesis_project")

_OPENALEX_BASE = "https://api.openalex.org/works"
_S2_BASE = "https://api.semanticscholar.org/graph/v1/paper/search"


def _safe_get_json(url: str, timeout: int = 15) -> dict | None:
    """发起 GET 请求并解析 JSON；网络/解析失败返回 None。"""
    req = urllib.request.Request(url, headers={"User-Agent": "thesis-project/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:  # noqa: S310
            return json.load(resp)
    except Exception as exc:  # noqa: BLE001
        _logger.warning("文献检索请求失败 %s: %s", url, exc)
        return None


def search_openalex(query: str, limit: int = 10) -> list[dict]:
    """OpenAlex 语义检索，返回 [{"title","authors","year","doi","url","abstract"}]。"""
    params = urllib.parse.urlencode({
        "search": query, "per-page": str(limit),
        "select": "title,authorships,publication_year,doi,abstract_inverted_index",
    })
    url = f"{_OPENALEX_BASE}?{params}"
    _logger.info("OpenAlex 检索：%s", url)
    data = _safe_get_json(url)
    if not data or not isinstance(data.get("results"), list):
        return []
    out = []
    for item in data["results"][:limit]:
        authors = []
        for au in (item.get("authorships") or [])[:5]:
            a = au.get("author", {})
            name = a.get("display_name", "")
            if name:
                authors.append(name)
        # 反转倒排索引得到摘要文本
        abstract = ""
        inv = item.get("abstract_inverted_index") or {}
        if inv:
            positions = []
            for word, idxs in inv.items():
                for idx in idxs:
                    positions.append((idx, word))
            positions.sort()
            abstract = " ".join(w for _, w in positions)[:500]
        doi = item.get("doi") or ""
        out.append({
            "title": item.get("title") or "",
            "authors": authors,
            "year": str(item.get("publication_year") or ""),
            "doi": doi.replace("https://doi.org/", "") if doi else "",
            "url": doi or f"https://openalex.org/{item.get('id', '')}",
            "abstract": abstract,
        })
    return out


def search_semantic_scholar(query: str, limit: int = 10) -> list[dict]:
    """Semantic Scholar 检索，返回同结构列表。"""
    params = urllib.parse.urlencode({
        "query": query, "limit": str(limit),
        "fields": "title,authors,year,externalIds,abstract,url",
    })
    url = f"{_S2_BASE}?{params}"
    _logger.info("Semantic Scholar 检索：%s", url)
    data = _safe_get_json(url)
    if not data or not isinstance(data.get("data"), list):
        return []
    out = []
    for item in data["data"][:limit]:
        authors = [a.get("name", "") for a in (item.get("authors") or [])[:5]
                   if a.get("name")]
        ext = item.get("externalIds") or {}
        doi = ext.get("DOI", "")
        out.append({
            "title": item.get("title") or "",
            "authors": authors,
            "year": str(item.get("year") or ""),
            "doi": doi,
            "url": item.get("url") or (f"https://doi.org/{doi}" if doi else ""),
            "abstract": (item.get("abstract") or "")[:500],
        })
    return out


def search_literature(query: str, source: str = "openalex",
                     limit: int = 10) -> list[dict]:
    """统一入口：按 source 分发到 OpenAlex 或 Semantic Scholar。

    source: "openalex"（默认）| "s2" | "both"
    """
    if source == "both":
        return search_openalex(query, limit) + search_semantic_scholar(query, limit)
    if source == "s2":
        return search_semantic_scholar(query, limit)
    return search_openalex(query, limit)


def results_to_cards(results: list[dict]) -> list[dict]:
    """将检索结果转为 synthesizer 可消费的摘要卡格式。"""
    cards = []
    for i, r in enumerate(results):
        cards.append({
            "title": r.get("title") or f"检索结果{i+1}",
            "authors": r.get("authors") or [],
            "year": r.get("year") or "",
            "type": "J",
            "topic": r.get("abstract", "")[:100],
            "method": "",
            "conclusion": "",
            "quotes": [],
            "source": f"文献检索:{r.get('url', '')}",
            "doi": r.get("doi") or "",
            "url": r.get("url") or "",
            "_local_meta": {
                "title": r.get("title") or "",
                "authors": r.get("authors") or [],
                "year": r.get("year") or "",
                "doi": r.get("doi") or "",
                "source": f"文献检索:{r.get('url', '')}",
            },
            "_from_search": True,
        })
    return cards
