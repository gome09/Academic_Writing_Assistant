# -*- coding: utf-8 -*-
"""Offline-first reference metadata, GB/T 7714 formatting and citation checks."""
from __future__ import annotations

import json
import os
import re
import urllib.parse
import urllib.request

_DOI_RE = re.compile(r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+\b", re.I)
_YEAR_RE = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
_CITE_RE = re.compile(r"\[(\d+(?:\s*[,，]\s*\d+)*)\]")


def _doc_text(doc: dict, limit: int = 12000) -> str:
    parts = []
    for block in doc.get("blocks", []):
        if block.get("text"):
            parts.append(str(block["text"]))
        elif block.get("kind") == "table":
            parts.extend(" | ".join(str(c).strip() for c in row)
                         for row in block.get("rows") or [])
    return "\n".join(parts)[:limit]


def extract_local_metadata(doc: dict) -> dict:
    """Extract bibliographic fields without network or LLM calls."""
    text = _doc_text(doc)
    source = os.path.basename(doc.get("source", ""))
    meta = doc.get("meta") or {}
    title = str(meta.get("title") or "").strip()
    if not title:
        for block in doc.get("blocks", []):
            if block.get("kind") == "heading" and block.get("text"):
                title = str(block["text"]).strip()
                break
    author_match = re.search(r"(?:作者|author(?:s)?)\s*[：:]\s*(.+)", text, re.I)
    year_match = _YEAR_RE.search(text) or _YEAR_RE.search(source)
    doi_match = _DOI_RE.search(text)
    return {
        "title": title,
        "authors": ([str(meta["author"]).strip()] if meta.get("author") else
                    ([author_match.group(1).strip()] if author_match else [])),
        "year": year_match.group(0) if year_match else "",
        "doi": doi_match.group(0).rstrip(".,;") if doi_match else "",
        "source": source,
    }


def entry_from_card(card: dict, local: dict | None = None) -> dict:
    local = local or {}
    return {
        "title": local.get("title") or card.get("title") or card.get("source", ""),
        "authors": list(local.get("authors") or card.get("authors") or []),
        "year": local.get("year") or card.get("year") or "",
        "type": card.get("type") or "J",
        "journal": card.get("journal") or "",
        "publisher": card.get("publisher") or "",
        "volume": card.get("volume") or "",
        "issue": card.get("issue") or "",
        "pages": card.get("pages") or "",
        "doi": local.get("doi") or card.get("doi") or "",
        "url": card.get("url") or "",
        "source": card.get("source") or local.get("source") or "",
    }


def format_gbt(entry: dict) -> str:
    authors = ", ".join(a for a in entry.get("authors", []) if a)
    authors = authors or "«请补全作者»"
    title = entry.get("title") or "«请补全题名»"
    kind = entry.get("type") or "J"
    tail = []
    if entry.get("journal"):
        tail.append(entry["journal"])
    elif kind == "J":
        tail.append("«请补全期刊»")
    if entry.get("publisher"):
        tail.append(entry["publisher"])
    if entry.get("year"):
        tail.append(str(entry["year"]))
    vi = ""
    if entry.get("volume"):
        vi += str(entry["volume"])
    if entry.get("issue"):
        vi += f"({entry['issue']})"
    if vi:
        tail.append(vi)
    if entry.get("pages"):
        tail.append(str(entry["pages"]))
    if entry.get("doi"):
        tail.append("DOI:" + entry["doi"])
    if not tail:
        tail.append("«请补全著录信息»")
    return f"{authors}. {title}[{kind}]. " + ", ".join(tail) + "."


def validate_citations(paragraphs: list[str], allowed_ids: set[int], total: int) -> list[str]:
    issues = []
    found = False
    for pi, para in enumerate(paragraphs, 1):
        for group in _CITE_RE.findall(para or ""):
            found = True
            ids = [int(x) for x in re.split(r"\s*[,，]\s*", group)]
            for ident in ids:
                if ident < 1 or ident > total:
                    issues.append(f"第{pi}段引用[{ident}]超出文献范围")
                elif ident not in allowed_ids:
                    issues.append(f"第{pi}段引用[{ident}]不属于本章关联文献")
    if allowed_ids and not found:
        issues.append("综述存在关联文献但正文没有任何引用")
    return issues


def lookup_crossref(title: str, doi: str = "", cache_path: str | None = None) -> dict:
    """Optional explicit Crossref lookup; never called by default."""
    cache = {}
    if cache_path and os.path.isfile(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as fh:
                cache = json.load(fh)
        except (OSError, ValueError):
            cache = {}
    key = doi or title.strip()
    if not key:
        return {}
    if key in cache:
        return cache[key]
    query = {"query.bibliographic": title, "rows": 1}
    if doi:
        query = {"filter": "doi:" + doi, "rows": 1}
    url = "https://api.crossref.org/works?" + urllib.parse.urlencode(query)
    req = urllib.request.Request(url, headers={"User-Agent": "thesis-project/1.0"})
    with urllib.request.urlopen(req, timeout=10) as response:  # noqa: S310
        data = json.load(response)
    items = data.get("message", {}).get("items", [])
    result = items[0] if items else {}
    if cache_path:
        cache_dir = os.path.dirname(cache_path)
        if cache_dir:
            os.makedirs(cache_dir, exist_ok=True)
        cache[key] = result
        with open(cache_path, "w", encoding="utf-8") as fh:
            json.dump(cache, fh, ensure_ascii=False, indent=2)
    return result
