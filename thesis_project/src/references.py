# -*- coding: utf-8 -*-
"""Offline-first reference metadata, GB/T 7714 formatting and citation checks."""
from __future__ import annotations

import json
import logging
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request

_logger = logging.getLogger("thesis_project")

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


def _detect_type(card: dict, local: dict | None = None) -> str:
    """启发式文献类型识别（T1-1）：有出版社→M、有会议名→C、有学位授予单位→D、默认→J。

    仅依据已有字段推断，不调用 LLM 不编造；card.type 优先（来自 LLM 或上游显式指定）。
    """
    local = local or {}
    if card.get("type"):
        return str(card["type"]).strip().upper()[:1] or "J"
    # 学位论文：有学位授予单位/学校字段
    for key in ("degree_grantor", "university", "school", "institution"):
        if card.get(key) or local.get(key):
            return "D"
    # 会议论文：有会议名
    for key in ("conference", "meeting", "proceedings"):
        if card.get(key) or local.get(key):
            return "C"
    # 专著：有出版社但无期刊名
    if (card.get("publisher") or local.get("publisher")) and not (
            card.get("journal") or local.get("journal")):
        return "M"
    return "J"


def entry_from_card(card: dict, local: dict | None = None) -> dict:
    local = local or {}
    return {
        "title": local.get("title") or card.get("title") or card.get("source", ""),
        "authors": list(local.get("authors") or card.get("authors") or []),
        "year": local.get("year") or card.get("year") or "",
        "type": _detect_type(card, local),
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


# ---------------------------------------------------------------------------
#  T1-2：多样式参考文献著录（APA / MLA / Chicago）
#  全部由本地代码确定性生成，不交给 LLM 排版（红线：格式确定性）。
# ---------------------------------------------------------------------------
def _fmt_authors_apa(authors: list) -> str:
    """APA 作者格式：Last, F. M.，多作者用 & 连接末位。"""
    out = []
    for a in authors:
        if not a:
            continue
        parts = a.replace(",", " ").split()
        if len(parts) == 1:
            out.append(parts[0])
        else:
            last = parts[0]
            initials = ". ".join(p[0] for p in parts[1:] if p) + "."
            out.append(f"{last}, {initials}")
    if not out:
        return "«请补全作者»"
    if len(out) == 1:
        return out[0]
    return ", ".join(out[:-1]) + ", & " + out[-1]


def format_apa(entry: dict) -> str:
    """APA 7th：Author. (Year). Title. Journal, Vol(Iss), Pages. DOI"""
    authors = _fmt_authors_apa(entry.get("authors", []))
    year = entry.get("year") or "«请补全年份»"
    title = entry.get("title") or "«请补全题名»"
    parts = [f"{authors} ({year}). {title}."]
    if entry.get("journal"):
        j = entry["journal"]
        vi = ""
        if entry.get("volume"):
            vi += str(entry["volume"])
            if entry.get("issue"):
                vi += f"({entry['issue']})"
        if entry.get("pages"):
            vi += f", {entry['pages']}" if vi else str(entry["pages"])
        parts.append(f"{j}" + (f", {vi}" if vi else "") + ".")
    elif entry.get("publisher"):
        parts.append(f"{entry['publisher']}.")
    if entry.get("doi"):
        parts.append(f"https://doi.org/{entry['doi']}")
    return " ".join(parts)


def format_mla(entry: dict) -> str:
    """MLA 9th：Author. "Title." Journal, vol. V, no. I, Year, pp. Pages."""
    authors = entry.get("authors", [])
    author_str = ", ".join(a for a in authors if a) or "«请补全作者»"
    title = entry.get("title") or "«请补全题名»"
    parts = [f'{author_str}. "{title}."']
    seg = []
    if entry.get("journal"):
        seg.append(entry["journal"])
    if entry.get("volume"):
        seg.append(f"vol. {entry['volume']}")
    if entry.get("issue"):
        seg.append(f"no. {entry['issue']}")
    if entry.get("year"):
        seg.append(str(entry["year"]))
    if entry.get("pages"):
        seg.append(f"pp. {entry['pages']}")
    if seg:
        parts.append(", ".join(seg) + ".")
    if entry.get("publisher") and not entry.get("journal"):
        parts.append(f"{entry['publisher']}.")
    if entry.get("doi"):
        parts.append(f"https://doi.org/{entry['doi']}.")
    return " ".join(parts)


def format_chicago(entry: dict) -> str:
    """Chicago Author-Date：Author. Year. "Title." Journal Vol (Iss): Pages."""
    authors = entry.get("authors", [])
    author_str = ", ".join(a for a in authors if a) or "«请补全作者»"
    year = entry.get("year") or "«请补全年份»"
    title = entry.get("title") or "«请补全题名»"
    parts = [f'{author_str}. {year}. "{title}."']
    if entry.get("journal"):
        loc = entry["journal"]
        if entry.get("volume"):
            loc += f" {entry['volume']}"
            if entry.get("issue"):
                loc += f" ({entry['issue']})"
        if entry.get("pages"):
            loc += f": {entry['pages']}"
        parts.append(loc + ".")
    elif entry.get("publisher"):
        parts.append(f"{entry['publisher']}.")
    if entry.get("doi"):
        parts.append(f"https://doi.org/{entry['doi']}")
    return " ".join(parts)


_FORMATTERS = {
    "GB/T 7714": format_gbt,
    "APA": format_apa,
    "MLA": format_mla,
    "Chicago": format_chicago,
}


def format_reference(entry: dict, standard: str = "GB/T 7714") -> str:
    """按指定著录样式格式化文献条目（T1-2 分发器）。

    确定性本地生成，不调用 LLM。未知样式回退 GB/T 7714。
    """
    fn = _FORMATTERS.get(standard, format_gbt)
    return fn(entry)


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


def check_orphan_references(all_paragraphs: list[str], total_refs: int) -> list[str]:
    """T1-5：检查是否有文献从未被引用（孤立文献），返回告警列表。

    遍历所有综述段落的 [n] 引用，汇总被引编号集合，
    与 1..total_refs 比对，未被引的文献给出告警。
    """
    if total_refs <= 0:
        return []
    cited: set[int] = set()
    for para in all_paragraphs or []:
        for group in _CITE_RE.findall(para or ""):
            ids = [int(x) for x in re.split(r"\s*[,，]\s*", group)]
            cited.update(ids)
    orphans = [i for i in range(1, total_refs + 1) if i not in cited]
    return [f"文献[{i}]在正文中从未被引用（孤立文献）" for i in orphans]


def lookup_crossref(title: str, doi: str = "", cache_path: str | None = None) -> dict:
    """Optional explicit Crossref lookup; never called by default.

    T1-3：网络异常（超时/URLError/HTTPError/JSON 错误）不中断流程，
    指数退避重试 1 次后仍失败则返回空字典（降级为跳过）。
    """
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
    _logger.info("Crossref 查询：%s", url)
    req = urllib.request.Request(url, headers={"User-Agent": "thesis-project/1.0"})

    data = None
    for attempt in range(2):  # 最多 2 次（初始 + 1 次退避重试）
        try:
            with urllib.request.urlopen(req, timeout=10) as response:  # noqa: S310
                data = json.load(response)
            break
        except (urllib.error.URLError, urllib.error.HTTPError,
                OSError, ValueError) as exc:
            if attempt == 0:
                _logger.warning("Crossref 查询失败（%s），1s 后重试：%s",
                                type(exc).__name__, exc)
                time.sleep(1)
                continue
            _logger.warning("Crossref 查询重试仍失败，跳过：%s", exc)
            return {}

    if data is None:
        return {}
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
