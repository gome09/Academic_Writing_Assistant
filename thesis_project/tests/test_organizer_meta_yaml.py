# -*- coding: utf-8 -*-
"""_extract_meta 解析 meta.keywords 的 4 种形态。"""
import pytest

from src.organizer import _extract_meta
from tests.factories import h, p, doc


def _docs_with_meta(keywords_value):
    return [doc([h(1, "绪论"), p("内容。")],
                meta={"title": "T", "keywords": keywords_value})]


@pytest.mark.parametrize("raw,expect_first", [
    (["深度学习", "图像分类"], "深度学习"),
    ("深度学习,图像分类", "深度学习"),
    ("深度学习；图像分类", "深度学习"),
    ('["深度学习", "图像分类"]', "深度学习"),
    ("深度学习 图像分类", "深度学习"),
])
def test_keywords_parsing(raw, expect_first):
    docs = _docs_with_meta(raw)
    meta = _extract_meta(docs)
    assert meta["keywords"][0] == expect_first
    # 必须有标签
    assert all(k for k in meta["keywords"])
    assert len(meta["keywords"]) <= 5
