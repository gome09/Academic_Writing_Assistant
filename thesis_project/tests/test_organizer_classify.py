# -*- coding: utf-8 -*-
"""_classify 关键词增强：环境/验证 归 result，LLM 标题归一化。"""
import unicodedata

from src.organizer import _classify


def test_classify_environment_to_result():
    assert _classify("实验环境") == "result"


def test_classify_validation_to_result():
    # 「验证」未在当前 result 桶关键词里，应归 result（新增关键词后）
    assert _classify("实验验证") == "result"


def test_classify_default_method():
    assert _classify("随笔") == "method"


def test_classify_demand_to_method():
    assert _classify("需求分析") == "method"


def test_classify_background_keeps():
    assert _classify("绪论") == "background"


def test_llm_normalize_title_for_lookup():
    s = "实验环境"
    norm = unicodedata.normalize("NFKC", s).strip().replace(" ", "").replace("\u3000", "")
    assert _classify(norm) == "result"
