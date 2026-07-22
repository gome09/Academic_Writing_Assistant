# -*- coding: utf-8 -*-
import pytest
from src import llm_enhancer


def test_parse_json_plain():
    assert llm_enhancer._parse_json('{"a": 1}') == {"a": 1}


def test_parse_json_fenced_with_prose():
    text = '好的，如下：\n```json\n{"a": [1, 2]}\n```\n以上。'
    assert llm_enhancer._parse_json(text) == {"a": [1, 2]}


def test_parse_json_no_json_raises():
    with pytest.raises(ValueError):
        llm_enhancer._parse_json("抱歉，我无法处理。")
