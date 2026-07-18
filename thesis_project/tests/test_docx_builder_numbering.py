# -*- coding: utf-8 -*-
import pytest
from src.docx_builder import _cn_num


@pytest.mark.parametrize("n,expect", [
    (1, "一"), (2, "二"), (9, "九"), (10, "十"), (11, "十一"),
    (19, "十九"), (20, "二十"), (21, "二十一"), (30, "三十"), (31, "三十一"),
    (59, "五十九"), (99, "九十九"),
])
def test_cn_num_basic(n, expect):
    assert _cn_num(n) == expect
