# -*- coding: utf-8 -*-
from config.format_spec import REFS_SPEC


def test_refs_spec_has_required_keys():
    assert set(REFS_SPEC) >= {"topic_filenames", "default_outline",
                              "review_notice", "materials_chapter"}


def test_topic_filenames_lowercase():
    # main.py 用 lower() 比对，配置里必须全小写
    assert all(n == n.lower() for n in REFS_SPEC["topic_filenames"])


def test_default_outline_kinds_valid():
    kinds = {c["kind"] for c in REFS_SPEC["default_outline"]}
    assert kinds <= {"intro", "review", "core", "conclusion"}
    assert "review" in kinds and "core" in kinds
