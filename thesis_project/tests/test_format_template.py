# -*- coding: utf-8 -*-
"""T0-2/T0-6: 配置模板不可变性与校验覆盖。

- apply_template 始终从原始默认值合并，多次调用不互相污染（T0-2）
- _validate 覆盖类型错误与负数分支（T0-6）
"""
import pytest

from config.format_spec import PPT_SPEC, WORD_SPEC
from config.template import apply_template, _validate


def _write(path, text):
    path.write_text(text, encoding="utf-8")


def test_external_template_deep_overrides(tmp_path):
    """apply_template 后不再需要 finally 还原——默认值快照保证不污染。"""
    path = tmp_path / "format.yml"
    _write(path, "word:\n  page:\n    orientation: landscape\nppt:\n  sizes:\n    body_pt: 22\n")
    apply_template(str(path))
    assert WORD_SPEC["page"]["orientation"] == "landscape"
    assert PPT_SPEC["sizes"]["body_pt"] == 22
    # 未覆盖的字段保持默认值
    assert WORD_SPEC["page"]["margin_top_cm"] == 2.5
    assert PPT_SPEC["sizes"]["cover_title_pt"] == 40


def test_external_template_rejects_unknown_field(tmp_path):
    path = tmp_path / "bad.yml"
    _write(path, "word:\n  mystery: 1\n")
    with pytest.raises(ValueError, match="未知字段"):
        apply_template(str(path))


def test_repeated_apply_does_not_contaminate(tmp_path):
    """T0-2 核心：两次 apply_template 不会从前次修改后的状态合并。"""
    p1 = tmp_path / "a.yml"
    _write(p1, "word:\n  page:\n    margin_top_cm: 5.0\n")
    apply_template(str(p1))
    assert WORD_SPEC["page"]["margin_top_cm"] == 5.0
    # 第二次用一个不同的字段，第一次修改的 margin_top 不应泄漏进基准
    p2 = tmp_path / "b.yml"
    _write(p2, "ppt:\n  sizes:\n    body_pt: 20\n")
    apply_template(str(p2))
    # 第一次覆盖的 margin_top 已不在新模板中 -> 应恢复为默认值 2.5
    assert WORD_SPEC["page"]["margin_top_cm"] == 2.5
    assert PPT_SPEC["sizes"]["body_pt"] == 20


def test_validate_rejects_non_dict_root():
    with pytest.raises(ValueError, match="根节点必须是对象"):
        _validate([1, 2, 3])


def test_validate_rejects_non_dict_word():
    with pytest.raises(ValueError, match="word 模板必须是对象"):
        _validate({"word": "not-a-dict"})


def test_validate_rejects_non_dict_ppt():
    with pytest.raises(ValueError, match="ppt 模板必须是对象"):
        _validate({"ppt": 42})


def test_validate_rejects_unknown_root_field():
    with pytest.raises(ValueError, match="未知根字段"):
        _validate({"surprise": {}})


def test_validate_rejects_nested_non_dict():
    with pytest.raises(ValueError, match="必须是对象"):
        _validate({"word": {"page": "not-a-dict"}})


def test_validate_rejects_negative_number(tmp_path):
    path = tmp_path / "neg.yml"
    _write(path, "word:\n  page:\n    margin_top_cm: -1\n")
    with pytest.raises(ValueError, match="不能为负数"):
        apply_template(str(path))


def test_validate_rejects_type_mismatch():
    with pytest.raises(ValueError, match="类型不正确"):
        _validate({"word": {"page": {"margin_top_cm": "abc"}}})


def test_validate_rejects_unknown_nested_field():
    with pytest.raises(ValueError, match="未知字段"):
        _validate({"word": {"page": {"no_such_field": 1}}})


def test_empty_template_is_noop(tmp_path):
    """空 YAML -> 保持全部默认值。"""
    path = tmp_path / "empty.yml"
    _write(path, "")
    apply_template(str(path))
    assert WORD_SPEC["page"]["margin_top_cm"] == 2.5
    assert PPT_SPEC["sizes"]["body_pt"] == 24


def test_defaults_snapshot_independent():
    """默认值快照与 live spec 是不同对象（深拷贝）。"""
    from config.format_spec import _WORD_DEFAULTS
    assert _WORD_DEFAULTS is not WORD_SPEC
    assert _WORD_DEFAULTS["page"]["margin_top_cm"] == 2.5
