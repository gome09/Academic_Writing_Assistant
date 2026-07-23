# -*- coding: utf-8 -*-
"""Load and validate external YAML formatting overrides."""
from __future__ import annotations

import copy

from config.format_spec import PPT_SPEC, WORD_SPEC


def _merge(base, override):
    out = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = _merge(out[key], value)
        else:
            out[key] = value
    return out


def _validate(spec):
    if not isinstance(spec, dict):
        raise ValueError("格式模板根节点必须是对象")
    unknown = set(spec) - {"word", "ppt"}
    if unknown:
        raise ValueError(f"格式模板包含未知根字段：{', '.join(sorted(unknown))}")
    if "word" in spec and not isinstance(spec["word"], dict):
        raise ValueError("word 模板必须是对象")
    if "ppt" in spec and not isinstance(spec["ppt"], dict):
        raise ValueError("ppt 模板必须是对象")
    _validate_keys(spec.get("word", {}), WORD_SPEC, "word")
    _validate_keys(spec.get("ppt", {}), PPT_SPEC, "ppt")
    for section in (spec.get("word", {}), spec.get("ppt", {})):
        for key, value in section.items():
            if isinstance(value, dict):
                continue
            if isinstance(value, (int, float)) and value < 0:
                raise ValueError(f"模板字段 {key} 不能为负数")


def _validate_keys(override, base, path):
    for key, value in override.items():
        if key not in base:
            raise ValueError(f"格式模板包含未知字段：{path}.{key}")
        expected = base[key]
        if isinstance(expected, dict):
            if not isinstance(value, dict):
                raise ValueError(f"模板字段 {path}.{key} 必须是对象")
            _validate_keys(value, expected, f"{path}.{key}")
        elif not isinstance(value, type(expected)) and not (
                isinstance(expected, float) and isinstance(value, int)):
            raise ValueError(f"模板字段 {path}.{key} 类型不正确")


def apply_template(path: str) -> tuple[dict, dict]:
    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError("外部模板需要 PyYAML：pip install pyyaml") from exc
    with open(path, "r", encoding="utf-8-sig") as fh:
        raw = yaml.safe_load(fh) or {}
    _validate(raw)
    word = _merge(WORD_SPEC, raw.get("word", {}))
    ppt = _merge(PPT_SPEC, raw.get("ppt", {}))
    WORD_SPEC.clear()
    WORD_SPEC.update(word)
    PPT_SPEC.clear()
    PPT_SPEC.update(ppt)
    return WORD_SPEC, PPT_SPEC
