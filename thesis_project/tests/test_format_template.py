# -*- coding: utf-8 -*-
import copy

import pytest

from config.format_spec import PPT_SPEC, WORD_SPEC
from config.template import apply_template


def test_external_template_deep_overrides(tmp_path):
    old_w, old_p = copy.deepcopy(WORD_SPEC), copy.deepcopy(PPT_SPEC)
    path = tmp_path / "format.yml"
    path.write_text("word:\n  page:\n    orientation: landscape\nppt:\n  sizes:\n    body_pt: 22\n",
                    encoding="utf-8")
    try:
        apply_template(str(path))
        assert WORD_SPEC["page"]["orientation"] == "landscape"
        assert PPT_SPEC["sizes"]["body_pt"] == 22
    finally:
        WORD_SPEC.clear()
        WORD_SPEC.update(old_w)
        PPT_SPEC.clear()
        PPT_SPEC.update(old_p)


def test_external_template_rejects_unknown_field(tmp_path):
    path = tmp_path / "bad.yml"
    path.write_text("word:\n  mystery: 1\n", encoding="utf-8")
    with pytest.raises(ValueError, match="未知字段"):
        apply_template(str(path))
