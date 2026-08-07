# -*- coding: utf-8 -*-
"""T0-1: 外发确认不能被 PYTEST_CURRENT_TEST 在生产环境绕过。

pyest 会向子进程注入 PYTEST_CURRENT_TEST，但该变量也能被手动设置。
确认逻辑必须额外校验 pytest 确实已加载（sys.modules），否则不应豁免。
"""
from __future__ import annotations

import sys

from src import main as main_mod


class _Args:
    def __init__(self, yes=False, llm=False, polish=None):
        self.yes = yes
        self.llm = llm
        self.polish = polish


def _doc(source="input/a.md"):
    return {"source": source, "type": "md", "meta": {},
            "blocks": [{"kind": "paragraph", "text": "x"}]}


def test_consent_yes_bypasses_prompt(monkeypatch):
    """--yes 仍然在非交互环境豁免（既有行为，保持）。"""
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.delenv("PYTEST_CURRENT_TEST", raising=False)
    monkeypatch.delenv("THESIS_LLM_CONSENT", raising=False)
    monkeypatch.setattr(sys, "stdin", type("S", (), {"isatty": lambda self: False})())
    assert main_mod._confirm_llm_transfer(_Args(yes=True, llm=True), [_doc()]) is True


def test_consent_env_token_bypasses(monkeypatch):
    """THESIS_LLM_CONSENT=1 豁免（既有行为，保持）。"""
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("THESIS_LLM_CONSENT", "1")
    monkeypatch.delenv("PYTEST_CURRENT_TEST", raising=False)
    monkeypatch.setattr(sys, "stdin", type("S", (), {"isatty": lambda self: False})())
    assert main_mod._confirm_llm_transfer(_Args(llm=True), [_doc()]) is True


def test_consent_rejects_noninteractive_without_yes(monkeypatch, capsys):
    """非交互、无 --yes、无 token -> 拒绝。"""
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.delenv("PYTEST_CURRENT_TEST", raising=False)
    monkeypatch.delenv("THESIS_LLM_CONSENT", raising=False)
    fake_stdin = type("S", (), {"isatty": lambda self: False})()
    monkeypatch.setattr(sys, "stdin", fake_stdin)
    assert main_mod._confirm_llm_transfer(_Args(llm=True), [_doc()]) is False
    assert "非交互" in capsys.readouterr().out


def test_pytest_var_alone_does_not_bypass_in_production(monkeypatch, capsys):
    """关键修复：仅设 PYTEST_CURRENT_TEST 但 pytest 未加载（生产模拟）-> 不豁免。

    在测试进程中 pytest 确在 sys.modules，因此需要临时把它移出再还原。
    """
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("PYTEST_CURRENT_TEST", "tests/test_llm_consent.py::test_x")
    monkeypatch.delenv("THESIS_LLM_CONSENT", raising=False)
    # 模拟生产环境：pytest 不在 sys.modules
    saved = sys.modules.pop("pytest", None)
    fake_stdin = type("S", (), {"isatty": lambda self: False})()
    monkeypatch.setattr(sys, "stdin", fake_stdin)
    try:
        assert main_mod._confirm_llm_transfer(_Args(llm=True), [_doc()]) is False
    finally:
        if saved is not None:
            sys.modules["pytest"] = saved
    assert "非交互" in capsys.readouterr().out


def test_pytest_var_bypasses_when_pytest_loaded(monkeypatch):
    """测试进程内 pytest 在 sys.modules 且 PYTEST_CURRENT_TEST 已设 -> 豁免（正常）。"""
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("PYTEST_CURRENT_TEST", "tests/test_llm_consent.py::test_x")
    monkeypatch.delenv("THESIS_LLM_CONSENT", raising=False)
    # pytest 已在 sys.modules（运行中），保持
    assert "pytest" in sys.modules
    assert main_mod._confirm_llm_transfer(_Args(llm=True), [_doc()]) is True


def test_no_api_key_short_circuits(monkeypatch):
    """无 LLM_API_KEY -> 直接 True（无需外发）。"""
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    monkeypatch.delenv("PYTEST_CURRENT_TEST", raising=False)
    monkeypatch.delenv("THESIS_LLM_CONSENT", raising=False)
    assert main_mod._confirm_llm_transfer(_Args(), [_doc()]) is True
