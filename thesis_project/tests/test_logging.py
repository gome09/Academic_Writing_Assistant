# -*- coding: utf-8 -*-
"""T0-3: 日志统一验收——关键事件落盘，含降级步骤。"""
from __future__ import annotations

import logging

from src import synthesizer
from src.logging_setup import configure_logging


def test_note_degrade_writes_to_log_file(tmp_path):
    """_note_degrade 应写入日志文件，且含「降级」关键字。"""
    log_path = str(tmp_path / "运行日志.log")
    configure_logging(log_path)
    degraded = []
    synthesizer._note_degrade("摘要卡", RuntimeError("模拟失败"), degraded)
    assert degraded == ["摘要卡"]
    content = open(log_path, "r", encoding="utf-8").read()
    assert "降级" in content
    assert "摘要卡" in content
    logging.getLogger("thesis_project").handlers.clear()


def test_log_file_nonempty_after_degradation(tmp_path):
    """端到端：配置日志后调用会降级的流程，日志文件非空。"""
    log_path = str(tmp_path / "运行日志.log")
    configure_logging(log_path)
    # 触发一次降级记录
    synthesizer._note_degrade("测试步骤", ValueError("err"), [])
    content = open(log_path, "r", encoding="utf-8").read()
    assert len(content) > 0
    assert "WARNING" in content
    logging.getLogger("thesis_project").handlers.clear()
