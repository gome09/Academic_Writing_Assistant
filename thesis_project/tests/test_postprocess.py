# -*- coding: utf-8 -*-
import sys
import types
from unittest.mock import MagicMock

from src import postprocess


def test_refresh_fields_graceful_without_pywin32(monkeypatch, tmp_path, capsys):
    monkeypatch.setitem(sys.modules, "win32com", None)
    monkeypatch.setitem(sys.modules, "win32com.client", None)
    ok = postprocess.refresh_word_fields(str(tmp_path / "x.docx"))
    assert ok is False
    assert "pywin32" in capsys.readouterr().out


def test_refresh_fields_graceful_when_word_missing(monkeypatch, tmp_path, capsys):
    fake_client = types.ModuleType("win32com.client")

    def boom(name):
        raise OSError("Word not installed")

    fake_client.DispatchEx = boom
    fake_pkg = types.ModuleType("win32com")
    fake_pkg.client = fake_client
    monkeypatch.setitem(sys.modules, "win32com", fake_pkg)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    ok = postprocess.refresh_word_fields(str(tmp_path / "x.docx"))
    assert ok is False
    assert "F9" in capsys.readouterr().out


def test_quit_called_when_open_fails(monkeypatch, tmp_path, capsys):
    word = MagicMock()
    word.Documents.Open.side_effect = OSError("locked")
    fake_client = types.ModuleType("win32com.client")
    fake_client.DispatchEx = lambda name: word
    fake_pkg = types.ModuleType("win32com")
    fake_pkg.client = fake_client
    monkeypatch.setitem(sys.modules, "win32com", fake_pkg)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    ok = postprocess.refresh_word_fields(str(tmp_path / "x.docx"))
    assert ok is False
    word.Quit.assert_called_once()


# ---------------------------------------------------------------------------
#  T2-4: PPT → PDF 导出
# ---------------------------------------------------------------------------
def test_export_pptx_pdf_graceful_without_pywin32(monkeypatch, tmp_path, capsys):
    monkeypatch.setitem(sys.modules, "win32com", None)
    monkeypatch.setitem(sys.modules, "win32com.client", None)
    result = postprocess.export_pptx_to_pdf(str(tmp_path / "x.pptx"))
    assert result is None
    assert "pywin32" in capsys.readouterr().out


def test_export_pptx_pdf_graceful_when_powerpoint_missing(monkeypatch, tmp_path, capsys):
    fake_client = types.ModuleType("win32com.client")
    fake_client.DispatchEx = lambda name: (_ for _ in ()).throw(
        OSError("PowerPoint not installed"))
    fake_pkg = types.ModuleType("win32com")
    fake_pkg.client = fake_client
    monkeypatch.setitem(sys.modules, "win32com", fake_pkg)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    result = postprocess.export_pptx_to_pdf(str(tmp_path / "x.pptx"))
    assert result is None
    assert "PPT PDF 导出失败" in capsys.readouterr().out


def test_export_pptx_pdf_success(monkeypatch, tmp_path, capsys):
    """模拟 PowerPoint COM 成功导出 PDF。"""
    prs = MagicMock()
    ppt_app = MagicMock()
    ppt_app.Presentations.Open.return_value = prs
    fake_client = types.ModuleType("win32com.client")
    fake_client.DispatchEx = lambda name: ppt_app
    fake_pkg = types.ModuleType("win32com")
    fake_pkg.client = fake_client
    monkeypatch.setitem(sys.modules, "win32com", fake_pkg)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    result = postprocess.export_pptx_to_pdf(str(tmp_path / "x.pptx"))
    assert result is not None
    assert result.endswith(".pdf")
    prs.SaveAs.assert_called_once()
    prs.Close.assert_called_once()
    ppt_app.Quit.assert_called_once()
    assert "PPT PDF" in capsys.readouterr().out
