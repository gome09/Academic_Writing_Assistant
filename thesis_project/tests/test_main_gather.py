# -*- coding: utf-8 -*-
"""main.gather_docs 错误聚合：失败项统一汇总，正例不被聚合吞掉。"""
from src import main as main_module


def test_gather_files_individual_logs_skip(monkeypatch, capsys, tmp_path):
    """单文件抛错时，汇总会列出来。"""
    f = tmp_path / "boom.txt"
    f.write_text("hello", encoding="utf-8")

    def boom(path):
        raise RuntimeError("simulated read failure")

    monkeypatch.setattr(main_module, "read_file", boom)

    docs, errors = main_module.gather_docs([str(f)])
    out = capsys.readouterr().out
    assert docs == []
    assert len(errors) == 1
    assert "[跳过 1 个文件]" in out
    assert "boom.txt" in out


def test_gather_files_real_file_passes(tmp_path, capsys):
    """未抛错分支：返回 docs，无汇总。"""
    f = tmp_path / "ok.md"
    f.write_text("# 绪论\n\n正文。", encoding="utf-8")
    docs, errors = main_module.gather_docs([str(f)])
    out = capsys.readouterr().out
    assert len(docs) == 1
    assert errors == []
    assert "[读取]" in out
    assert "[跳过" not in out


def test_gather_docs_dir_missing(capsys):
    """不存在路径也记入 errors，由 main() 端汇总。"""
    docs, errors = main_module.gather_docs(["Z:/__definitely_not_there__"])
    out = capsys.readouterr().out
    assert docs == []
    assert len(errors) == 1
    assert errors[0][1] == "不存在"
    assert "[忽略]" in out
