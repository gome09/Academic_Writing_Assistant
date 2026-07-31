# -*- coding: utf-8 -*-
from src import main as main_mod


def test_build_with_retry_renames_when_locked(tmp_path):
    calls = []

    def fake_build(data, path):
        calls.append(path)
        if len(calls) == 1:
            raise PermissionError(13, "file in use")
        return path

    out = str(tmp_path / "论文草案.docx")
    result = main_mod._build_with_retry(fake_build, {}, out)
    assert result.endswith("论文草案(2).docx")
    assert len(calls) == 2


def test_build_with_retry_gives_up_after_all_locked(tmp_path, capsys):
    def fake_build(data, path):
        raise PermissionError(13, "file in use")

    result = main_mod._build_with_retry(fake_build, {}, str(tmp_path / "a.docx"))
    assert result is None
    assert "占用" in capsys.readouterr().out


def test_draft_completion_only_mentions_generated_ppt(capsys):
    main_mod._print_draft_completion(["output/答辩PPT草案.pptx"])

    out = capsys.readouterr().out
    assert "已生成 PPT" in out
    assert "Word 打开" not in out
    assert "两份" not in out


def test_draft_completion_mentions_word_field_refresh(capsys):
    main_mod._print_draft_completion(["output/论文草案.docx"])

    out = capsys.readouterr().out
    assert "已生成 Word" in out
    assert "F9" in out
