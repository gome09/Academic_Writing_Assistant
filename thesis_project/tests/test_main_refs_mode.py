# -*- coding: utf-8 -*-
import pytest
from src import main as main_mod


def _doc(source, dtype="md"):
    return {"source": source, "type": dtype, "meta": {}, "blocks": []}


def test_split_topic_detects_by_name_case_insensitive():
    docs = [_doc("input/Topic.MD"), _doc("input/a.pdf")]
    topic, refs = main_mod._split_topic(docs)
    assert topic is docs[0]
    assert refs == [docs[1]]


def test_split_topic_chinese_filename():
    docs = [_doc("input/题目.txt"), _doc("input/b.docx")]
    topic, refs = main_mod._split_topic(docs)
    assert topic is docs[0]


def test_split_topic_none_when_absent():
    topic, refs = main_mod._split_topic([_doc("input/a.pdf")])
    assert topic is None and len(refs) == 1


def test_refs_mode_requires_api_key(monkeypatch, tmp_path, capsys):
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    rc = main_mod._run_refs_mode_checks(_doc("input/topic.md"),
                                        [_doc("input/a.pdf")])
    assert rc == 1
    assert "LLM_API_KEY" in capsys.readouterr().out


def test_refs_mode_requires_topic(monkeypatch, capsys):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    rc = main_mod._run_refs_mode_checks(None, [_doc("input/a.pdf")])
    assert rc == 1
    assert "题目文件" in capsys.readouterr().out


def test_refs_mode_requires_references(monkeypatch, capsys):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    rc = main_mod._run_refs_mode_checks(_doc("input/topic.md"), [])
    assert rc == 1
    assert "参考资料" in capsys.readouterr().out


def test_refs_mode_checks_pass(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    rc = main_mod._run_refs_mode_checks(_doc("input/topic.md"),
                                        [_doc("input/a.pdf")])
    assert rc == 0
