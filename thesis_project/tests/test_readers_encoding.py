# -*- coding: utf-8 -*-
import json

from src.readers import read_txt, read_md, read_json, _read_text


def test_read_txt_gbk(tmp_path):
    f = tmp_path / "gbk.txt"
    f.write_bytes("这是一段中文测试文本。".encode("gbk"))
    d = read_txt(str(f))
    assert d["blocks"][0]["text"] == "这是一段中文测试文本。"


def test_read_txt_utf8_bom(tmp_path):
    f = tmp_path / "bom.txt"
    f.write_bytes("你好世界。".encode("utf-8-sig"))
    d = read_txt(str(f))
    assert d["blocks"][0]["text"] == "你好世界。"


def test_read_md_gbk_heading(tmp_path):
    f = tmp_path / "gbk.md"
    f.write_bytes("# 绪论\n\n正文内容。".encode("gbk"))
    d = read_md(str(f))
    assert d["blocks"][0]["text"] == "绪论"
    assert d["blocks"][0]["kind"] == "heading"


def test_read_text_invalid_bytes_no_crash(tmp_path):
    f = tmp_path / "bad.txt"
    f.write_bytes(b"\xff\xfe\x00invalid\x80")
    # 不抛异常即可（最后兜底 errors="replace"）
    assert isinstance(_read_text(str(f)), str)


def test_read_json_gbk(tmp_path):
    f = tmp_path / "gbk.json"
    data = {"title": "毕业论文大纲", "content": "第一章：绪论内容。"}
    f.write_bytes(json.dumps(data, ensure_ascii=False).encode("gbk"))
    d = read_json(str(f))
    headings = [b["text"] for b in d["blocks"] if b["kind"] == "heading"]
    paras = [b["text"] for b in d["blocks"] if b["kind"] == "paragraph"]
    assert "毕业论文大纲" in headings
    assert "第一章：绪论内容。" in paras
