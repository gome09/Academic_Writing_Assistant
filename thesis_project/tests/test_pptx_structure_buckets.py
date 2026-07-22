# -*- coding: utf-8 -*-
from src import organizer, pptx_builder


def _chapters():
    def ch(title, n_paras):
        return {"title": title, "level": 1,
                "paras": [f"{title}的第{i}段内容，用于生成要点。" for i in range(n_paras)],
                "subs": []}
    return [ch("研究背景", 2), ch("系统设计", 2), ch("实验结果", 2), ch("总结", 2)]


def test_build_deck_slides_carry_bucket():
    meta = {"title": "t", "author": "a", "abstract": "x", "keywords": ["k"]}
    deck = organizer._build_deck(meta, _chapters())
    content_buckets = {s.get("bucket") for s in deck["slides"]
                       if s["type"] == "content"}
    assert content_buckets <= {"background", "method", "result", "conclusion"}
    assert "background" in content_buckets
    assert all(s.get("bucket") is None for s in deck["slides"]
               if s["type"] in ("cover", "outline", "thanks"))


def test_check_structure_warns_on_bucket_overflow(capsys):
    slides = ([{"type": "cover", "title": "t"},
               {"type": "outline", "title": "目录", "items": []}]
              + [{"type": "content", "title": f"方法{i}", "bullets": ["x"],
                  "bucket": "method"} for i in range(9)]      # 超过 method 上限
              + [{"type": "thanks", "title": "致谢"}])
    pptx_builder._check_structure(slides)
    out = capsys.readouterr().out
    assert "研究方法" in out and "9" in out


def test_check_structure_ok_no_bucket_warning(capsys):
    slides = ([{"type": "cover", "title": "t"},
               {"type": "outline", "title": "目录", "items": []}]
              + [{"type": "content", "title": t, "bullets": ["x"], "bucket": b}
                 for b, t in [("background", "背景"), ("background", "意义"),
                              ("method", "方法1"), ("method", "方法2"),
                              ("method", "方法3"),
                              ("result", "结果1"), ("result", "结果2"),
                              ("result", "结果3"),
                              ("conclusion", "结论")]]
              + [{"type": "thanks", "title": "致谢"}])
    pptx_builder._check_structure(slides)
    out = capsys.readouterr().out
    for seg in ("研究背景", "研究方法", "研究成果", "结论"):
        assert f"「{seg}" not in out or "警告" not in out
