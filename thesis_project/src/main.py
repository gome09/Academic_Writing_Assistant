# -*- coding: utf-8 -*-
"""
主管道：读取 input/ 下所有源文件 -> 整理 -> 生成 Word 草案 + PPT 草案。

用法：
    python src/main.py                      # 读 input/，输出到 output/
    python src/main.py --input 某目录        # 指定输入目录
    python src/main.py --input a.pdf b.md    # 指定若干文件
    python src/main.py --only word           # 只生成 Word
    python src/main.py --only ppt            # 只生成 PPT
"""
from __future__ import annotations
import argparse
import os
import sys

# Windows 控制台默认 GBK，重新配置为 UTF-8 以正确输出中文与符号
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.readers import read_file, read_dir
from src.organizer import organize
from src import docx_builder, pptx_builder

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_INPUT = os.path.join(ROOT, "input")
DEFAULT_OUTPUT = os.path.join(ROOT, "output")


def gather_docs(inputs):
    docs = []
    for item in inputs:
        if os.path.isdir(item):
            print(f"[目录] {item}")
            docs.extend(read_dir(item))
        elif os.path.isfile(item):
            try:
                docs.append(read_file(item))
                print(f"  [读取] {os.path.basename(item)}")
            except Exception as e:  # noqa: BLE001
                print(f"  [跳过] {item}: {e}")
        else:
            print(f"  [忽略] 不存在：{item}")
    return docs


def main():
    ap = argparse.ArgumentParser(description="论文Word草案 + 答辩PPT草案 生成器")
    ap.add_argument("--input", nargs="+", default=[DEFAULT_INPUT],
                    help="输入目录或文件（可多个）")
    ap.add_argument("--output", default=DEFAULT_OUTPUT, help="输出目录")
    ap.add_argument("--only", choices=["word", "ppt"], help="只生成其中一种")
    args = ap.parse_args()

    os.makedirs(args.output, exist_ok=True)

    print("=" * 56)
    print("① 读取源文件")
    docs = gather_docs(args.input)
    if not docs:
        print("\n⚠ 未读取到任何文件。请把 Word/PDF/TXT/md/json 放进 input/ 后重试。")
        print("  （可先用示例：--input sample_input）")
        return 1
    print(f"  共读取 {len(docs)} 个文件，"
          f"{sum(len(d['blocks']) for d in docs)} 个内容块。")

    print("② 整理内容结构")
    thesis, deck = organize(docs)
    print(f"  论文：{len(thesis['chapters'])} 章；PPT：{len(deck['slides'])} 页。")

    print("③ 生成草案")
    if args.only != "ppt":
        wp = os.path.join(args.output, "论文草案.docx")
        docx_builder.build(thesis, wp)
        print(f"  ✔ Word: {wp}")
    if args.only != "word":
        pp = os.path.join(args.output, "答辩PPT草案.pptx")
        pptx_builder.build(deck, pp)
        print(f"  ✔ PPT : {pp}")

    print("=" * 56)
    print("完成。请打开草案：Word 中按 F9 更新目录；两份均需人工润色占位符 <请填写>。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
