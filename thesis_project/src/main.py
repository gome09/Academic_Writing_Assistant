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
    """读取所有输入，返回 (docs, errors)。

    errors 是 [(path, reason), ...]，调用方在 main() 末尾聚合打印。
    """
    docs = []
    errors = []
    for item in inputs:
        if os.path.isdir(item):
            print(f"[目录] {item}")
            for d in read_dir(item):
                docs.append(d)
        elif os.path.isfile(item):
            try:
                docs.append(read_file(item))
                print(f"  [读取] {os.path.basename(item)}")
            except Exception as e:  # noqa: BLE001
                errors.append((item, str(e)))
                print(f"  [跳过] {item}: {e}")
        else:
            errors.append((item, '不存在'))
            print(f"  [忽略] 不存在：{item}")
    if errors:
        print(f"  [跳过 {len(errors)} 个文件]")
        for p_, reason in errors:
            print(f"    - {p_}: {reason}")
    return docs, errors


def _build_with_retry(build_fn, data, out_path):
    """构建并保存；文件被 Word/WPS 占用时自动加序号改名重试，全部失败返回 None。"""
    candidates = [out_path]
    root, ext = os.path.splitext(out_path)
    candidates += [f"{root}({i}){ext}" for i in range(2, 6)]
    for i, path in enumerate(candidates):
        try:
            result = build_fn(data, path)
            if i > 0:
                print(f"  [提示] {os.path.basename(out_path)} 正被占用"
                      f"（可能在 Word/WPS 中打开），已改存：{os.path.basename(path)}")
            return result
        except PermissionError:
            continue
    print(f"  [错误] 无法写入 {out_path}：文件被占用。请关闭 Word/WPS 后重试。")
    return None


def main():
    ap = argparse.ArgumentParser(description="论文Word草案 + 答辩PPT草案 生成器")
    ap.add_argument("--input", nargs="+", default=[DEFAULT_INPUT],
                    help="输入目录或文件（可多个）")
    ap.add_argument("--output", default=DEFAULT_OUTPUT, help="输出目录")
    ap.add_argument("--only", choices=["word", "ppt"], help="只生成其中一种")
    ap.add_argument("--llm", action="store_true",
                    help="用 LLM 增强草案质量（需设置 LLM_API_KEY，"
                         "可选 LLM_BASE_URL / LLM_MODEL，OpenAI 兼容接口）")
    ap.add_argument("--refresh-fields", action="store_true",
                    help="生成后用本机 Word 静默刷新目录/页码域（需已装 Word 和 pywin32）")
    ap.add_argument("--pdf", action="store_true",
                    help="刷新域后同时导出 PDF（隐含 --refresh-fields）")
    args = ap.parse_args()

    os.makedirs(args.output, exist_ok=True)

    print("=" * 56)
    print("① 读取源文件")
    docs, errors = gather_docs(args.input)
    if not docs:
        print("\n⚠ 未读取到任何文件。请把 Word/PDF/TXT/md/json 放进 input/ 后重试。")
        print("  （可先用示例：--input sample_input）")
        return 1
    print(f"  共读取 {len(docs)} 个文件，"
          f"{sum(len(d['blocks']) for d in docs)} 个内容块。")

    print("② 整理内容结构")
    thesis, deck = organize(docs)
    if args.llm:
        print("②+ LLM 增强")
        from src import llm_enhancer
        thesis, deck = llm_enhancer.enhance(thesis, deck, docs)
    print(f"  论文：{len(thesis['chapters'])} 章；PPT：{len(deck['slides'])} 页。")

    print("③ 生成草案")
    ok = True
    if args.only != "ppt":
        wp = os.path.join(args.output, "论文草案.docx")
        wp = _build_with_retry(docx_builder.build, thesis, wp)
        if wp:
            print(f"  ✔ Word: {wp}")
            if args.refresh_fields or args.pdf:
                from src import postprocess
                postprocess.refresh_word_fields(wp, export_pdf=args.pdf)
        else:
            ok = False
    if args.only != "word":
        pp = os.path.join(args.output, "答辩PPT草案.pptx")
        pp = _build_with_retry(pptx_builder.build, deck, pp)
        if pp:
            print(f"  ✔ PPT : {pp}")
        else:
            ok = False

    print("=" * 56)
    print("完成。请打开草案：Word 打开时按提示更新域（或按 F9 更新目录）；两份均需人工润色占位符 <请填写>。")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
