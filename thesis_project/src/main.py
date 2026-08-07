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
import logging
import os
import sys
from urllib.parse import urlparse

# Windows 控制台默认 GBK，重新配置为 UTF-8 以正确输出中文与符号
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.readers import read_file, read_dir_detailed
from src.organizer import organize
from src import docx_builder, pptx_builder
from config.format_spec import REFS_SPEC

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_INPUT = os.path.join(ROOT, "input")
DEFAULT_OUTPUT = os.path.join(ROOT, "output")

_logger = logging.getLogger("thesis_project")


def gather_docs(inputs):
    """读取所有输入，返回 (docs, errors)。

    errors 是 [(path, reason), ...]，调用方在 main() 末尾聚合打印。
    """
    docs = []
    errors = []
    for item in inputs:
        if os.path.isdir(item):
            print(f"[目录] {item}")
            dir_docs, dir_errors = read_dir_detailed(item)
            docs.extend(dir_docs)
            errors.extend(dir_errors)
        elif os.path.isfile(item):
            try:
                docs.append(read_file(item))
                print(f"  [读取] {os.path.basename(item)}")
            except Exception as e:  # noqa: BLE001
                errors.append((item, str(e)))
                _logger.warning("读取失败 %s: %s", item, e)
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
                _logger.warning("写盘重试 %s -> %s（文件被占用）",
                                os.path.basename(out_path), os.path.basename(path))
                print(f"  [提示] {os.path.basename(out_path)} 正被占用"
                      f"（可能在 Word/WPS 中打开），已改存：{os.path.basename(path)}")
            return result
        except PermissionError:
            continue
    _logger.error("无法写入 %s：文件被占用，全部重试失败", out_path)
    print(f"  [错误] 无法写入 {out_path}：文件被占用。请关闭 Word/WPS 后重试。")
    return None


def _split_topic(docs):
    """按约定文件名拆出题目 Document；返回 (topic_doc|None, 其余docs)。"""
    names = {n.lower() for n in REFS_SPEC["topic_filenames"]}
    topic_doc, refs = None, []
    for d in docs:
        if topic_doc is None and os.path.basename(d["source"]).lower() in names:
            topic_doc = d
        else:
            refs.append(d)
    return topic_doc, refs


def _run_refs_mode_checks(topic_doc, ref_docs) -> int:
    """参考资料模式入口硬校验；通过返回 0，否则打印原因返回 1。"""
    from src import llm_enhancer
    if topic_doc is None:
        print("[错误] 参考资料模式需要题目文件。请在 input/ 放置 "
              "topic.md（或 题目.txt），写明论文题目与研究方向。")
        return 1
    if not llm_enhancer.is_available():
        print("[错误] 参考资料模式需要 LLM。请设置环境变量后重试：\n"
              "    set LLM_API_KEY=sk-...\n"
              "    set LLM_BASE_URL=https://api.deepseek.com   （可选）\n"
              "    set LLM_MODEL=deepseek-chat                 （可选）\n"
              "  （PowerShell 用 $env:LLM_API_KEY=\"sk-...\"）")
        return 1
    if not ref_docs:
        print("[错误] 未发现参考资料。请把文献（PDF/Word/md/txt/json）、"
              "数据（xlsx/csv）或截图放入 input/。")
        return 1
    return 0


def _run_refs_mode(args, topic_doc, ref_docs) -> int:
    """参考资料模式主流程：综合 -> 仅生成 Word。"""
    rc = _run_refs_mode_checks(topic_doc, ref_docs)
    if rc:
        return rc
    from src import synthesizer
    print("② 综合参考资料（LLM）")
    thesis = synthesizer.synthesize(
        topic_doc, ref_docs, lookup_metadata=args.lookup_metadata,
        cache_path=os.path.join(ROOT, ".cache", "reference_metadata.json"))
    print("③ 生成草案（参考资料模式只生成 Word，不生成 PPT）")
    wp = os.path.join(args.output, "论文草案.docx")
    wp = _build_with_retry(docx_builder.build, thesis, wp)
    if not wp:
        return 1
    print(f"  ✔ Word: {wp}")
    if args.refresh_fields or args.pdf:
        from src import postprocess
        postprocess.refresh_word_fields(wp, export_pdf=args.pdf)
    print("=" * 56)
    print("完成。综述正文带 <AI生成，请核对> 标记，请逐条核对文献后改写为"
          "自己的表述；核心章节按【写作要点】撰写；检索 <请填写> 补全占位符。")
    from src.runtime_report import write_report
    write_report(args.report or os.path.join(args.output, "运行报告.json"),
                 {"status": "ok", "mode": "refs", "files": len(ref_docs) + 1,
                  "outputs": [wp], "metadata_lookup": args.lookup_metadata,
                  "references": thesis.get("reference_entries", []),
                  "degraded_steps": thesis.get("degraded_steps", [])})
    return 0


def _confirm_llm_transfer(args, docs, required=False) -> bool:
    if not (args.llm or args.polish or required):
        return True
    if not os.environ.get("LLM_API_KEY"):
        return True
    base = os.environ.get("LLM_BASE_URL") or "https://api.openai.com"
    host = urlparse(base).netloc or base
    chars = sum(len(b.get("text") or "") for d in docs for b in d["blocks"])
    images = sum(b.get("kind") == "image" for d in docs for b in d["blocks"])
    print(f"  [LLM外发确认] 端点：{host}；文件：{len(docs)}；"
          f"文本约 {chars} 字符；图片 {images} 张")
    _logger.info("LLM外发确认：端点=%s 文件=%d 文本=%d字符 图片=%d张",
                 host, len(docs), chars, images)
    if args.yes or os.environ.get("THESIS_LLM_CONSENT") == "1" or (
            os.environ.get("PYTEST_CURRENT_TEST")
            and sys.modules.get("pytest") is not None):
        return True
    if sys.stdin.isatty():
        return input("  是否继续发送？[y/N] ").strip().lower() in ("y", "yes")
    print("  [错误] 非交互环境必须传 --yes 才能发送资料。")
    return False


def _print_llm_transfer_summary(docs):
    base = os.environ.get("LLM_BASE_URL") or "https://api.openai.com"
    host = urlparse(base).netloc or base
    chars = sum(len(b.get("text") or "") for d in docs for b in d["blocks"])
    images = sum(b.get("kind") == "image" for d in docs for b in d["blocks"])
    print(f"  [LLM外发清单] 端点：{host}；文件：{len(docs)}；"
          f"文本约 {chars} 字符；图片 {images} 张")


def _print_draft_completion(outputs):
    kinds = []
    if any(str(path).lower().endswith(".docx") for path in outputs):
        kinds.append("Word")
    if any(str(path).lower().endswith(".pptx") for path in outputs):
        kinds.append("PPT")
    generated = " + ".join(kinds) if kinds else "草案"
    print(f"完成。已生成 {generated}；生成内容需人工润色，占位符 <请填写> 需补全。")
    if "Word" in kinds:
        print("Word 打开时请按提示更新域（或按 F9 更新目录）。")


def main():
    ap = argparse.ArgumentParser(description="论文Word草案 + 答辩PPT草案 生成器")
    ap.add_argument("--input", nargs="+", default=[DEFAULT_INPUT],
                    help="输入目录或文件（可多个）")
    ap.add_argument("--output", default=DEFAULT_OUTPUT, help="输出目录")
    ap.add_argument("--only", choices=["word", "ppt"],
                    help="draft 模式只生成其中一种")
    ap.add_argument("--mode", choices=["auto", "refs", "draft"],
                    default="auto",
                    help="auto: input/ 有题目文件(topic.md等)时走参考资料模式；"
                         "refs/draft 强制指定")
    ap.add_argument("--llm", action="store_true",
                    help="用 LLM 增强草案质量（需设置 LLM_API_KEY，"
                         "可选 LLM_BASE_URL / LLM_MODEL，OpenAI 兼容接口）")
    ap.add_argument("--refresh-fields", action="store_true",
                    help="生成后用本机 Word 静默刷新目录/页码域（需已装 Word 和 pywin32）")
    ap.add_argument("--pdf", action="store_true",
                    help="生成 Word 时刷新域并导出 PDF（隐含 --refresh-fields）")
    ap.add_argument("--format-template", metavar="FILE",
                    help="从 YAML 文件加载 Word/PPT 格式覆盖")
    ap.add_argument("--polish", choices=["conservative", "standard", "strong"],
                    metavar="LEVEL", help="draft 模式对正文段落多轮润色"
                    "（需 LLM_API_KEY；conservative/standard/strong 三档；"
                    "输出带 <AI润色，请核对> 标记，失败回退原文）")
    ap.add_argument("--dry-run", action="store_true",
                    help="仅检查输入可读性、模式和 LLM 外发清单，不调用外部服务或生成产物")
    ap.add_argument("--yes", action="store_true",
                    help="非交互环境确认将资料发送至 LLM 端点")
    ap.add_argument("--lookup-metadata", action="store_true",
                    help="允许参考资料模式查询 Crossref（默认离线）")
    ap.add_argument("--report", metavar="FILE",
                    help="运行报告 JSON 输出路径")
    args = ap.parse_args()

    if args.format_template:
        from config.template import apply_template
        apply_template(args.format_template)
        from src import organizer as organizer_mod
        organizer_mod.reload_spec()
        pptx_builder.reload_spec()

    os.makedirs(args.output, exist_ok=True)
    from src.logging_setup import configure_logging
    configure_logging(os.path.join(args.output, "运行日志.log"))

    print("=" * 56)
    print("① 读取源文件")
    docs, errors = gather_docs(args.input)
    if not docs:
        print("\n⚠ 未读取到任何文件。请把 Word/PDF/TXT/Markdown/JSON/Excel/图片"
              "放进 input/ 后重试。")
        print("  （可先用示例：--input sample_input）")
        return 1
    print(f"  共读取 {len(docs)} 个文件，"
          f"{sum(len(d['blocks']) for d in docs)} 个内容块。")

    topic_doc, ref_docs = _split_topic(docs)
    mode = args.mode
    if mode == "auto":
        mode = "refs" if topic_doc is not None else "draft"

    if args.dry_run:
        if args.llm or args.polish or mode == "refs":
            _print_llm_transfer_summary(docs)
        from src.runtime_report import write_report
        write_report(args.report or os.path.join(args.output, "运行报告.json"),
                     {"status": "dry_run", "files": [d["source"] for d in docs],
                      "mode": mode,
                      "llm_requested": bool(args.llm or args.polish or mode == "refs")})
        print("  [dry-run] 已完成输入检查，未生成产物。")
        return 0

    if not _confirm_llm_transfer(args, docs, required=(mode == "refs")):
        return 2
    if mode == "refs":
        return _run_refs_mode(args, topic_doc, ref_docs)
    # draft 模式沿用原流程（docs 不剔除题目文件）

    print("② 整理内容结构")
    thesis, deck = organize(docs)
    if args.llm:
        print("②+ LLM 增强")
        from src import llm_enhancer
        thesis, deck = llm_enhancer.enhance(thesis, deck, docs)
    if args.polish:
        print("②+ 正文润色")
        from src import llm_enhancer
        if not llm_enhancer.is_available():
            print("  [LLM] 未设置 LLM_API_KEY，跳过润色。")
        else:
            n = llm_enhancer.polish_paragraphs(thesis, args.polish)
            print(f"  [LLM] 润色完成，{n} 个段落已标记 <AI润色，请核对>。")
    print(f"  论文：{len(thesis['chapters'])} 章；PPT：{len(deck['slides'])} 页。")

    print("③ 生成草案")
    ok = True
    outputs = []
    ppt_warnings: list = []
    if args.only != "ppt":
        wp = os.path.join(args.output, "论文草案.docx")
        wp = _build_with_retry(docx_builder.build, thesis, wp)
        if wp:
            outputs.append(wp)
            print(f"  ✔ Word: {wp}")
            if args.refresh_fields or args.pdf:
                from src import postprocess
                postprocess.refresh_word_fields(wp, export_pdf=args.pdf)
        else:
            ok = False
    if args.only != "word":
        pp = os.path.join(args.output, "答辩PPT草案.pptx")
        result = _build_with_retry(pptx_builder.build, deck, pp)
        if result:
            pp = result
            outputs.append(pp)
            print(f"  ✔ PPT : {pp}")
            # BuildResult 携带告警列表（T0-4）
            ppt_warnings = getattr(result, "warnings", [])
        else:
            ok = False

    print("=" * 56)
    _print_draft_completion(outputs)
    from src.runtime_report import write_report
    write_report(args.report or os.path.join(args.output, "运行报告.json"),
                 {"status": "ok" if ok else "error", "mode": "draft",
                  "files": len(docs), "read_errors": errors,
                  "outputs": outputs, "llm": bool(args.llm),
                  "warnings": ppt_warnings})
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
