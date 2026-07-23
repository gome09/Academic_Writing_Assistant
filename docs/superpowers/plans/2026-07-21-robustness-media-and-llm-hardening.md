# 健壮性修复 + 图表保真 + LLM 加固 实施计划

> **实施状态更新（2026-07-23）**：本计划的稳定性、媒体、LLM、域更新和 PPT
> 结构校验任务已完成。后续增量实现又加入了章节有序 `blocks`、附录保留、PPT
> 媒体自动版式、YAML 格式模板、LLM 外发确认、确定性参考文献和运行报告。
> 本文中的任务步骤与旧行号是历史实施记录，当前入口和用法以
> [thesis_project/README.md](../../../thesis_project/README.md) 为准。

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 修复个人日常使用中最高频的翻车点（文件占用、扫描 PDF、手工排版标题、摘要误判），让源文件中的表格/图片真实进入 Word 草案，加固 LLM 层（超时/JSON mode/批量化/键归一化），并补上演讲备注、域自动更新与 PPT 结构校验。

**Architecture:** 沿现有四层管道（readers → organizer → builders，llm_enhancer 旁路增强）
就地改进。当前章节以有序 `blocks` 为规范模型，`tables`/`images`/`paras` 保留为兼容视图；
Word 域刷新用可选的 win32com 后处理模块，缺 pywin32/Word 时优雅降级。

**Tech Stack:** Python 3.13、python-docx、python-pptx、pdfplumber、openai>=1.0、pytest；可选 pywin32（仅 Windows 域刷新/导 PDF）。

**执行约定：**
- 所有命令在 `thesis_project/` 目录下执行；测试命令为 `python -m pytest tests/<文件> -v`。
- 每个任务完成后跑一次全量 `python -m pytest tests/ -q` 再提交，防止跨模块回归。
- 提交信息沿用仓库现有风格（`feat(readers): ...` / `fix(organizer): ...`），并以下行结尾：
  `Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`
- 测试文件头部加 `# -*- coding: utf-8 -*-`，import 风格与现有测试一致（`from src import xxx`，conftest.py 已处理 sys.path）。

**关键现状事实（写代码前请核对，行号为 2026-07-21 版本）：**
- `main.py:96-103` 两处 `build()` 调用无异常保护；`docx_builder.py:311`、`pptx_builder.py:211` 的 `save()` 被占用时 PermissionError 直接冒泡。
- `readers.py:187-229` `read_docx` 只认 `heading\s*(\d)` 样式名（208 行）；`readers.py:253` `_PDF_HEADING` 会把 "2023 年…" 识别为标题；`read_pdf` 对 0 文本块不报警。
- `organizer.py:144-149` 摘要正则 `^(摘[，\s]*要|abstract)` 无词边界；147-148 取下一块不跳过关键词行/标题。
- `organizer.py:102-103` `_DROP_TITLE` 命中的章（含附录）被静默 `continue` 丢弃。
- `organizer.py:207-213` 表格经 `_table_to_text` 拍平进 `paras`；`docx_builder.py` 无 `add_table`/`add_picture` 逻辑；`format_spec.py` 的 `figure`/`table` 块整块未被读取。
- `llm_enhancer.py:33-36` `_client()` 无 timeout/max_retries；`39-47` `_chat` 无 `response_format`；`230` `data.get(name, [])` 未归一化键；`rebuild_deck:195-197` 经 `bullets_fn=_safe_bullets` 产生每章一次的 N+1 调用；无演讲备注功能。
- `organizer.py:315` `_build_deck` 调 `to_bullets(paras)` 不传标题；slide dict 无 `bucket` 字段，`pptx_builder._check_structure:215-238` 因此只校验 cover/outline/thanks。
- 章节节点 dict 在 `organizer.py` 有 9 处字面量创建（179、186、188、196、199、204、216-218、224、227），`rechapter`（`llm_enhancer.py:240、244`）另有 2 处。

---

## Phase 1 稳定性

### Task 1: 输出文件被 Word/WPS 占用时自动改名重试

**Files:**
- Modify: `thesis_project/src/main.py`（新增 `_build_with_retry`，改写 95-103 行的生成段）
- Test: `thesis_project/tests/test_main_build_retry.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
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
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_main_build_retry.py -v`
Expected: FAIL，`AttributeError: ... has no attribute '_build_with_retry'`

- [ ] **Step 3: 实现**

在 `main.py` 的 `gather_docs` 之后新增：

```python
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
```

改写 `main()` 的生成段（原 95-103 行）：

```python
    print("③ 生成草案")
    ok = True
    if args.only != "ppt":
        wp = os.path.join(args.output, "论文草案.docx")
        wp = _build_with_retry(docx_builder.build, thesis, wp)
        if wp:
            print(f"  ✔ Word: {wp}")
        else:
            ok = False
    if args.only != "word":
        pp = os.path.join(args.output, "答辩PPT草案.pptx")
        pp = _build_with_retry(pptx_builder.build, deck, pp)
        if pp:
            print(f"  ✔ PPT : {pp}")
        else:
            ok = False
```

并把末尾 `return 0` 改为 `return 0 if ok else 1`。

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_main_build_retry.py tests/test_main_gather.py -v`
Expected: 全部 PASS

- [ ] **Step 5: 全量测试 + 提交**

```bash
python -m pytest tests/ -q
git add src/main.py tests/test_main_build_retry.py
git commit -m "fix(main): retry with numbered filename when output locked by Word/WPS

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 2: 扫描版 PDF（无文本层）明确报错

**Files:**
- Modify: `thesis_project/src/readers.py`（新增 `_ensure_has_text`，`read_pdf` 末尾调用）
- Test: `thesis_project/tests/test_readers_pdf_scanned.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import pytest
from src import readers


def test_ensure_has_text_raises_for_scanned_pdf():
    blocks = [readers._block("paragraph", "")]
    with pytest.raises(RuntimeError, match="扫描件"):
        readers._ensure_has_text(blocks, "a.pdf")


def test_ensure_has_text_raises_for_empty_blocks():
    with pytest.raises(RuntimeError):
        readers._ensure_has_text([], "a.pdf")


def test_ensure_has_text_passes_with_text():
    readers._ensure_has_text([readers._block("paragraph", "有内容")], "a.pdf")


def test_ensure_has_text_passes_with_table_only():
    readers._ensure_has_text(
        [readers._block("table", "", rows=[["a", "b"]])], "a.pdf")
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_readers_pdf_scanned.py -v`
Expected: FAIL，`AttributeError: ... no attribute '_ensure_has_text'`

- [ ] **Step 3: 实现**

在 `readers.py` 的 PDF 区块（`_pdf_lines_to_blocks` 附近）新增：

```python
def _ensure_has_text(blocks: list, path: str) -> None:
    """扫描件（图片型 PDF）提取不到文字时报错，避免静默生成全占位符骨架。"""
    if any(b.get("text") for b in blocks):
        return
    if any(b.get("kind") == "table" for b in blocks):
        return
    raise RuntimeError(
        "未提取到任何文字，可能是扫描件（图片型 PDF）。"
        "请先用 OCR 工具（如 WPS/Acrobat/umi-ocr）转成可复制文本的 PDF 再试")
```

在 `read_pdf` 的 `return {...}`（约 330 行）之前加一行 `_ensure_has_text(blocks, path)`。
注意 pdfplumber 分支与 pypdf 回退分支共用同一个 return，只需加这一处。

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_readers_pdf_scanned.py tests/test_readers_pdf.py -v`
Expected: 全部 PASS

- [ ] **Step 5: 全量测试 + 提交**

```bash
python -m pytest tests/ -q
git add src/readers.py tests/test_readers_pdf_scanned.py
git commit -m "feat(readers): raise clear error for scanned PDFs with no text layer

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 3: docx 手工排版标题的启发式识别 + 年份行误判修复

**Files:**
- Modify: `thesis_project/src/readers.py`（`_PDF_HEADING` 253 行、`read_docx` 208-214 行、新增 `_looks_like_manual_heading`）
- Test: `thesis_project/tests/test_readers_docx_manual_heading.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import docx
from src import readers


def _save(tmp_path, paragraphs, name="t.docx"):
    d = docx.Document()
    for p in paragraphs:
        d.add_paragraph(p)   # Normal 样式，模拟"手工排版"文档
    path = str(tmp_path / name)
    d.save(path)
    return path


def test_manual_numbered_headings_promoted(tmp_path):
    path = _save(tmp_path, ["1 绪论",
                            "这是一段足够长的正文内容，描述研究背景与意义。",
                            "1.1 研究背景",
                            "另一段正文。"])
    doc = readers.read_docx(path)
    kinds = [(b["kind"], b["level"]) for b in doc["blocks"]]
    assert ("heading", 1) in kinds
    assert ("heading", 2) in kinds


def test_chapter_word_heading_promoted(tmp_path):
    path = _save(tmp_path, ["第一章 绪论", "正文。"])
    doc = readers.read_docx(path)
    assert doc["blocks"][0]["kind"] == "heading"
    assert doc["blocks"][0]["level"] == 1


def test_year_line_stays_paragraph(tmp_path):
    path = _save(tmp_path, ["2023 年国内研究综述指出该方向发展迅速。"])
    doc = readers.read_docx(path)
    assert doc["blocks"][0]["kind"] == "paragraph"


def test_long_numbered_sentence_stays_paragraph(tmp_path):
    path = _save(tmp_path,
                 ["1. 本文首先分析了现有方法的不足，然后提出了改进方案，最后进行了实验验证。"])
    doc = readers.read_docx(path)
    assert doc["blocks"][0]["kind"] == "paragraph"


def test_pdf_heading_regex_rejects_year():
    assert readers._PDF_HEADING.match("2023 年国内研究综述") is None
    assert readers._PDF_HEADING.match("2.3 实验设计") is not None
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_readers_docx_manual_heading.py -v`
Expected: FAIL（手工标题被读成 paragraph；年份正则断言失败）

- [ ] **Step 3: 实现**

253 行 `_PDF_HEADING` 改为（负向前瞻排除 4 位年份，注意用非捕获组保持 group(1) 语义不变）：

```python
_PDF_HEADING = re.compile(
    r"^(第\s*[一二三四五六七八九十百\d]+\s*章"
    r"|(?!\d{4}(?:[\s年]|$))\d+(\.\d+)*[\s、.．])")
```

在 `_PDF_HEADING` 定义之后新增：

```python
def _looks_like_manual_heading(text: str) -> bool:
    """无标题样式但形如标题：编号开头 + 短（≤25字）+ 不含句中/句末标点。"""
    if len(text) > 25 or re.search(r"[。！？；，,;]", text):
        return False
    return bool(_PDF_HEADING.match(text))
```

`read_docx` 中 208-214 行的分支改为（list 判断保持在启发式之前）：

```python
            m = re.search(r"heading\s*(\d)", style)
            if m:
                blocks.append(_block("heading", text, level=int(m.group(1))))
            elif style.startswith("list") or style.startswith("bullet"):
                blocks.append(_block("list_item", text))
            elif _looks_like_manual_heading(text):
                hm = _PDF_HEADING.match(text)
                blocks.append(_block("heading", text,
                                     level=_pdf_heading_level(hm.group(1))))
            else:
                blocks.append(_block("paragraph", text))
```

（`_looks_like_manual_heading`/`_pdf_heading_level` 定义在 read_docx 之后不影响运行时解析。）

- [ ] **Step 4: 运行确认通过（含既有 docx/pdf 测试防回归）**

Run: `python -m pytest tests/test_readers_docx_manual_heading.py tests/test_readers_docx_heading_chain.py tests/test_readers_pdf.py -v`
Expected: 全部 PASS。若 `test_readers_pdf.py` 有对旧正则的断言失败，先确认新行为正确再更新该断言。

- [ ] **Step 5: 全量测试 + 提交**

```bash
python -m pytest tests/ -q
git add src/readers.py tests/test_readers_docx_manual_heading.py
git commit -m "feat(readers): promote manually numbered docx paragraphs to headings; exclude year-prefixed lines

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 4: 摘要识别正则加词边界，取"下一段"时跳过关键词行与标题

**Files:**
- Modify: `thesis_project/src/organizer.py`（模块级新增 `_ABSTRACT_LABEL`，改写 `_extract_meta` 144-149 行）
- Test: `thesis_project/tests/test_organizer_abstract.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import organizer
from src.readers import _block


def _doc(blocks):
    return {"source": "t.txt", "type": "txt", "blocks": blocks, "meta": {}}


def test_abstraction_paragraph_not_treated_as_abstract():
    d = _doc([_block("paragraph",
                     "Abstraction is a key concept in computer science.")])
    meta = organizer._extract_meta([d])
    assert meta["abstract"] == organizer.PLACEHOLDER


def test_abstract_label_with_body_same_block():
    d = _doc([_block("paragraph", "摘要：本文研究了基于深度学习的图像识别方法。")])
    meta = organizer._extract_meta([d])
    assert meta["abstract"].startswith("本文研究了")


def test_abstract_next_block_skips_keyword_line():
    d = _doc([
        _block("paragraph", "摘要"),
        _block("paragraph", "关键词：深度学习 图像识别 卷积网络"),
        _block("paragraph", "本文研究了基于深度学习的图像识别方法，并完成了系统实现。"),
    ])
    meta = organizer._extract_meta([d])
    assert meta["abstract"].startswith("本文研究了")
    assert "关键词" not in meta["abstract"]


def test_abstract_next_block_skips_heading():
    d = _doc([
        _block("paragraph", "摘要"),
        _block("heading", "第一章 绪论", level=1),
        _block("paragraph", "正文段落。"),
    ])
    meta = organizer._extract_meta([d])
    # 下一块是标题：不应把标题文本当摘要
    assert meta["abstract"] != "第一章 绪论"
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_organizer_abstract.py -v`
Expected: 至少 `test_abstraction_...` 与 `skips_keyword_line` FAIL

- [ ] **Step 3: 实现**

在 `organizer.py` 模块级（`_DROP_TITLE` 附近）新增：

```python
_ABSTRACT_LABEL = re.compile(r"^(摘[，\s]*要|abstract\b)[：:\s]*", re.I)
```

`_extract_meta` 中 144-149 行改为：

```python
        # 摘要
        if abstract is None and _ABSTRACT_LABEL.match(t):
            # 摘要正文：同块去掉标签；太短则向后找第一个普通段落
            body = _ABSTRACT_LABEL.sub("", t).strip()
            if len(body) < 10:
                for nb in all_blocks[i + 1:i + 4]:
                    if nb["kind"] != "paragraph":
                        continue
                    if re.match(r"^(关键词|关键字|key\s*words)", nb["text"], re.I):
                        continue
                    body = nb["text"]
                    break
            abstract = body or None
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_organizer_abstract.py tests/test_organizer_meta.py tests/test_organizer_meta_yaml.py -v`
Expected: 全部 PASS

- [ ] **Step 5: 全量测试 + 提交**

```bash
python -m pytest tests/ -q
git add src/organizer.py tests/test_organizer_abstract.py
git commit -m "fix(organizer): word-boundary abstract label; skip keywords/headings when taking next block

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 5: 附录/目录/致谢剔除时打印提示

**Files:**
- Modify: `thesis_project/src/organizer.py`（`_split_special_chapters` 102-103 行）
- Test: `thesis_project/tests/test_organizer_special.py`（追加用例）

- [ ] **Step 1: 写失败测试（追加到 test_organizer_special.py 末尾）**

```python
def test_dropped_chapter_with_content_prints_notice(capsys):
    chapters = [
        {"title": "绪论", "level": 1, "paras": ["正文。"], "subs": []},
        {"title": "附录", "level": 1, "paras": ["问卷原文。", "代码清单。"], "subs": []},
    ]
    kept, refs = organizer._split_special_chapters(chapters)
    out = capsys.readouterr().out
    assert "附录" in out and "2 段" in out
    assert all(c["title"] != "附录" for c in kept)


def test_dropped_empty_chapter_silent(capsys):
    chapters = [{"title": "目录", "level": 1, "paras": [], "subs": []}]
    organizer._split_special_chapters(chapters)
    assert capsys.readouterr().out == ""
```

（若该文件没有 `import organizer`，按其现有 import 风格补 `from src import organizer`。）

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_organizer_special.py -v`
Expected: 新增两条 FAIL（无输出）

- [ ] **Step 3: 实现**

`_split_special_chapters` 中 `_DROP_TITLE` 命中分支（原 102-103 行）改为：

```python
        if _DROP_TITLE.match(title):
            n = len(ch.get("paras", [])) + sum(
                len(s.get("paras", [])) for s in ch.get("subs", []))
            if n:
                print(f"  [提示] 已从正文剔除章节《{ch['title']}》（{n} 段）"
                      "——摘要/目录/致谢/附录不进入正文")
            continue
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_organizer_special.py -v`
Expected: 全部 PASS

- [ ] **Step 5: 全量测试 + 提交**

```bash
python -m pytest tests/ -q
git add src/organizer.py tests/test_organizer_special.py
git commit -m "feat(organizer): notice when dropping non-empty special chapters

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Phase 2 内容保真（表格/图片进 Word）

**设计约定（三个任务共享）：**
- Block 新增 kind：`image`，携带 `data`（bytes）与 `ext`（如 `.png`）两个额外键。
- 章节节点（章/节/条）新增 `tables: [rows...]` 与 `images: [{"data","ext"}...]` 两个键；`paras` 保持纯字符串列表。
- 表格**不再**经 `_table_to_text` 拍平进 `paras`（PPT 要点因此不再混入 `a | b` 行——这是有意的行为变化）；`_table_to_text` 函数保留（既有单测仍然有效，未来可复用）。
- docx 渲染：媒体挂在所属节点末尾输出（不做段落级插入点还原），题注按章编号 `表{章}-{序}` / `图{章}-{序}`，激活 `WORD_SPEC["figure"]/["table"]` 这两块死配置。

### Task 6: readers 提取 docx 内嵌图片；PDF 仅计数提示

**Files:**
- Modify: `thesis_project/src/readers.py`（`read_docx` 201-221 行、`read_pdf` pdfplumber 分支）
- Test: `thesis_project/tests/test_readers_docx_images.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import base64
import io

import docx
from src import readers

# 1x1 像素 PNG
_PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQ"
    "DwAEhQGAhKmMIQAAAABJRU5ErkJggg==")


def test_read_docx_extracts_inline_image(tmp_path):
    d = docx.Document()
    d.add_paragraph("前文段落。")
    d.add_picture(io.BytesIO(_PNG_1PX))
    d.add_paragraph("后文段落。")
    path = str(tmp_path / "img.docx")
    d.save(path)

    doc = readers.read_docx(path)
    imgs = [b for b in doc["blocks"] if b["kind"] == "image"]
    assert len(imgs) == 1
    assert imgs[0]["data"][:4] == b"\x89PNG"
    assert imgs[0]["ext"] == ".png"
    # 图片出现在两段文字之间
    kinds = [b["kind"] for b in doc["blocks"]]
    assert kinds == ["paragraph", "image", "paragraph"]


def test_read_docx_without_images_unchanged(tmp_path):
    d = docx.Document()
    d.add_paragraph("只有文字。")
    path = str(tmp_path / "plain.docx")
    d.save(path)
    doc = readers.read_docx(path)
    assert all(b["kind"] != "image" for b in doc["blocks"])
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_readers_docx_images.py -v`
Expected: FAIL（无 image 块；`add_picture` 产生的空文本段被 `continue` 跳过）

- [ ] **Step 3: 实现**

`read_docx` 的 `w:p` 分支（原 202-214 行）：**先提图片，再判空文本**（图片段落通常无文字，旧代码会在 `if not text: continue` 处把它跳过）：

```python
        if el.tag == qn("w:p"):
            p = Paragraph(el, doc)
            # 段内图片：a:blip 的 r:embed 指向图片关系部件
            for blip in el.findall(".//" + qn("a:blip")):
                rid = blip.get(qn("r:embed"))
                part = doc.part.related_parts.get(rid) if rid else None
                if part is None:
                    continue
                b = _block("image")
                b["data"] = part.blob
                b["ext"] = os.path.splitext(str(part.partname))[1] or ".png"
                blocks.append(b)
            text = _clean(p.text)
            if not text:
                continue
            ...  # 以下标题/列表/段落分支保持 Task 3 之后的样子，不动
```

（`readers.py` 顶部已 `import os`、`import re`；`qn` 已在函数内导入，`a:`/`r:` 命名空间均在 python-docx 的 nsmap 中。）

`read_pdf` 的 pdfplumber 分支：页循环内累计 `img_count += len(page.images)`（在 `for page in pdf.pages:` 内任意位置加一行，循环前初始化 `img_count = 0`），循环结束后：

```python
            if img_count:
                print(f"  [提示] {os.path.basename(path)}：检测到 {img_count} 张图片，"
                      "PDF 图片暂不导入，请在草案中手工补图")
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_readers_docx_images.py tests/test_readers_docx_table_order.py -v`
Expected: 全部 PASS

- [ ] **Step 5: 全量测试 + 提交**

```bash
python -m pytest tests/ -q
git add src/readers.py tests/test_readers_docx_images.py
git commit -m "feat(readers): extract inline docx images as image blocks; count-only notice for PDF images

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 7: organizer 把表格/图片挂到章节树（tables/images 键）

**Files:**
- Modify: `thesis_project/src/organizer.py`（新增 `_node`，改写 `_build_chapters` 167-230 行）
- Modify: `thesis_project/src/llm_enhancer.py`（`rechapter` 媒体保留，237-246 行一带）
- Modify: `thesis_project/tests/test_organizer_table.py`（更新"拍平进 paras"的旧断言）
- Test: `thesis_project/tests/test_organizer_media.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import organizer, llm_enhancer
from src.readers import _block


def _doc(blocks):
    return {"source": "t.txt", "type": "txt", "blocks": blocks, "meta": {}}


def _img_block():
    b = _block("image")
    b["data"] = b"\x89PNGfake"
    b["ext"] = ".png"
    return b


def test_table_attached_to_chapter_not_paras():
    blocks = [_block("heading", "第一章 绪论", level=1),
              _block("paragraph", "正文段落。"),
              _block("table", "", rows=[["指标", "数值"], ["准确率", "94%"]])]
    chapters, _ = organizer._build_chapters([_doc(blocks)])
    ch = chapters[0]
    assert ch["tables"] == [[["指标", "数值"], ["准确率", "94%"]]]
    assert ch["paras"] == ["正文段落。"]          # 表格文本不再混入段落


def test_image_attached_to_sub():
    blocks = [_block("heading", "第一章", level=1),
              _block("heading", "1.1 背景", level=2),
              _img_block()]
    chapters, _ = organizer._build_chapters([_doc(blocks)])
    sub = chapters[0]["subs"][0]
    assert len(sub["images"]) == 1
    assert sub["images"][0]["ext"] == ".png"


def test_media_before_any_heading_goes_to_preface():
    chapters, _ = organizer._build_chapters([_doc([_img_block()])])
    assert chapters[0]["title"] == "前言"
    assert len(chapters[0]["images"]) == 1


def test_skeleton_branch_collects_media():
    blocks = [_block("paragraph", "无标题文档的正文。"), _img_block(),
              _block("table", "", rows=[["a", "b"]])]
    chapters, auto = organizer._build_chapters([_doc(blocks)])
    assert auto is True
    body = next(c for c in chapters if c["title"] == "研究内容")
    assert len(body["images"]) == 1 and len(body["tables"]) == 1


def test_rechapter_carries_media(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        lambda s, u: {"绪论": [0]})
    src_ch = {"title": "研究内容", "level": 1,
              "paras": ["第一段。", "第二段。"], "subs": [],
              "tables": [[["a", "b"]]], "images": [{"data": b"x", "ext": ".png"}]}
    thesis = {"auto_skeleton": True, "chapters": [src_ch]}
    llm_enhancer.rechapter(thesis)
    all_tables = [t for c in thesis["chapters"] for t in c.get("tables", [])]
    all_images = [i for c in thesis["chapters"] for i in c.get("images", [])]
    assert len(all_tables) == 1 and len(all_images) == 1
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_organizer_media.py -v`
Expected: FAIL（`KeyError: 'tables'` 等）

- [ ] **Step 3: 实现 organizer**

在 `_build_chapters` 之前新增节点工厂（消除 9 处字面量重复）：

```python
def _node(title, level):
    """章节树节点：章/节/条统一结构。"""
    return {"title": title, "level": level, "paras": [], "subs": [],
            "tables": [], "images": []}
```

`_build_chapters` 改写要点：
1. 9 处 `{"title": ..., "level": ..., "paras": [...], "subs": []}` 字面量全部换成 `_node(...)`（骨架分支先 `ch = _node(name, 1)` 再 `ch["paras"].append(PLACEHOLDER)`；sub3 也用 `_node(title, 3)`，多出的空 `subs` 键无害）。
2. 在 `for b in blocks:` 循环之前定义目标节点选择器：

```python
        def target_node():
            t = current_sub3 or current_sub or current_ch
            if t is not None:
                return t
            if not chapters or chapters[0]["title"] != "前言":
                chapters.insert(0, _node("前言", 1))
            return chapters[0]
```

3. 原 207-219 行的 else 分支拆为三支：

```python
            elif b["kind"] == "table":
                if b.get("rows"):
                    target_node()["tables"].append(b["rows"])
            elif b["kind"] == "image":
                target_node()["images"].append(
                    {"data": b.get("data"), "ext": b.get("ext", ".png")})
            else:  # 正文/列表/代码
                text = b["text"]
                if not text:
                    continue
                target_node()["paras"].append(text)
```

4. 无标题骨架分支（原 220-228 行）改为：

```python
        for name in DEFAULT_CHAPTERS:
            ch = _node(name, 1)
            ch["paras"].append(PLACEHOLDER)
            chapters.append(ch)
        body = _node("研究内容", 1)
        for b in blocks:
            if b["kind"] == "table" and b.get("rows"):
                body["tables"].append(b["rows"])
            elif b["kind"] == "image":
                body["images"].append(
                    {"data": b.get("data"), "ext": b.get("ext", ".png")})
            elif b["text"]:
                body["paras"].append(b["text"])
        chapters.insert(3, body)
```

- [ ] **Step 4: 实现 rechapter 媒体保留（llm_enhancer.py）**

`rechapter` 中构建 `new_chapters` 后、`thesis["chapters"] = new_chapters` 之前插入：

```python
    # 媒体不参与语义分配：统一挂到"研究内容"（无则最后一章）
    tables = [t for c in thesis["chapters"] for t in c.get("tables", [])]
    images = [im for c in thesis["chapters"] for im in c.get("images", [])]
    for c in new_chapters:
        c.setdefault("tables", [])
        c.setdefault("images", [])
    if tables or images:
        carrier = next((c for c in new_chapters if c["title"] == "研究内容"),
                       new_chapters[-1])
        carrier["tables"].extend(tables)
        carrier["images"].extend(images)
```

- [ ] **Step 5: 更新旧断言**

读 `tests/test_organizer_table.py`：凡断言"表格文本出现在 `paras`/deck bullets 中"的用例，改为断言 `ch["tables"]` 挂载了 rows；`_table_to_text` 的直接单测保留不动。

- [ ] **Step 6: 运行确认通过**

Run: `python -m pytest tests/test_organizer_media.py tests/test_organizer_table.py tests/test_organizer_skeleton.py tests/test_llm_rechapter.py -v`
Expected: 全部 PASS

- [ ] **Step 7: 全量测试 + 提交**

```bash
python -m pytest tests/ -q
git add src/organizer.py src/llm_enhancer.py tests/test_organizer_media.py tests/test_organizer_table.py
git commit -m "feat(organizer): attach tables/images to chapter tree instead of flattening

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 8: docx_builder 输出三线表与图片（含题注编号）

**Files:**
- Modify: `thesis_project/src/docx_builder.py`（新增 `_add_three_line_table`/`_add_image`/`_render_media`，正文循环 274-297 行插调用）
- Test: `thesis_project/tests/test_docx_builder_media.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import base64

import docx
from src import docx_builder

_PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQ"
    "DwAEhQGAhKmMIQAAAABJRU5ErkJggg==")


def _min_thesis():
    return {"title": "题目", "author": "作者",
            "abstract": "本文摘要。", "abstract_en": "EN abstract.",
            "keywords": ["a", "b", "c"], "keywords_en": ["a", "b", "c"],
            "chapters": [{"title": "绪论", "level": 1, "paras": ["正文。"],
                          "subs": [], "tables": [], "images": []}],
            "auto_skeleton": False, "references": ["某文献[J]. 2024."]}


def test_real_table_rendered_with_caption(tmp_path):
    thesis = _min_thesis()
    thesis["chapters"][0]["tables"] = [[["指标", "数值"], ["准确率", "94%"]]]
    out = str(tmp_path / "o.docx")
    docx_builder.build(thesis, out)

    d = docx.Document(out)
    assert len(d.tables) == 1
    assert d.tables[0].cell(0, 0).text == "指标"
    assert d.tables[0].cell(1, 1).text == "94%"
    texts = [p.text for p in d.paragraphs]
    assert any(t.startswith("表1-1") for t in texts)


def test_image_rendered_with_caption(tmp_path):
    thesis = _min_thesis()
    thesis["chapters"][0]["images"] = [{"data": _PNG_1PX, "ext": ".png"}]
    out = str(tmp_path / "o.docx")
    docx_builder.build(thesis, out)

    d = docx.Document(out)
    assert len(d.inline_shapes) == 1
    texts = [p.text for p in d.paragraphs]
    assert any(t.startswith("图1-1") for t in texts)


def test_broken_image_falls_back_to_placeholder(tmp_path):
    thesis = _min_thesis()
    thesis["chapters"][0]["images"] = [{"data": b"not-an-image", "ext": ".png"}]
    out = str(tmp_path / "o.docx")
    docx_builder.build(thesis, out)          # 不应抛异常
    d = docx.Document(out)
    assert any("图片插入失败" in p.text for p in d.paragraphs)


def test_chapters_without_media_keys_still_build(tmp_path):
    thesis = _min_thesis()
    del thesis["chapters"][0]["tables"], thesis["chapters"][0]["images"]
    docx_builder.build(thesis, str(tmp_path / "o.docx"))   # .get() 容错
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_docx_builder_media.py -v`
Expected: 前三条 FAIL（无表格/图片/题注）

- [ ] **Step 3: 实现**

顶部补充 import（与现有 import 合并，已有的不重复加）：

```python
import io
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
```

（`qn`、`Cm`、`WD_ALIGN_PARAGRAPH` 等按文件现状复用；`_ALIGN` 字典已提供对齐映射。）

在 `_cn_num` 之前新增三个函数：

```python
_CAPTION_SPEC = {"font_cn": "宋体", "font_en": "Times New Roman",
                 "size_pt": 10.5, "line_spacing_pt": 20}


def _set_three_line_borders(tbl):
    """三线表：顶线/底线 1.5 磅、表头下线 0.75 磅，其余无框线。"""
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for tag, sz in (("top", "12"), ("bottom", "12")):
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:color"), "000000")
        borders.append(el)
    for tag in ("left", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:val"), "none")
        borders.append(el)
    tblPr.append(borders)
    for tc in tbl.rows[0]._tr.findall(qn("w:tc")):
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is None:
            tcPr = OxmlElement("w:tcPr")
            tc.insert(0, tcPr)
        tcB = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:color"), "000000")
        tcB.append(bottom)
        tcPr.append(tcB)


def _add_three_line_table(doc, rows, ci, ti):
    spec = W.get("table", {})
    caption = f"{spec.get('prefix', '表')}{ci}-{ti}　<请填写表题>"
    if spec.get("caption_position", "above") == "above":
        _add_para(doc, caption, _CAPTION_SPEC, align="center")
    n_cols = max(len(r) for r in rows)
    tbl = doc.add_table(rows=len(rows), cols=n_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_i, row in enumerate(rows):
        for c_i in range(n_cols):
            cell = tbl.rows[r_i].cells[c_i]
            cell.text = row[c_i] if c_i < len(row) else ""
            for p in cell.paragraphs:
                for run in p.runs:
                    _set_run_font(run, "宋体", "Times New Roman", 10.5,
                                  bold=(r_i == 0))
    _set_three_line_borders(tbl)
    if spec.get("caption_position", "above") != "above":
        _add_para(doc, caption, _CAPTION_SPEC, align="center")


def _add_image(doc, img, ci, fi):
    spec = W.get("figure", {})
    p = doc.add_paragraph()
    p.alignment = _ALIGN["center"]
    try:
        p.add_run().add_picture(io.BytesIO(img["data"]), width=Cm(14))
    except Exception:  # noqa: BLE001 损坏/不支持的图片格式
        _add_para(doc, "<图片插入失败，请手工补图>", W["body"])
        return
    caption = f"{spec.get('prefix', '图')}{ci}-{fi}　<请填写图题>"
    _add_para(doc, caption, _CAPTION_SPEC, align="center")


def _render_media(doc, node, ci, counters):
    """渲染节点挂载的表格与图片；counters 按章累计编号。"""
    for rows in node.get("tables", []):
        counters["table"] += 1
        _add_three_line_table(doc, rows, ci, counters["table"])
    for img in node.get("images", []):
        counters["figure"] += 1
        _add_image(doc, img, ci, counters["figure"])
```

（`_ALIGN["center"]` 若键名不同，以文件里 `_ALIGN` 字典实际键为准；`Cm` 若未导入则从 `docx.shared` 补。）

正文循环（274-297 行）插入调用——每章开头重置计数器，各层段落渲染完后渲染媒体：

```python
    for ci, ch in enumerate(thesis["chapters"], 1):
        counters = {"table": 0, "figure": 0}
        ...                                   # 章标题 + 章 paras（原代码不动）
        _render_media(doc, ch, ci, counters)  # 加在章 paras 循环之后
        for si, sub in enumerate(ch.get("subs", []), 1):
            ...                               # 节标题 + 节 paras（原代码不动）
            _render_media(doc, sub, ci, counters)   # 加在节 paras 循环之后
            for ti, sub3 in enumerate(sub.get("subs", []), 1):
                ...                           # 条标题 + 条 paras（原代码不动）
                _render_media(doc, sub3, ci, counters)  # 加在条 paras 循环之后
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_docx_builder_media.py tests/test_e2e.py -v`
Expected: 全部 PASS

- [ ] **Step 5: 样例冒烟 + 全量测试 + 提交**

```bash
python src/main.py --input sample_input --output output_smoke
python -m pytest tests/ -q
git add src/docx_builder.py tests/test_docx_builder_media.py
git commit -m "feat(docx): render real three-line tables and inline images with chapter-numbered captions

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

（冒烟跑完删除 `output_smoke/`，不要提交。）

---

## Phase 3 LLM 层加固

### Task 9: OpenAI 客户端加超时与重试上限

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py`（`_client` 33-36 行）
- Test: `thesis_project/tests/test_llm_client.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import llm_enhancer


def test_client_default_timeout(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.delenv("LLM_TIMEOUT", raising=False)
    captured = {}

    class FakeOpenAI:
        def __init__(self, **kw):
            captured.update(kw)

    monkeypatch.setattr("openai.OpenAI", FakeOpenAI)
    llm_enhancer._client()
    assert captured["timeout"] == 60.0
    assert captured["max_retries"] == 1


def test_client_timeout_env_override(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("LLM_TIMEOUT", "30")
    captured = {}

    class FakeOpenAI:
        def __init__(self, **kw):
            captured.update(kw)

    monkeypatch.setattr("openai.OpenAI", FakeOpenAI)
    llm_enhancer._client()
    assert captured["timeout"] == 30.0
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_llm_client.py -v`
Expected: FAIL，`KeyError: 'timeout'`

- [ ] **Step 3: 实现**

先看 `_client`（33-36 行）现状：若 `from openai import OpenAI` 在模块顶部，测试里 `monkeypatch.setattr("openai.OpenAI", ...)` 不会生效——此时把导入移进函数内（下方写法）。改为：

```python
def _client():
    import openai
    timeout = float(os.environ.get("LLM_TIMEOUT", "60"))
    return openai.OpenAI(api_key=os.environ["LLM_API_KEY"],
                         base_url=os.environ.get("LLM_BASE_URL") or None,
                         timeout=timeout, max_retries=1)
```

（`api_key`/`base_url` 的取法以原实现为准，只新增 `timeout`/`max_retries` 两个参数与 `LLM_TIMEOUT` 读取。）

- [ ] **Step 4: 运行确认通过 + 提交**

```bash
python -m pytest tests/test_llm_client.py tests/test_llm_base.py -v
python -m pytest tests/ -q
git add src/llm_enhancer.py tests/test_llm_client.py
git commit -m "feat(llm): configurable timeout (LLM_TIMEOUT, default 60s) and max_retries=1

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

同时在 `README.md` 的 LLM 环境变量表格中加一行 `LLM_TIMEOUT`（否/单步超时秒数，默认 60），一并提交。

---

### Task 10: JSON mode（response_format），保留剥壳解析兜底

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py`（`_chat` 39-47、`_chat_json` 50-60）
- Modify: `thesis_project/tests/test_llm_base.py`（既有 `_chat` 打桩 lambda 加 `**kw`）
- Test: `thesis_project/tests/test_llm_json_mode.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import llm_enhancer


class _FakeResp:
    def __init__(self, content):
        msg = type("M", (), {"content": content})()
        self.choices = [type("C", (), {"message": msg})()]


def test_chat_passes_response_format(monkeypatch):
    calls = []

    class FakeCompletions:
        @staticmethod
        def create(**kw):
            calls.append(kw)
            return _FakeResp('{"a": 1}')

    fake_client = type("Cl", (), {})()
    fake_client.chat = type("Ch", (), {"completions": FakeCompletions})()
    monkeypatch.setattr(llm_enhancer, "_client", lambda: fake_client)

    llm_enhancer._chat("s", "u", json_mode=True)
    assert calls[0]["response_format"] == {"type": "json_object"}


def test_chat_plain_has_no_response_format(monkeypatch):
    calls = []

    class FakeCompletions:
        @staticmethod
        def create(**kw):
            calls.append(kw)
            return _FakeResp("ok")

    fake_client = type("Cl", (), {})()
    fake_client.chat = type("Ch", (), {"completions": FakeCompletions})()
    monkeypatch.setattr(llm_enhancer, "_client", lambda: fake_client)

    llm_enhancer._chat("s", "u")
    assert "response_format" not in calls[0]


def test_chat_falls_back_when_endpoint_rejects_json_mode(monkeypatch):
    calls = []

    class FakeCompletions:
        @staticmethod
        def create(**kw):
            calls.append(kw)
            if "response_format" in kw:
                raise TypeError("unexpected keyword")
            return _FakeResp('{"a": 1}')

    fake_client = type("Cl", (), {})()
    fake_client.chat = type("Ch", (), {"completions": FakeCompletions})()
    monkeypatch.setattr(llm_enhancer, "_client", lambda: fake_client)

    assert llm_enhancer._chat("s", "u", json_mode=True) == '{"a": 1}'
    assert len(calls) == 2 and "response_format" not in calls[1]


def test_chat_json_requests_json_mode(monkeypatch):
    seen = {}

    def fake_chat(system, user, json_mode=False):
        seen["json_mode"] = json_mode
        return '{"a": 1}'

    monkeypatch.setattr(llm_enhancer, "_chat", fake_chat)
    assert llm_enhancer._chat_json("s", "u") == {"a": 1}
    assert seen["json_mode"] is True
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_llm_json_mode.py -v`
Expected: FAIL（`_chat` 不接受 `json_mode`）

- [ ] **Step 3: 实现**

`_chat` 改为（消息构造、model、temperature 沿用原实现）：

```python
def _chat(system: str, user: str, json_mode: bool = False) -> str:
    """唯一网络出口。json_mode 尝试 response_format，端点不支持时自动降级。"""
    kwargs = {"model": os.environ.get("LLM_MODEL", "gpt-4o-mini"),
              "messages": [{"role": "system", "content": system},
                           {"role": "user", "content": user}],
              "temperature": 0.2}
    if json_mode:
        try:
            resp = _client().chat.completions.create(
                response_format={"type": "json_object"}, **kwargs)
            return resp.choices[0].message.content.strip()
        except Exception:  # noqa: BLE001 端点不支持 response_format 时降级
            pass
    resp = _client().chat.completions.create(**kwargs)
    return resp.choices[0].message.content.strip()
```

`_chat_json`（50-60 行）里对 `_chat` 的调用改为 `_chat(sys2, user, json_mode=True)`（`sys2` 指该函数内拼接了"只输出 JSON"的 system 变量名，以实际代码为准——JSON mode 要求消息中出现 "JSON" 字样，现有提示词已满足）。剥壳与 `raw_decode` 解析逻辑**保持不变**作为兜底。

- [ ] **Step 4: 更新既有打桩**

`grep -rn "monkeypatch.setattr(llm_enhancer, \"_chat\"" tests/`，把所有 `lambda s, u: ...` 改为 `lambda s, u, **kw: ...`（test_llm_base.py 有 5 处；其它 test_llm_* 若打的是 `_chat_json` 则无需改）。

- [ ] **Step 5: 运行确认通过 + 提交**

```bash
python -m pytest tests/test_llm_json_mode.py tests/test_llm_base.py -v
python -m pytest tests/ -q
git add src/llm_enhancer.py tests/test_llm_json_mode.py tests/test_llm_base.py
git commit -m "feat(llm): use JSON mode with graceful fallback for non-supporting endpoints

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 11: rechapter 对 LLM 返回的章节名键做归一化

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py`（`rechapter` 227-234 行）
- Test: `thesis_project/tests/test_llm_rechapter.py`（追加用例）

- [ ] **Step 1: 写失败测试（追加）**

```python
def test_rechapter_normalizes_returned_keys(monkeypatch):
    # LLM 返回的键带全角空格与编号前缀，仍应命中骨架章节
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        lambda s, u: {"1. 绪论　": [0], "总结与展望 ": [1]})
    thesis = {"auto_skeleton": True,
              "chapters": [{"title": "研究内容", "level": 1,
                            "paras": ["第一段。", "第二段。"], "subs": []}]}
    llm_enhancer.rechapter(thesis)
    by_title = {c["title"]: c["paras"] for c in thesis["chapters"]}
    assert by_title["绪论"] == ["第一段。"]
    assert by_title["总结与展望"] == ["第二段。"]
```

（import 与既有 test_llm_rechapter.py 风格一致。）

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_llm_rechapter.py -v`
Expected: 新用例 FAIL（键不匹配，段落全部留在"研究内容"）

- [ ] **Step 3: 实现**

`rechapter` 中 `data.get(name, [])` 的查表（227-234 行）改为先归一化：

```python
    from src.organizer import _strip_numbering
    norm_data = {_norm_title(_strip_numbering(str(k))): v
                 for k, v in data.items()}
    used = set()
    assign = {}
    for name in DEFAULT_CHAPTERS:
        idxs = sorted(i for i in norm_data.get(_norm_title(name), [])
                      if isinstance(i, int) and 0 <= i < len(paras)
                      and i not in used)
        used.update(idxs)
        assign[name] = idxs
```

（`from src.organizer import ...` 与函数内既有的 organizer 导入合并成一行。）

- [ ] **Step 4: 运行确认通过 + 提交**

```bash
python -m pytest tests/test_llm_rechapter.py -v
python -m pytest tests/ -q
git add src/llm_enhancer.py tests/test_llm_rechapter.py
git commit -m "fix(llm): normalize chapter-name keys from rechapter response before lookup

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 12: PPT 要点提炼批量化（N+1 次调用 → 1 次）

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py`（新增 `_BULLET_BATCH_SYS`/`_llm_bullets_batch`，重写 `rebuild_deck` 186-197 行）
- Modify: `thesis_project/src/organizer.py`（`_to_bullets` 加 `title=None` 形参；`_build_deck` 315 行传 title，带 TypeError 兼容回退）
- Test: `thesis_project/tests/test_llm_bullets_batch.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import llm_enhancer


def _thesis():
    def ch(title, para):
        return {"title": title, "level": 1, "paras": [para], "subs": []}
    return {"title": "题目", "author": "作者",
            "chapters": [ch("绪论", "研究背景是重要的现实问题。"),
                         ch("系统设计", "系统分为三个模块，各司其职。"),
                         ch("总结与展望", "本文完成了预期目标。")]}


def test_rebuild_deck_makes_exactly_two_llm_calls(monkeypatch):
    calls = []

    def fake_chat_json(system, user):
        calls.append(system)
        if len(calls) == 1:          # 第一次：章节分类
            return {}
        return {"绪论": ["要点A", "要点B"]}   # 第二次：批量要点

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    deck = llm_enhancer.rebuild_deck(_thesis())

    assert len(calls) == 2            # 3 章也只有 2 次调用（分类 + 批量要点）
    content = next(s for s in deck["slides"]
                   if s["type"] == "content" and s["title"] == "绪论")
    assert content["bullets"] == ["要点A", "要点B"]


def test_rebuild_deck_batch_failure_falls_back_to_rules(monkeypatch):
    def fake_chat_json(system, user):
        raise RuntimeError("接口超时")

    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    deck = llm_enhancer.rebuild_deck(_thesis())   # 不应抛异常
    contents = [s for s in deck["slides"] if s["type"] == "content"]
    assert contents and all(s["bullets"] for s in contents)


def test_llm_bullets_batch_normalizes_and_caps(monkeypatch):
    monkeypatch.setattr(
        llm_enhancer, "_chat_json",
        lambda s, u: {"绪 论": ["x" * 60] + [f"要点{i}" for i in range(8)]})
    out = llm_enhancer._llm_bullets_batch({"绪论": ["段落。"]})
    key = llm_enhancer._norm_title("绪论")
    assert key in out
    assert len(out[key]) <= 6                 # 上限 6 条
    assert all(len(b) <= 40 for b in out[key])  # 单条截 40 字
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_llm_bullets_batch.py -v`
Expected: FAIL（无 `_llm_bullets_batch`；调用次数为 4 而非 2）

- [ ] **Step 3: 实现 llm_enhancer**

在 `_llm_bullets` 附近新增：

```python
_BULLET_BATCH_SYS = ("你是答辩PPT助手。为每一章提炼要点：每章最多6条，"
                     "每条不超过40字，只做压缩改写，不得新增事实。"
                     '只输出 JSON：{"章节标题": ["要点", ...], ...}')


def _llm_bullets_batch(sections: dict) -> dict:
    """一次请求为全部章节提炼要点。sections: {标题: [段落...]}。

    返回 {归一化标题: bullets}；结果非 dict 时抛错交上层回退。
    """
    payload = {t: "\n".join(ps)[:2000] for t, ps in sections.items() if ps}
    data = _chat_json(_BULLET_BATCH_SYS,
                      json.dumps(payload, ensure_ascii=False))
    if not isinstance(data, dict):
        raise ValueError("批量要点结果不是 JSON 对象")
    out = {}
    for k, v in data.items():
        if isinstance(v, list):
            bullets = [str(x).strip()[:40] for x in v if str(x).strip()]
            if bullets:
                out[_norm_title(str(k))] = bullets[:6]
    return out
```

`rebuild_deck` 重写为：

```python
def rebuild_deck(thesis: dict) -> dict:
    """LLM 分类 + 批量要点重建 PPT 大纲；任一步失败回退规则实现。"""
    from src.organizer import _build_deck, _classify, _to_bullets
    try:
        mapping = classify_chapters([c["title"] for c in thesis["chapters"]])
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] 章节分类失败，改用关键词规则：{e}")
        mapping = {}

    sections = {}
    for ch in thesis["chapters"]:
        paras = list(ch["paras"])
        for sub in ch.get("subs", []):
            paras.extend(sub["paras"])
            for sub3 in sub.get("subs", []):
                paras.extend(sub3["paras"])
        sections[ch["title"]] = paras
    try:
        batch = _llm_bullets_batch(sections)
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] 批量要点失败，改用规则提炼：{e}")
        batch = {}

    def bullets_fn(paras, title=None):
        if title is not None:
            hit = batch.get(_norm_title(title))
            if hit:
                return hit
        return _to_bullets(paras)

    meta = {"title": thesis["title"], "author": thesis["author"]}
    return _build_deck(meta, thesis["chapters"],
                       classify_fn=lambda t: mapping.get(_norm_title(t)) or _classify(t),
                       bullets_fn=bullets_fn)
```

说明：批量失败时整体回退**规则**提炼（风格统一、零额外调用），不再逐章调 LLM；`_llm_bullets`/`_safe_bullets` 保留（既有测试仍覆盖），不再被 `rebuild_deck` 引用。

- [ ] **Step 4: 实现 organizer 侧 title 透传**

`_to_bullets` 签名改为 `def _to_bullets(paras, max_bullets=12, max_len=40, title=None):`（`title` 仅占位，函数体不用）。
`_build_deck` 315 行改为兼容注入方旧签名的调用：

```python
        try:
            bl = to_bullets(paras, title=ch["title"])
        except TypeError:      # 注入的 bullets_fn 不接受 title 时退回旧签名
            bl = to_bullets(paras)
        buckets[key].append({"title": ch["title"], "bullets": bl})
```

- [ ] **Step 5: 运行确认通过 + 提交**

```bash
python -m pytest tests/test_llm_bullets_batch.py tests/test_llm_bullets.py tests/test_safe_bullets_summary.py tests/test_deck_overflow.py -v
python -m pytest tests/ -q
git add src/llm_enhancer.py src/organizer.py tests/test_llm_bullets_batch.py
git commit -m "feat(llm): batch bullet extraction in one call instead of per-chapter N+1

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 13: 答辩演讲备注（speaker notes）

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py`（新增 `_NOTES_SYS`/`add_speaker_notes`，`enhance` 268-273 行后接入）
- Modify: `thesis_project/src/pptx_builder.py`（build 循环 195-202 行写入 notes_slide）
- Test: `thesis_project/tests/test_llm_notes.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from pptx import Presentation

from src import llm_enhancer, pptx_builder


def test_add_speaker_notes_attaches_marked_notes(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        lambda s, u: {"绪论": "各位老师好，本章介绍研究背景。"})
    deck = {"title": "t", "slides": [
        {"type": "cover", "title": "t", "subtitle": "s"},
        {"type": "content", "title": "绪论", "bullets": ["背景"]},
    ]}
    llm_enhancer.add_speaker_notes(deck)
    assert deck["slides"][1]["notes"].startswith(llm_enhancer.AI_MARK)
    assert "研究背景" in deck["slides"][1]["notes"]
    assert "notes" not in deck["slides"][0]          # 非 content 页不写


def test_add_speaker_notes_failure_is_silent(monkeypatch):
    def boom(s, u):
        raise RuntimeError("超时")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    deck = {"title": "t", "slides": [
        {"type": "content", "title": "绪论", "bullets": ["a"]}]}
    llm_enhancer.add_speaker_notes(deck)             # 不抛异常
    assert "notes" not in deck["slides"][0]


def test_pptx_builder_writes_notes(tmp_path):
    deck = {"title": "题目", "slides": [
        {"type": "cover", "title": "题目", "subtitle": "答辩人：张三"},
        {"type": "content", "title": "绪论", "bullets": ["背景"],
         "notes": "这里是口播备注。"},
    ]}
    out = str(tmp_path / "o.pptx")
    pptx_builder.build(deck, out)
    prs = Presentation(out)
    assert prs.slides[1].notes_slide.notes_text_frame.text == "这里是口播备注。"
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_llm_notes.py -v`
Expected: FAIL（无 `add_speaker_notes`；pptx 无备注）

- [ ] **Step 3: 实现 llm_enhancer**

```python
_NOTES_SYS = ("你是答辩教练。为每页幻灯片写80~120字的口语化演讲备注，"
              "只依据给定要点组织语言，不得编造数据或结论。"
              '只输出 JSON：{"页标题": "备注", ...}')


def add_speaker_notes(deck: dict) -> None:
    """为 content 页生成演讲备注写入 slide["notes"]；失败不影响主流程。"""
    contents = [s for s in deck.get("slides", []) if s.get("type") == "content"]
    if not contents:
        return
    payload = {s["title"]: s.get("bullets", []) for s in contents}
    try:
        data = _chat_json(_NOTES_SYS, json.dumps(payload, ensure_ascii=False))
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] 演讲备注生成失败：{e}")
        return
    if not isinstance(data, dict):
        return
    norm = {_norm_title(str(k)): v for k, v in data.items()}
    for s in contents:
        note = norm.get(_norm_title(s["title"]))
        if isinstance(note, str) and note.strip():
            s["notes"] = f"{AI_MARK} {note.strip()}"
```

`enhance` 中 `rebuild_deck` 的 try/except 之后追加：

```python
    try:
        add_speaker_notes(deck)
        print("  [LLM] 演讲备注生成完成")
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] 演讲备注失败：{e}")
```

- [ ] **Step 4: 实现 pptx_builder**

build 的 slide 循环（195-202 行）里，`_DISPATCH` 分发之后追加：

```python
        if s.get("notes"):
            prs.slides[-1].notes_slide.notes_text_frame.text = s["notes"]
```

- [ ] **Step 5: 运行确认通过 + 提交 + README**

```bash
python -m pytest tests/test_llm_notes.py tests/test_pptx_warnings.py -v
python -m pytest tests/ -q
git add src/llm_enhancer.py src/pptx_builder.py tests/test_llm_notes.py README.md
git commit -m "feat(llm): generate speaker notes for content slides into pptx notes area

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

README 的"LLM 增强"一节增强内容清单里补一句"每页演讲备注（写入 PPT 备注区，带 AI 标记）"。

---

## Phase 4 产物落地（域更新 / PPT 结构校验）

### Task 14: docx 设置 updateFields=true（打开即提示更新目录/页码）

**Files:**
- Modify: `thesis_project/src/docx_builder.py`（新增 `_enable_update_fields`，`build` 末尾调用）
- Test: `thesis_project/tests/test_docx_update_fields.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import docx
from docx.oxml.ns import qn

from src import docx_builder


def _min_thesis():
    return {"title": "题目", "author": "作者",
            "abstract": "本文摘要。", "abstract_en": "EN abstract.",
            "keywords": ["a", "b", "c"], "keywords_en": ["a", "b", "c"],
            "chapters": [{"title": "绪论", "level": 1, "paras": ["正文。"],
                          "subs": []}],
            "auto_skeleton": False, "references": ["某文献[J]. 2024."]}


def test_update_fields_flag_present(tmp_path):
    out = str(tmp_path / "o.docx")
    docx_builder.build(_min_thesis(), out)
    d = docx.Document(out)
    el = d.settings.element.find(qn("w:updateFields"))
    assert el is not None
    assert el.get(qn("w:val")) == "true"
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_docx_update_fields.py -v`
Expected: FAIL，`el is None`

- [ ] **Step 3: 实现**

在 `_add_toc` 附近新增（`OxmlElement` 已在 Task 8 引入；若 Task 8 未先做，则此处补 import）：

```python
def _enable_update_fields(doc):
    """settings.xml 写 updateFields=true：Word 打开时提示一键更新目录/页码域。"""
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is not None:
        return
    el = OxmlElement("w:updateFields")
    el.set(qn("w:val"), "true")
    settings.append(el)
```

`build` 中 `doc.save(out_path)`（原 311 行）之前加 `_enable_update_fields(doc)`。
同时把 `_add_toc` 内的提示文本 `"【在 Word 中按 F9 更新目录】"`（158 行）改为 `"【打开文档后按提示更新域，或按 F9】"`，README"完成"提示同步微调（main.py 106 行的收尾文案也提到 F9，一并改）。

- [ ] **Step 4: 运行确认通过 + 提交**

```bash
python -m pytest tests/test_docx_update_fields.py tests/test_docx_toc.py -v
python -m pytest tests/ -q
git add src/docx_builder.py src/main.py tests/test_docx_update_fields.py README.md
git commit -m "feat(docx): set updateFields=true so Word offers to refresh TOC/page fields on open

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 15: win32com 后处理：静默刷新域 + 可选导出 PDF

**Files:**
- Create: `thesis_project/src/postprocess.py`
- Modify: `thesis_project/src/main.py`（新增 `--refresh-fields` / `--pdf` 参数与调用）
- Modify: `thesis_project/requirements.txt`（追加 pywin32 条目）
- Test: `thesis_project/tests/test_postprocess.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import sys
import types

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
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_postprocess.py -v`
Expected: FAIL，`ModuleNotFoundError: src.postprocess`

- [ ] **Step 3: 实现 postprocess.py（完整新文件）**

```python
# -*- coding: utf-8 -*-
"""Word COM 后处理：静默更新目录/页码等域，可选导出 PDF。

仅 Windows + 本机安装 Word 时可用；任何失败都只打印提示、返回 False，
不影响主流程（docx 已带 updateFields 标记作为兜底）。
"""
from __future__ import annotations
import os

_WD_EXPORT_PDF = 17  # WdExportFormat.wdExportFormatPDF


def refresh_word_fields(docx_path: str, export_pdf: bool = False) -> bool:
    """用本机 Word 更新 docx 全部域（目录/页码/交叉引用）并保存。"""
    try:
        import win32com.client
    except ImportError:
        print("  [提示] 未安装 pywin32，跳过域刷新"
              "（打开 Word 后按提示更新域即可）：pip install pywin32")
        return False
    path = os.path.abspath(docx_path)
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(path)
        doc.Fields.Update()
        for i in range(1, doc.TablesOfContents.Count + 1):
            doc.TablesOfContents(i).Update()
        doc.Save()
        print("  ✔ 已用 Word 刷新目录/页码域")
        if export_pdf:
            pdf = os.path.splitext(path)[0] + ".pdf"
            doc.ExportAsFixedFormat(pdf, _WD_EXPORT_PDF)
            print(f"  ✔ PDF : {pdf}")
        return True
    except Exception as e:  # noqa: BLE001 COM 异常类型不可枚举
        print(f"  [提示] Word 域刷新失败（{e}），请打开文档后按 F9 手动更新")
        return False
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:  # noqa: BLE001
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:  # noqa: BLE001
            pass
```

- [ ] **Step 4: 接入 main.py**

argparse 增加两个参数（`--llm` 之后）：

```python
    ap.add_argument("--refresh-fields", action="store_true",
                    help="生成后用本机 Word 静默刷新目录/页码域（需已装 Word 和 pywin32）")
    ap.add_argument("--pdf", action="store_true",
                    help="刷新域后同时导出 PDF（隐含 --refresh-fields）")
```

Word 生成成功分支（Task 1 改造后的 `if wp:` 内）追加：

```python
        if wp and (args.refresh_fields or args.pdf):
            from src import postprocess
            postprocess.refresh_word_fields(wp, export_pdf=args.pdf)
```

`requirements.txt` 追加一行：

```
pywin32>=306; sys_platform == "win32"
```

- [ ] **Step 5: 运行确认通过 + 真机冒烟 + 提交**

```bash
python -m pytest tests/test_postprocess.py -v
python -m pytest tests/ -q
python src/main.py --input sample_input --output output_smoke --only word --refresh-fields
```

冒烟预期：本机装了 Word 时打印"已用 Word 刷新目录/页码域"，打开产物目录已生成、无 F9 提示占位；没装则打印降级提示且退出码仍为 0。完成后删除 `output_smoke/`。

```bash
git add src/postprocess.py src/main.py requirements.txt tests/test_postprocess.py README.md
git commit -m "feat(postprocess): optional win32com pass to refresh fields and export PDF

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

README"其它用法"代码块补 `--refresh-fields` / `--pdf` 两行说明。

---

### Task 16: deck 携带 bucket，补全 PPT 结构页数校验

**Files:**
- Modify: `thesis_project/src/organizer.py`（`_build_deck` 329-341 行的 section/content slide 加 `bucket`）
- Modify: `thesis_project/src/pptx_builder.py`（`_check_structure` 215-238 行补 bucket 页数校验，docstring 更新）
- Test: `thesis_project/tests/test_pptx_structure_buckets.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
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
```

（若 `PPT_SPEC["structure"]` 各段 min/max 与上面第三个用例的页数不匹配，按 `config/format_spec.py:162-170` 的实际值调整页数，使其全部落在区间内。）

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_pptx_structure_buckets.py -v`
Expected: FAIL（slide 无 bucket 键；无 bucket 警告输出）

- [ ] **Step 3: 实现 organizer**

`_build_deck` 中 4 处 slide 构造加 `"bucket": key`（329、330-331、333、340-341 行）：

```python
        if not group:
            slides.append({"type": "section", "title": label[key], "bucket": key})
            slides.append({"type": "content", "title": label[key],
                           "bullets": ["<待补充要点>"], "bucket": key})
            continue
        slides.append({"type": "section", "title": label[key], "bucket": key})
        ...
                slides.append({"type": "content", "title": title,
                               "bullets": chunk, "bucket": key})
```

- [ ] **Step 4: 实现 pptx_builder**

`_check_structure` 末尾追加（并把 docstring 里"无分段归属标记而不校验"的说明改为"content 页按 bucket 校验"）：

```python
    label_by_key = {seg["key"]: seg for seg in P["structure"]}
    bucket_counts = {}
    for s in slides:
        if s.get("type") == "content" and s.get("bucket"):
            bucket_counts[s["bucket"]] = bucket_counts.get(s["bucket"], 0) + 1
    for key in ("background", "method", "result", "conclusion"):
        seg = label_by_key.get(key)
        if seg is None:
            continue
        n = bucket_counts.get(key, 0)
        if n < seg.get("min", 0) or n > seg.get("max", 99):
            print(f"  [警告] 「{seg['title']}」内容页 {n} 页，"
                  f"规范建议 {seg.get('min')}~{seg.get('max')} 页")
```

- [ ] **Step 5: 运行确认通过 + 提交**

```bash
python -m pytest tests/test_pptx_structure_buckets.py tests/test_pptx_page_count.py tests/test_pptx_warnings.py tests/test_deck_overflow.py -v
python -m pytest tests/ -q
git add src/organizer.py src/pptx_builder.py tests/test_pptx_structure_buckets.py
git commit -m "feat(pptx): per-bucket page-count validation via bucket field on deck slides

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Phase 5 清理与测试补强

### Task 17: 标注 format_spec 未落实字段 + README 对齐

> **回顾修订（2026-07-23）**：本任务原始清单中的 `page.size`、
> `page.orientation`、`layout.max_lines_per_bullet`、`layout.content_max_ratio`
> 和 `layout.text_align` 已接入生成流程，并支持 YAML 覆盖；下列内容保留为当时
> 的实施步骤，当前状态以 README 和 `config/template.py` 为准。

**Files:**
- Modify: `thesis_project/config/format_spec.py`
- Modify: `thesis_project/README.md`

- [ ] **Step 1: 给未被代码读取的字段加行尾注释 `# 暂未落实`**

逐项核对（Task 8/16 完成后 `figure`/`table`/`structure` 已激活，无需标注）：
- `WORD_SPEC`：`headings.*.outline_level`、`toc.leader`、`toc.page_number_align`、`reference.standard`
- `PPT_SPEC`：`slide.ratio`、`font.max_font_kinds`、`layout.background`、`principle.talk_minutes`、`principle.rule`、`principle.narrative`

- [ ] **Step 2: README"自定义"一节补一句**

"仍标注`# 暂未落实`的字段目前仅作文档参考；已经接入生成器的字段可通过 YAML 模板生效。"

- [ ] **Step 3: 全量测试 + 提交**

```bash
python -m pytest tests/ -q
git add config/format_spec.py README.md
git commit -m "docs(spec): mark not-yet-implemented spec fields to avoid misleading customization

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

### Task 18: 删除注释断言测试；e2e 补内容级断言

**Files:**
- Delete: `thesis_project/tests/test_readers_docstring.py`（用 inspect 断言 docstring 措辞，属负价值测试；其行为已被 test_readers_encoding.py 覆盖）
- Modify: `thesis_project/tests/test_e2e.py`

- [ ] **Step 1: 在 test_e2e.py 生成产物后追加内容级断言**

先读 test_e2e.py 确认产物路径变量名，然后在"文件大小 > 0"断言之后追加（变量名按实际替换）：

```python
    import docx as _docx
    from pptx import Presentation as _P

    d = _docx.Document(word_path)
    all_text = "\n".join(p.text for p in d.paragraphs)
    assert "参考文献" in all_text
    assert "[1]" in all_text                      # 参考文献编号
    assert "关键词" in all_text

    prs = _P(ppt_path)
    cover_texts = [sh.text_frame.text for sh in prs.slides[0].shapes
                   if sh.has_text_frame]
    assert any(t.strip() for t in cover_texts)    # 封面非空
    assert len(prs.slides) >= 5
```

- [ ] **Step 2: 运行确认通过**

Run: `python -m pytest tests/test_e2e.py -v`
Expected: PASS（若断言文本与实际产物有出入，打开产物核实后修正断言，而不是放宽到无意义）

- [ ] **Step 3: 删除负价值测试 + 提交**

```bash
git rm tests/test_readers_docstring.py
python -m pytest tests/ -q
git add tests/test_e2e.py
git commit -m "test: content-level e2e assertions; drop docstring-wording test

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## 收尾验证（全部任务完成后）

- [ ] `python -m pytest tests/ -q` 全绿
- [ ] `python src/main.py --input sample_input` 冒烟：两份产物生成、控制台无 traceback
- [ ] 打开 `论文草案.docx`：确认弹出"更新域"提示（或 `--refresh-fields` 后目录/页码直接就绪）
- [ ] 用一个含表格与图片的 docx 源文件跑一遍：确认三线表、图片与`表1-1`/`图1-1`题注出现在产物中
- [ ] `git log --oneline` 确认每个任务一个提交

## 遗留候选（本计划不做，按需另立计划）

1. **PPT 模板填充引擎**（`--template xxx.pptx`，复制学校模板原生页替换内容）——参考 PPTAgent、thesis-defense-pptx-skill 的路线，工作量大，值得单独规划。
2. **高级 GB/T 7714 CSL 排版**——当前已提供离线确定性 formatter；如需 citeproc-py /
   Zotero 中文 CSL 样式，可另立计划升级著录规则。
3. **`--check-only` 格式/内容检查器**——规则集外置 YAML、分级输出，参考 thesis-format-checker。
4. **PDF 页眉页脚噪声过滤**（按坐标裁掉页面顶/底 5% 或去除每页重复行）与英文断词连字符合并。
5. **多文件输入同名章节合并** + README 说明按文件名排序。
6. **封面元信息透传**（学院/专业/指导教师经 frontmatter 进封面）。
7. **复杂 PDF 可选接入 MinerU/marker** 输出结构化 Markdown。

## 2026-07-23 验收补充

- [x] 全量 pytest：264 项通过。
- [x] `src/config` Ruff 检查与 Python 编译检查通过。
- [x] 普通模式 CLI 生成 Word/PPT、`--dry-run`、YAML 模板加载均已冒烟。
- [ ] 真实 LLM/Crossref、Windows 双击 `run.bat`、Word COM 域刷新仍需在目标环境人工验证。
