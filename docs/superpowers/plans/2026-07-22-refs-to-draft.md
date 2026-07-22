# 参考资料 → 论文初稿（研究写作辅助）实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新增"参考资料模式"：`input\` 放题目文件 + 参考文献/Excel/截图时，LLM 生成带文献综述、大纲、写作要点与 GB/T 7714 参考文献表的 `论文草案.docx`（不写核心研究章节正文，不生成 PPT）。

**Architecture:** 扩展 readers（xlsx/csv/图片）；新建 `synthesizer.py` 作为 organizer 的平级替代品，输出与现有完全同构的 thesis dict，docx_builder/pptx_builder 零改动；新建 `llm_vision.py` 承载视觉调用；main.py 按 topic 文件自动分流，`--mode` 可覆盖。

**Tech Stack:** Python 3.13、python-docx、pdfplumber、openai>=1.0、openpyxl（新增）、pytest。

**与 spec 的一处偏差（已确认更优）:** spec 3 节原计划把 `_chat/_chat_json` 抽到 `llm_client.py`。但现有测试（test_llm_base/test_llm_json_mode 等十余个文件）都 patch `llm_enhancer._chat` 并依赖同模块内 `_chat_json → _chat → _client` 的属性查找链，抽走必断。改为：`llm_enhancer` 文本函数**原地不动**，仅把 JSON 解析逻辑提为 `_parse_json` 复用；新建 `src/llm_vision.py` 只装视觉调用。spec 将同步修订。

**执行约定：**
- 所有命令在 `thesis_project/` 目录下执行；测试命令为 `python -m pytest tests/<文件> -v`。
- 每个任务完成后跑一次全量 `python -m pytest tests/ -q` 再提交，防止跨模块回归。
- 提交信息沿用仓库风格（`feat(readers): ...` 等），并以下行结尾：
  `Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`
- 测试文件头部加 `# -*- coding: utf-8 -*-`；import 用 `from src import xxx`（conftest.py 已处理 sys.path）。

**关键现状事实（2026-07-22 版本，写码前核对）：**
- `readers.py:387-395` `_READERS` 目前只有 txt/text/md/markdown/json/docx/pdf；`_block()` 在 35 行，`_read_text()`（三段编码回退）在 43 行，`_clean()` 在 39 行。
- `organizer.py:199-202` `_node()` 定义章节节点 `{"title", "level", "paras", "subs", "tables", "images"}`；`PLACEHOLDER = "<请填写>"` 在 21 行；`DEFAULT_CHAPTERS` 在 26 行。
- `llm_enhancer.py:42-60` `_chat`（唯一文本网络出口）、`63-74` `_chat_json`（内联 JSON 解析）、`26` `AI_MARK`、`30-31` `is_available`、`34-39` `_client`。
- `main.py:35-61` `gather_docs`、`64-79` `_build_with_retry`、`82-139` `main()`（argparse 在 83-95）。
- thesis dict 必备键（`organizer.organize` 331-351 行）：title/author/abstract/abstract_en/keywords/keywords_en/chapters/auto_skeleton/references。
- 测试打桩惯例：`monkeypatch.setattr(llm_enhancer, "_chat_json", fake)`——新模块同样把 `_chat_json` import 进自己命名空间供 patch。

---

## Phase 1 读取器扩展

### Task 1: read_xlsx（openpyxl）

**Files:**
- Modify: `thesis_project/requirements.txt`（加 `openpyxl`）
- Modify: `thesis_project/src/readers.py`（新增 `read_xlsx`，注册 `.xlsx`）
- Test: `thesis_project/tests/test_readers_xlsx.py`（新建）

- [ ] **Step 1: 安装依赖并写失败测试**

Run: `python -m pip install openpyxl`，并在 `requirements.txt` 追加一行 `openpyxl`。

```python
# -*- coding: utf-8 -*-
import openpyxl
import pytest
from src import readers


def _make_wb(path, sheets):
    """sheets: {表名: [[行], ...]}；写一个真实 xlsx 供读取。"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


def test_read_xlsx_each_nonempty_sheet_becomes_table(tmp_path):
    p = str(tmp_path / "数据.xlsx")
    _make_wb(p, {"实验": [["组别", "精度"], ["A", 0.91]], "空表": [], "问卷": [["Q1", 5]]})
    doc = readers.read_xlsx(p)
    assert doc["type"] == "xlsx"
    tables = [b for b in doc["blocks"] if b["kind"] == "table"]
    assert len(tables) == 2                      # 空表被跳过
    assert tables[0]["rows"][0] == ["组别", "精度"]
    assert tables[0]["rows"][1] == ["A", "0.91"]  # 数值转字符串


def test_read_xlsx_none_cells_become_empty_string(tmp_path):
    p = str(tmp_path / "n.xlsx")
    _make_wb(p, {"s": [["a", None, "c"]]})
    doc = readers.read_xlsx(p)
    assert doc["blocks"][0]["rows"][0] == ["a", "", "c"]


def test_read_xlsx_all_empty_raises(tmp_path):
    p = str(tmp_path / "e.xlsx")
    _make_wb(p, {"s1": [], "s2": [[None, None]]})
    with pytest.raises(RuntimeError):
        readers.read_xlsx(p)


def test_read_xlsx_registered_in_dispatch(tmp_path):
    p = str(tmp_path / "d.xlsx")
    _make_wb(p, {"s": [["x"]]})
    assert readers.read_file(p)["type"] == "xlsx"
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_readers_xlsx.py -v`
Expected: FAIL，`AttributeError: module 'src.readers' has no attribute 'read_xlsx'`

- [ ] **Step 3: 实现 read_xlsx**

在 `readers.py` 的 `read_pdf` 之后、`_READERS` 之前插入：

```python
# ---------------------------------------------------------------------------
#  XLSX / CSV
# ---------------------------------------------------------------------------
def read_xlsx(path: str) -> dict:
    """每个非空工作表 -> 一个 table 块。合并单元格的非锚点格读出 None -> ""。

    data_only=True 读公式的缓存计算值；文件从未被 Excel 打开过时可能为
    None，同样落为空串，属可接受降级。
    """
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("需要 openpyxl：pip install openpyxl")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    blocks = []
    try:
        for ws in wb.worksheets:
            rows = []
            for row in ws.iter_rows(values_only=True):
                cells = ["" if c is None else _clean(str(c)) for c in row]
                if any(cells):
                    rows.append(cells)
            if rows:
                blocks.append(_block("table", "", rows=rows))
    finally:
        wb.close()
    if not blocks:
        raise RuntimeError("工作簿中没有任何非空工作表")
    return {"source": path, "type": "xlsx", "blocks": blocks, "meta": {}}
```

并在 `_READERS` 字典中加入 `".xlsx": read_xlsx,`。

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_readers_xlsx.py -v`
Expected: 4 PASS

- [ ] **Step 5: 全量回归 + 提交**

```bash
python -m pytest tests/ -q
git add requirements.txt src/readers.py tests/test_readers_xlsx.py
git commit -m "feat(readers): read xlsx workbooks as table blocks via openpyxl"
```

### Task 2: read_csv

**Files:**
- Modify: `thesis_project/src/readers.py`
- Test: `thesis_project/tests/test_readers_csv.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import pytest
from src import readers


def test_read_csv_basic(tmp_path):
    p = tmp_path / "数据.csv"
    p.write_text("组别,精度\nA,0.91\n,\n", encoding="utf-8")
    doc = readers.read_csv(str(p))
    assert doc["type"] == "csv"
    assert doc["blocks"][0]["kind"] == "table"
    assert doc["blocks"][0]["rows"] == [["组别", "精度"], ["A", "0.91"]]  # 全空行跳过


def test_read_csv_gbk_fallback(tmp_path):
    p = tmp_path / "g.csv"
    p.write_bytes("名称,数值\n测试,1\n".encode("gb18030"))
    doc = readers.read_csv(str(p))
    assert doc["blocks"][0]["rows"][0] == ["名称", "数值"]


def test_read_csv_empty_raises(tmp_path):
    p = tmp_path / "e.csv"
    p.write_text("\n,\n", encoding="utf-8")
    with pytest.raises(RuntimeError):
        readers.read_csv(str(p))


def test_read_csv_registered(tmp_path):
    p = tmp_path / "d.csv"
    p.write_text("x\n", encoding="utf-8")
    assert readers.read_file(str(p))["type"] == "csv"
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_readers_csv.py -v`
Expected: FAIL，`has no attribute 'read_csv'`

- [ ] **Step 3: 实现 read_csv**

在 `read_xlsx` 之后插入（复用 `_read_text` 的三段编码回退）：

```python
def read_csv(path: str) -> dict:
    """整个 CSV -> 单个 table 块；全空行跳过。编码回退复用 _read_text。"""
    import csv
    import io
    raw = _read_text(path)
    rows = []
    for row in csv.reader(io.StringIO(raw)):
        cells = [_clean(c) for c in row]
        if any(cells):
            rows.append(cells)
    if not rows:
        raise RuntimeError("CSV 中没有任何非空行")
    return {"source": path, "type": "csv",
            "blocks": [_block("table", "", rows=rows)], "meta": {}}
```

`_READERS` 加入 `".csv": read_csv,`。

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_readers_csv.py -v`
Expected: 4 PASS

- [ ] **Step 5: 全量回归 + 提交**

```bash
python -m pytest tests/ -q
git add src/readers.py tests/test_readers_csv.py
git commit -m "feat(readers): read csv files as a single table block"
```

### Task 3: read_image（独立图片文件）

**Files:**
- Modify: `thesis_project/src/readers.py`
- Test: `thesis_project/tests/test_readers_image.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import pytest
from src import readers

# 1x1 红点 PNG（最小合法 PNG 字节）
PNG_BYTES = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108020000009077"
    "3df80000000c4944415408d763f8cfc00000030101cf9e46a80000000049454e44ae426082")


def test_read_image_returns_single_image_block(tmp_path):
    p = tmp_path / "架构图.png"
    p.write_bytes(PNG_BYTES)
    doc = readers.read_image(str(p))
    assert doc["type"] == "image"
    assert len(doc["blocks"]) == 1
    b = doc["blocks"][0]
    assert b["kind"] == "image"
    assert b["data"] == PNG_BYTES
    assert b["ext"] == ".png"


def test_read_image_empty_file_raises(tmp_path):
    p = tmp_path / "空.jpg"
    p.write_bytes(b"")
    with pytest.raises(RuntimeError):
        readers.read_image(str(p))


@pytest.mark.parametrize("ext", [".png", ".jpg", ".jpeg", ".bmp", ".webp"])
def test_image_extensions_registered(tmp_path, ext):
    p = tmp_path / ("t" + ext)
    p.write_bytes(PNG_BYTES)
    assert readers.read_file(str(p))["type"] == "image"
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_readers_image.py -v`
Expected: FAIL，`has no attribute 'read_image'`

- [ ] **Step 3: 实现 read_image**

在 `read_csv` 之后插入：

```python
# ---------------------------------------------------------------------------
#  独立图片（截图等）
# ---------------------------------------------------------------------------
def read_image(path: str) -> dict:
    """整个文件 -> 单个 image 块（原始字节 + 扩展名），供插图与视觉理解。"""
    with open(path, "rb") as f:
        data = f.read()
    if not data:
        raise RuntimeError("图片文件为空")
    b = _block("image")
    b["data"] = data
    b["ext"] = os.path.splitext(path)[1].lower() or ".png"
    return {"source": path, "type": "image", "blocks": [b], "meta": {}}
```

`_READERS` 加入：

```python
    ".png": read_image,
    ".jpg": read_image,
    ".jpeg": read_image,
    ".bmp": read_image,
    ".webp": read_image,
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_readers_image.py -v`
Expected: 7 PASS

- [ ] **Step 5: 全量回归 + 提交**

```bash
python -m pytest tests/ -q
git add src/readers.py tests/test_readers_image.py
git commit -m "feat(readers): import standalone images (png/jpg/jpeg/bmp/webp)"
```

## Phase 2 配置与 LLM 基础设施

### Task 4: REFS_SPEC 配置块

**Files:**
- Modify: `thesis_project/config/format_spec.py`（文件末尾追加）
- Test: `thesis_project/tests/test_refs_spec.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from config.format_spec import REFS_SPEC


def test_refs_spec_has_required_keys():
    assert set(REFS_SPEC) >= {"topic_filenames", "default_outline",
                              "review_notice", "materials_chapter"}


def test_topic_filenames_lowercase():
    # main.py 用 lower() 比对，配置里必须全小写
    assert all(n == n.lower() for n in REFS_SPEC["topic_filenames"])


def test_default_outline_kinds_valid():
    kinds = {c["kind"] for c in REFS_SPEC["default_outline"]}
    assert kinds <= {"intro", "review", "core", "conclusion"}
    assert "review" in kinds and "core" in kinds
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_refs_spec.py -v`
Expected: FAIL，`ImportError: cannot import name 'REFS_SPEC'`

- [ ] **Step 3: 在 format_spec.py 末尾追加**

```python
# =============================================================================
#  三、参考资料模式（研究写作辅助）
# =============================================================================

REFS_SPEC = {
    # input/ 下存在这些文件名之一（不区分大小写）即触发参考资料模式
    "topic_filenames": ["topic.md", "topic.txt", "题目.md", "题目.txt"],

    # 大纲生成失败时的降级骨架；kind: intro|review|core|conclusion
    # review 章由 LLM 撰写综述正文；其余章只给写作要点 + 占位符
    "default_outline": [
        {"title": "绪论", "kind": "intro"},
        {"title": "相关理论与技术", "kind": "review"},
        {"title": "国内外研究现状", "kind": "review"},
        {"title": "研究内容与方案设计", "kind": "core"},
        {"title": "系统实现与结果分析", "kind": "core"},
        {"title": "总结与展望", "kind": "conclusion"},
    ],

    # 插在每个综述章开头的醒目提示
    "review_notice": "本章由 AI 基于所给文献生成，请逐条核对原文后改写为自己的表述。",

    # 无法按语义匹配到章节的表格/截图统一挂到这一章
    "materials_chapter": "素材附录（整理用，定稿前删除）",
}
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_refs_spec.py -v`
Expected: 3 PASS

- [ ] **Step 5: 提交**

```bash
python -m pytest tests/ -q
git add config/format_spec.py tests/test_refs_spec.py
git commit -m "feat(config): add REFS_SPEC for reference-materials mode"
```

### Task 5: llm_enhancer._parse_json 提取（微重构，行为不变）

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py:63-74`（`_chat_json` 拆出解析函数）
- Test: `thesis_project/tests/test_llm_parse_json.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import pytest
from src import llm_enhancer


def test_parse_json_plain():
    assert llm_enhancer._parse_json('{"a": 1}') == {"a": 1}


def test_parse_json_fenced_with_prose():
    text = '好的，如下：\n```json\n{"a": [1, 2]}\n```\n以上。'
    assert llm_enhancer._parse_json(text) == {"a": [1, 2]}


def test_parse_json_no_json_raises():
    with pytest.raises(ValueError):
        llm_enhancer._parse_json("抱歉，我无法处理。")
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_llm_parse_json.py -v`
Expected: FAIL，`has no attribute '_parse_json'`

- [ ] **Step 3: 重构 _chat_json**

把 `llm_enhancer.py` 的 `_chat_json`（63-74 行）改为：

```python
def _parse_json(text: str):
    """容忍代码块包裹与前后废话的 JSON 解析。供 _chat_json 与视觉模块复用。"""
    m = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL)
    if m:
        text = m.group(1)
    starts = [i for i in (text.find("{"), text.find("[")) if i >= 0]
    if not starts:
        raise ValueError(f"LLM 未返回 JSON：{text[:80]!r}")
    obj, _ = json.JSONDecoder().raw_decode(text[min(starts):].strip())
    return obj


def _chat_json(system: str, user: str):
    """要求模型输出 JSON 并解析；容忍代码块包裹与前后废话。"""
    text = _chat(system + "\n只输出 JSON，不要任何其它文字。", user,
                 json_mode=True)
    return _parse_json(text)
```

注意：`_chat_json` 的对外行为、打桩点（`llm_enhancer._chat`）完全不变。

- [ ] **Step 4: 运行确认通过（含既有 LLM 测试回归）**

Run: `python -m pytest tests/test_llm_parse_json.py tests/test_llm_base.py tests/test_llm_json_mode.py -v`
Expected: 全 PASS

- [ ] **Step 5: 全量回归 + 提交**

```bash
python -m pytest tests/ -q
git add src/llm_enhancer.py tests/test_llm_parse_json.py
git commit -m "refactor(llm): extract _parse_json from _chat_json for reuse"
```

### Task 6: llm_vision.py（视觉调用，第二个网络出口）

**Files:**
- Create: `thesis_project/src/llm_vision.py`
- Test: `thesis_project/tests/test_llm_vision.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
import pytest
from src import llm_vision


def test_unavailable_without_vision_model(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.delenv("LLM_VISION_MODEL", raising=False)
    assert llm_vision.is_vision_available() is False


def test_available_with_both_envs(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("LLM_VISION_MODEL", "qwen-vl-plus")
    assert llm_vision.is_vision_available() is True


def test_describe_image_parses_json(monkeypatch):
    captured = {}

    def fake_chat_vision(prompt, image_b64, mime):
        captured["mime"] = mime
        return '{"caption": "系统架构图", "summary": "三层架构：表现层…"}'

    monkeypatch.setattr(llm_vision, "_chat_vision", fake_chat_vision)
    out = llm_vision.describe_image(b"\x89PNG...", ".png")
    assert out == {"caption": "系统架构图", "summary": "三层架构：表现层…"}
    assert captured["mime"] == "image/png"


def test_describe_image_unknown_ext_defaults_png(monkeypatch):
    captured = {}

    def fake_chat_vision(prompt, image_b64, mime):
        captured["mime"] = mime
        return '{"caption": "c", "summary": "s"}'

    monkeypatch.setattr(llm_vision, "_chat_vision", fake_chat_vision)
    llm_vision.describe_image(b"x", ".tiff")
    assert captured["mime"] == "image/png"


def test_describe_image_bad_json_raises(monkeypatch):
    monkeypatch.setattr(llm_vision, "_chat_vision", lambda *a: "看不懂")
    with pytest.raises(ValueError):
        llm_vision.describe_image(b"x", ".png")
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_llm_vision.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'src.llm_vision'`

- [ ] **Step 3: 实现 llm_vision.py**

```python
# -*- coding: utf-8 -*-
"""
视觉理解模块 —— 用 OpenAI 兼容多模态端点理解截图（可选）。

配置（环境变量）：
    LLM_API_KEY       复用文本模型的密钥（必填）
    LLM_VISION_MODEL  多模态模型名（如 gpt-4o / qwen-vl-plus / glm-4v）；
                      未设置则视觉理解整体跳过，截图仅作插图。
    LLM_BASE_URL / LLM_TIMEOUT  与文本调用共用。

原则：
  - _chat_vision 是本模块唯一网络出口，测试打桩即可。
  - 失败抛异常，由调用方（synthesizer）降级为"仅插图"。
"""
from __future__ import annotations
import base64
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.llm_enhancer import _parse_json

_MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
         ".bmp": "image/bmp", ".webp": "image/webp"}

_VISION_PROMPT = (
    "这是论文写作素材截图。请以 JSON 输出 "
    '{"caption": "适合作论文图题的一句话", "summary": "图中信息摘要，100字以内"}。'
    "只描述图中真实可见的内容，不得编造。只输出 JSON。")


def is_vision_available() -> bool:
    return bool(os.environ.get("LLM_API_KEY")) and \
        bool(os.environ.get("LLM_VISION_MODEL"))


def _chat_vision(prompt: str, image_b64: str, mime: str) -> str:
    """唯一网络出口：单图 + 文本 -> 原始回复文本。"""
    from openai import OpenAI
    timeout = float(os.environ.get("LLM_TIMEOUT", "60"))
    client = OpenAI(api_key=os.environ["LLM_API_KEY"],
                    base_url=os.environ.get("LLM_BASE_URL") or None,
                    timeout=timeout, max_retries=1)
    resp = client.chat.completions.create(
        model=os.environ["LLM_VISION_MODEL"],
        messages=[{"role": "user", "content": [
            {"type": "text", "text": prompt},
            {"type": "image_url",
             "image_url": {"url": f"data:{mime};base64,{image_b64}"}},
        ]}],
        temperature=0.2)
    return resp.choices[0].message.content or ""


def describe_image(data: bytes, ext: str) -> dict:
    """截图 -> {"caption": 图题建议, "summary": 内容摘要}；失败抛异常。"""
    b64 = base64.b64encode(data).decode("ascii")
    text = _chat_vision(_VISION_PROMPT, b64, _MIME.get(ext, "image/png"))
    obj = _parse_json(text)
    if not isinstance(obj, dict):
        raise ValueError("视觉理解结果不是 JSON 对象")
    return {"caption": str(obj.get("caption", "")).strip(),
            "summary": str(obj.get("summary", "")).strip()}
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_llm_vision.py -v`
Expected: 5 PASS

- [ ] **Step 5: 全量回归 + 提交**

```bash
python -m pytest tests/ -q
git add src/llm_vision.py tests/test_llm_vision.py
git commit -m "feat(llm): add optional vision module for screenshot understanding"
```

## Phase 3 综合器（synthesizer.py）

> 本阶段每个任务向 `src/synthesizer.py` 增量添加函数。Task 7 建立模块骨架，
> 后续任务只追加。所有 LLM 调用经模块内名字 `_chat_json`（从 llm_enhancer
> import 进来），测试统一 `monkeypatch.setattr(synthesizer, "_chat_json", fake)`。

### Task 7: 模块骨架 + parse_topic

**Files:**
- Create: `thesis_project/src/synthesizer.py`
- Test: `thesis_project/tests/test_synth_topic.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import synthesizer
from src.organizer import PLACEHOLDER


def _doc(blocks, meta=None, source="input/topic.md"):
    return {"source": source, "type": "md", "blocks": blocks, "meta": meta or {}}


def _h(text, level=1):
    return {"kind": "heading", "level": level, "text": text, "rows": None}


def _p(text):
    return {"kind": "paragraph", "level": 0, "text": text, "rows": None}


def test_parse_topic_title_from_first_heading():
    doc = _doc([_h("基于YOLOv8的课堂行为识别系统"), _p("研究内容：检测举手、低头等行为。")])
    t = synthesizer.parse_topic(doc)
    assert t["title"] == "基于YOLOv8的课堂行为识别系统"
    assert "举手" in t["background"]


def test_parse_topic_title_from_first_paragraph_when_no_heading():
    doc = _doc([_p("基于知识图谱的推荐系统研究"), _p("拟采用图神经网络。")])
    t = synthesizer.parse_topic(doc)
    assert t["title"] == "基于知识图谱的推荐系统研究"


def test_parse_topic_author_from_frontmatter():
    doc = _doc([_h("题目X")], meta={"author": "张三"})
    assert synthesizer.parse_topic(doc)["author"] == "张三"


def test_parse_topic_defaults_placeholder():
    doc = _doc([])
    t = synthesizer.parse_topic(doc)
    assert t["title"] == PLACEHOLDER
    assert t["author"] == PLACEHOLDER
    assert t["background"] == ""
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_synth_topic.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'src.synthesizer'`

- [ ] **Step 3: 创建 synthesizer.py 骨架**

```python
# -*- coding: utf-8 -*-
"""
参考资料综合器 —— 参考资料模式下 organizer 的平级替代品。

输入：题目 Document + 参考资料 Document 列表
输出：与 organizer.organize() 完全同构的 thesis dict（docx_builder 直接消费）

管道（见 docs/superpowers/specs/2026-07-22-refs-to-draft-design.md）：
  ① make_cards       逐篇文献 -> 摘要卡（逐篇容错）
  ② build_outline    题目 + 摘要卡 -> 章节大纲（失败退 REFS_SPEC 默认骨架）
  ③ write_review     综述章正文（带 [n] 引用与 AI 标记；失败留素材+占位）
  ④ write_points     核心章写作要点（批量；失败留占位）
  ⑤ format_references 摘要卡 -> GB/T 7714（失败罗列原始标题）
  ⑥ attach_media     xlsx/csv/截图挂载（纯规则）

原则：
  - LLM 生成综述与要点，不代写核心研究章节正文（学术诚信边界）。
  - 所有 AI 正文带 AI_MARK；任何步骤降级都记入 _degraded 并汇总打印。
  - 网络出口复用 llm_enhancer._chat_json / llm_vision._chat_vision，
    测试打桩 synthesizer._chat_json / llm_vision._chat_vision 即可。
"""
from __future__ import annotations
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.format_spec import REFS_SPEC
from src.llm_enhancer import AI_MARK, _chat_json
from src.organizer import PLACEHOLDER, _node

_VALID_KINDS = {"intro", "review", "core", "conclusion"}
_MEDIA_TYPES = ("xlsx", "csv", "image")

# 本次 synthesize() 调用中发生的降级记录（步骤名列表）
_degraded = []


def _note_degrade(step: str, err) -> None:
    _degraded.append(step)
    print(f"  [LLM告警] {step}失败，已降级：{err}")


def _doc_text(doc: dict, limit: int = 6000) -> str:
    """Document -> 纯文本（标题与段落顺序拼接，表格拍平），截断到 limit。"""
    parts = []
    for b in doc["blocks"]:
        if b.get("text"):
            parts.append(b["text"])
        elif b.get("kind") == "table" and b.get("rows"):
            parts.append("\n".join(
                " | ".join(c for c in r if c and c.strip())
                for r in b["rows"] if any(c.strip() for c in r if c)))
    return "\n".join(parts)[:limit]


# ---------------------------------------------------------------------------
#  题目解析
# ---------------------------------------------------------------------------
def parse_topic(topic_doc: dict) -> dict:
    """题目文件 -> {"title", "author", "background"}。

    title：meta.title > 首个一级标题 > 首个非空段落首行 > 占位符。
    background：全文文本（喂给大纲/综述做背景），截断 4000 字符。
    """
    meta = topic_doc.get("meta") or {}
    title = (meta.get("title") or "").strip()
    if not title:
        for b in topic_doc["blocks"]:
            if b["kind"] == "heading" and b["level"] <= 1 and b["text"]:
                title = b["text"]
                break
    if not title:
        for b in topic_doc["blocks"]:
            if b["kind"] == "paragraph" and b["text"]:
                title = b["text"].splitlines()[0].strip()
                break
    return {"title": title or PLACEHOLDER,
            "author": (meta.get("author") or "").strip() or PLACEHOLDER,
            "background": _doc_text(topic_doc, limit=4000)}
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_synth_topic.py -v`
Expected: 4 PASS

- [ ] **Step 5: 提交**

```bash
python -m pytest tests/ -q
git add src/synthesizer.py tests/test_synth_topic.py
git commit -m "feat(synthesizer): module skeleton and topic-file parsing"
```

### Task 8: make_cards（逐篇摘要卡，逐篇容错）

**Files:**
- Modify: `thesis_project/src/synthesizer.py`（追加）
- Test: `thesis_project/tests/test_synth_cards.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import synthesizer


def _ref(source, text):
    return {"source": source, "type": "pdf", "meta": {},
            "blocks": [{"kind": "paragraph", "level": 0, "text": text,
                        "rows": None}]}


def test_make_cards_one_per_doc(monkeypatch):
    def fake_chat_json(system, user):
        return {"title": "YOLO综述", "authors": ["李四"], "year": "2023",
                "topic": "目标检测", "method": "综述", "conclusion": "有效",
                "quotes": ["单阶段检测器速度快"]}
    monkeypatch.setattr(synthesizer, "_chat_json", fake_chat_json)
    cards = synthesizer.make_cards([_ref("a.pdf", "正文A"), _ref("b.pdf", "正文B")])
    assert len(cards) == 2
    assert cards[0]["title"] == "YOLO综述"
    assert cards[0]["source"] == "a.pdf"
    assert cards[0]["quotes"] == ["单阶段检测器速度快"]


def test_make_cards_failed_doc_falls_back_to_text_card(monkeypatch, capsys):
    calls = []

    def flaky(system, user):
        calls.append(1)
        if len(calls) == 1:
            raise RuntimeError("接口超时")
        return {"title": "B文", "authors": [], "year": "", "topic": "t",
                "method": "", "conclusion": "", "quotes": []}
    monkeypatch.setattr(synthesizer, "_chat_json", flaky)
    cards = synthesizer.make_cards(
        [_ref("坏.pdf", "这篇失败了" * 100), _ref("好.pdf", "ok")])
    assert len(cards) == 2                       # 失败篇不丢
    assert cards[0]["title"] == "坏.pdf"          # 退化卡用文件名当标题
    assert cards[0]["fallback_text"].startswith("这篇失败了")
    assert "摘要卡" in capsys.readouterr().out    # 打了降级告警


def test_make_cards_normalizes_missing_fields(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json",
                        lambda s, u: {"title": "X"})   # 其余字段全缺
    card = synthesizer.make_cards([_ref("x.pdf", "t")])[0]
    assert card["authors"] == [] and card["quotes"] == []
    assert card["year"] == "" and card["method"] == ""
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_synth_cards.py -v`
Expected: FAIL，`has no attribute 'make_cards'`

- [ ] **Step 3: 追加实现**

```python
# ---------------------------------------------------------------------------
#  ① 文献摘要卡
# ---------------------------------------------------------------------------
_CARD_SYS = ("你是文献调研助手。阅读给定的单篇文献片段，抽取结构化信息。"
             "只能从原文归纳，禁止编造；无法确定的字段输出空字符串或空列表。")


def make_cards(ref_docs: list) -> list:
    """逐篇 -> 摘要卡；单篇失败退化为文本卡，不影响其它篇。

    卡片字段：title/authors/year/topic/method/conclusion/quotes/source；
    退化卡额外带 fallback_text（原文截断），title 用文件名。
    """
    cards = []
    for d in ref_docs:
        name = os.path.basename(d["source"])
        try:
            data = _chat_json(
                _CARD_SYS,
                '请以 JSON 输出 {"title": "...", "authors": ["..."], '
                '"year": "...", "topic": "一句话主题", "method": "...", '
                '"conclusion": "...", "quotes": ["可直接引用的关键观点"]}：'
                "\n\n" + _doc_text(d))
            if not isinstance(data, dict):
                raise ValueError("摘要卡结果不是 JSON 对象")
            cards.append({
                "title": str(data.get("title") or name).strip(),
                "authors": [str(a).strip() for a in data.get("authors") or []
                            if str(a).strip()],
                "year": str(data.get("year") or "").strip(),
                "topic": str(data.get("topic") or "").strip(),
                "method": str(data.get("method") or "").strip(),
                "conclusion": str(data.get("conclusion") or "").strip(),
                "quotes": [str(q).strip() for q in data.get("quotes") or []
                           if str(q).strip()],
                "source": name,
            })
        except Exception as e:  # noqa: BLE001
            _note_degrade(f"《{name}》摘要卡", e)
            cards.append({"title": name, "authors": [], "year": "",
                          "topic": "", "method": "", "conclusion": "",
                          "quotes": [], "source": name,
                          "fallback_text": _doc_text(d, limit=500)})
    return cards
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_synth_cards.py -v`
Expected: 3 PASS

- [ ] **Step 5: 提交**

```bash
python -m pytest tests/ -q
git add src/synthesizer.py tests/test_synth_cards.py
git commit -m "feat(synthesizer): per-document literature cards with fallback"
```

### Task 9: build_outline（大纲生成，失败退默认骨架）

**Files:**
- Modify: `thesis_project/src/synthesizer.py`（追加）
- Test: `thesis_project/tests/test_synth_outline.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import synthesizer
from config.format_spec import REFS_SPEC

TOPIC = {"title": "基于X的Y系统", "author": "张三", "background": "研究背景……"}
CARDS = [{"title": "文A", "topic": "主题A", "source": "a.pdf"},
         {"title": "文B", "topic": "主题B", "source": "b.pdf"}]


def test_build_outline_valid(monkeypatch):
    def fake_chat_json(system, user):
        return {"chapters": [
            {"title": "绪论", "kind": "intro", "cards": []},
            {"title": "相关技术综述", "kind": "review", "cards": [0, 1]},
            {"title": "系统设计", "kind": "core", "cards": [0]},
            {"title": "总结", "kind": "conclusion", "cards": []},
        ]}
    monkeypatch.setattr(synthesizer, "_chat_json", fake_chat_json)
    outline = synthesizer.build_outline(TOPIC, CARDS, [])
    assert [c["kind"] for c in outline] == ["intro", "review", "core",
                                            "conclusion"]
    assert outline[1]["cards"] == [0, 1]


def test_build_outline_drops_bad_kind_and_card_ids(monkeypatch):
    def fake_chat_json(system, user):
        return {"chapters": [
            {"title": "绪论", "kind": "intro", "cards": []},
            {"title": "怪章", "kind": "weird", "cards": []},      # 非法 kind 丢弃
            {"title": "综述", "kind": "review", "cards": [0, 9]},  # 越界编号过滤
            {"title": "设计", "kind": "core", "cards": []},
            {"title": "总结", "kind": "conclusion", "cards": []},
        ]}
    monkeypatch.setattr(synthesizer, "_chat_json", fake_chat_json)
    outline = synthesizer.build_outline(TOPIC, CARDS, [])
    assert all(c["kind"] in synthesizer._VALID_KINDS for c in outline)
    assert outline[1]["cards"] == [0]


def test_build_outline_llm_failure_uses_default_skeleton(monkeypatch, capsys):
    def boom(system, user):
        raise RuntimeError("接口挂了")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    outline = synthesizer.build_outline(TOPIC, CARDS, [])
    assert [c["title"] for c in outline] == \
        [c["title"] for c in REFS_SPEC["default_outline"]]
    # 默认骨架下所有卡都关联到各综述章，素材不至于丢
    review = [c for c in outline if c["kind"] == "review"]
    assert all(c["cards"] == [0, 1] for c in review)
    assert "大纲" in capsys.readouterr().out


def test_build_outline_too_few_chapters_falls_back(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "chapters": [{"title": "只有一章", "kind": "core", "cards": []}]})
    outline = synthesizer.build_outline(TOPIC, CARDS, [])
    assert [c["title"] for c in outline] == \
        [c["title"] for c in REFS_SPEC["default_outline"]]
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_synth_outline.py -v`
Expected: FAIL，`has no attribute 'build_outline'`

- [ ] **Step 3: 追加实现**

```python
# ---------------------------------------------------------------------------
#  ② 论文大纲
# ---------------------------------------------------------------------------
_OUTLINE_SYS = (
    "你是论文结构顾问。根据论文题目/研究方向与文献摘要卡，为一篇本科毕业论文"
    "设计章节大纲（4~10章）。每章标注 kind："
    "intro（绪论）、review（文献综述类，可由AI基于文献撰写）、"
    "core（作者必须自己完成的研究/设计/实现/实验章）、conclusion（总结展望）。"
    "cards 列出与该章相关的文献编号（从0开始）。")


def _default_outline(n_cards: int) -> list:
    """降级骨架：全部卡关联到每个综述章，保证素材不丢。"""
    out = []
    for spec_ch in REFS_SPEC["default_outline"]:
        ch = dict(spec_ch)
        ch["cards"] = list(range(n_cards)) if ch["kind"] == "review" else []
        out.append(ch)
    return out


def build_outline(topic: dict, cards: list, img_notes: list) -> list:
    """返回 [{"title", "kind", "cards": [卡编号]}]；失败/不合格退默认骨架。"""
    card_lines = [f"[{i}] {c['title']}：{c.get('topic', '')}"
                  for i, c in enumerate(cards)]
    img_lines = [f"（图片素材）{n['caption']}：{n['summary']}"
                 for n in img_notes if n.get("summary")]
    try:
        data = _chat_json(
            _OUTLINE_SYS,
            '请以 JSON 输出 {"chapters": [{"title": "...", '
            '"kind": "intro|review|core|conclusion", "cards": [0]}]}。\n'
            f"论文题目与研究方向：\n{topic['background'] or topic['title']}\n\n"
            "文献摘要卡：\n" + "\n".join(card_lines + img_lines))
        raw = data.get("chapters") if isinstance(data, dict) else None
        if not isinstance(raw, list):
            raise ValueError("大纲结果缺少 chapters 列表")
        outline = []
        for ch in raw:
            if not isinstance(ch, dict):
                continue
            title = str(ch.get("title") or "").strip()
            kind = str(ch.get("kind") or "").strip()
            if not title or kind not in _VALID_KINDS:
                continue
            ids = sorted({i for i in (ch.get("cards") or [])
                          if isinstance(i, int) and 0 <= i < len(cards)})
            outline.append({"title": title, "kind": kind, "cards": ids})
        if not 4 <= len(outline) <= 10:
            raise ValueError(f"大纲章数不合理：{len(outline)}")
        return outline
    except Exception as e:  # noqa: BLE001
        _note_degrade("大纲生成", e)
        return _default_outline(len(cards))
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_synth_outline.py -v`
Expected: 4 PASS

- [ ] **Step 5: 提交**

```bash
python -m pytest tests/ -q
git add src/synthesizer.py tests/test_synth_outline.py
git commit -m "feat(synthesizer): LLM outline with default-skeleton fallback"
```

### Task 10: write_review（综述章正文，带 [n] 引用）

**Files:**
- Modify: `thesis_project/src/synthesizer.py`（追加）
- Test: `thesis_project/tests/test_synth_review.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import synthesizer
from src.llm_enhancer import AI_MARK

TOPIC = {"title": "题目X", "author": "张三", "background": "背景"}
CARDS = [{"title": "文A", "topic": "tA", "method": "mA", "conclusion": "cA",
          "quotes": ["观点A"], "source": "a.pdf", "authors": [], "year": ""},
         {"title": "文B", "topic": "tB", "method": "mB", "conclusion": "cB",
          "quotes": [], "source": "b.pdf", "authors": [], "year": ""}]
CH = {"title": "相关技术综述", "kind": "review", "cards": [0, 1]}


def test_write_review_paras_carry_ai_mark(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "paras": ["目标检测方法分为两类[1]。", "近年趋势是端到端[2]。"]})
    paras = synthesizer.write_review(CH, TOPIC, CARDS)
    assert len(paras) == 2
    assert all(p.endswith(AI_MARK) for p in paras)
    assert "[1]" in paras[0]


def test_write_review_citation_ids_are_one_based_positions(monkeypatch):
    captured = {}

    def fake_chat_json(system, user):
        captured["user"] = user
        return {"paras": ["p"]}
    monkeypatch.setattr(synthesizer, "_chat_json", fake_chat_json)
    synthesizer.write_review({"title": "综述", "kind": "review", "cards": [1]},
                             TOPIC, CARDS)
    # 只喂关联卡；提示词中的引用编号是全局 1-based（文B -> [2]）
    assert "[2] 文B" in captured["user"]
    assert "文A" not in captured["user"]


def test_write_review_failure_returns_material_fallback(monkeypatch, capsys):
    def boom(system, user):
        raise RuntimeError("挂了")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    paras = synthesizer.write_review(CH, TOPIC, CARDS)
    # 降级：素材摘录（卡片要点罗列）+ 占位符，无 AI 正文
    assert paras[-1] == synthesizer.PLACEHOLDER
    assert any("文A" in p for p in paras)
    assert "综述" in capsys.readouterr().out
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_synth_review.py -v`
Expected: FAIL，`has no attribute 'write_review'`

- [ ] **Step 3: 追加实现**

```python
# ---------------------------------------------------------------------------
#  ③ 综述章撰写
# ---------------------------------------------------------------------------
_REVIEW_SYS = (
    "你是学术写作助手，为本科毕业论文撰写文献综述章节的初稿段落。"
    "只能综合给定摘要卡的信息，禁止编造数据与文献；"
    "引用某张卡片的内容时，在句末用其方括号编号标注（如[1]）。"
    "输出 3~6 个中文段落，每段 100~200 字。")


def _card_brief(idx: int, card: dict) -> str:
    """卡片 -> 提示词行；idx 为全局 1-based 引用编号。"""
    bits = [f"[{idx}] {card['title']}"]
    for k, label in (("topic", "主题"), ("method", "方法"),
                     ("conclusion", "结论")):
        if card.get(k):
            bits.append(f"{label}：{card[k]}")
    if card.get("quotes"):
        bits.append("观点：" + "；".join(card["quotes"][:3]))
    if card.get("fallback_text"):
        bits.append("原文片段：" + card["fallback_text"][:300])
    return "。".join(bits)


def _material_paras(ch: dict, cards: list) -> list:
    """综述降级产物：素材摘录 + 占位符。"""
    paras = ["素材摘录（LLM 综述失败，以下为原始卡片信息）："]
    for i in ch["cards"]:
        paras.append(_card_brief(i + 1, cards[i]))
    paras.append(PLACEHOLDER)
    return paras


def write_review(ch: dict, topic: dict, cards: list) -> list:
    """综述章 -> 正文段落列表（每段带 AI_MARK）；失败退素材摘录。"""
    briefs = [_card_brief(i + 1, cards[i]) for i in ch["cards"]]
    try:
        data = _chat_json(
            _REVIEW_SYS,
            '请以 JSON 输出 {"paras": ["段落1", "段落2"]}。\n'
            f"论文题目：{topic['title']}\n章节标题:{ch['title']}\n\n"
            "文献摘要卡：\n" + "\n".join(briefs))
        paras = [str(p).strip() for p in (data.get("paras") or [])
                 if str(p).strip()] if isinstance(data, dict) else []
        if not paras:
            raise ValueError("综述结果没有段落")
        return [f"{p} {AI_MARK}" for p in paras]
    except Exception as e:  # noqa: BLE001
        _note_degrade(f"《{ch['title']}》综述", e)
        return _material_paras(ch, cards)
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_synth_review.py -v`
Expected: 3 PASS

- [ ] **Step 5: 提交**

```bash
python -m pytest tests/ -q
git add src/synthesizer.py tests/test_synth_review.py
git commit -m "feat(synthesizer): AI-drafted review chapters with [n] citations"
```

### Task 11: write_points（核心章写作要点，批量）

**Files:**
- Modify: `thesis_project/src/synthesizer.py`（追加）
- Test: `thesis_project/tests/test_synth_points.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import synthesizer

TOPIC = {"title": "题目X", "author": "张三", "background": "背景"}
CARDS = [{"title": "文A", "topic": "tA", "source": "a.pdf"}]
CHS = [{"title": "系统设计", "kind": "core", "cards": [0]},
       {"title": "总结与展望", "kind": "conclusion", "cards": []}]


def test_write_points_batch_maps_titles(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "系统设计": ["先画总体架构图", "说明模块划分依据，可参考[1]"],
        "总结与展望": ["概括三项工作"]})
    points = synthesizer.write_points(CHS, TOPIC, CARDS)
    assert points["系统设计"][0] == "先画总体架构图"
    assert len(points["总结与展望"]) == 1


def test_write_points_failure_returns_empty(monkeypatch, capsys):
    def boom(system, user):
        raise RuntimeError("挂了")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    assert synthesizer.write_points(CHS, TOPIC, CARDS) == {}
    assert "写作要点" in capsys.readouterr().out


def test_write_points_skips_when_no_chapters():
    # 不该发起任何调用（无 monkeypatch 也不能炸）
    assert synthesizer.write_points([], TOPIC, CARDS) == {}
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_synth_points.py -v`
Expected: FAIL，`has no attribute 'write_points'`

- [ ] **Step 3: 追加实现**

```python
# ---------------------------------------------------------------------------
#  ④ 核心章写作要点（一次批量调用）
# ---------------------------------------------------------------------------
_POINTS_SYS = (
    "你是论文写作教练。为每个章节生成写作要点：作者应当写什么内容、"
    "按什么顺序展开、可参考哪些文献编号（如[1]）。"
    "每章 3~6 条，每条不超过 60 字。只给指引，不代写正文。"
    '只输出 JSON：{"章节标题": ["要点", ...], ...}')


def write_points(chapters: list, topic: dict, cards: list) -> dict:
    """非 review 章 -> {章节标题: [要点...]}；失败返回 {}（调用方留占位）。"""
    if not chapters:
        return {}
    briefs = [_card_brief(i + 1, c) for i, c in enumerate(cards)]
    payload = {"论文题目": topic["title"],
               "章节": [ch["title"] for ch in chapters],
               "文献摘要卡": briefs}
    try:
        data = _chat_json(_POINTS_SYS, json.dumps(payload, ensure_ascii=False))
        if not isinstance(data, dict):
            raise ValueError("写作要点结果不是 JSON 对象")
        out = {}
        for ch in chapters:
            v = data.get(ch["title"])
            if isinstance(v, list):
                pts = [str(x).strip()[:60] for x in v if str(x).strip()]
                if pts:
                    out[ch["title"]] = pts[:6]
        return out
    except Exception as e:  # noqa: BLE001
        _note_degrade("写作要点", e)
        return {}
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_synth_points.py -v`
Expected: 3 PASS

- [ ] **Step 5: 提交**

```bash
python -m pytest tests/ -q
git add src/synthesizer.py tests/test_synth_points.py
git commit -m "feat(synthesizer): batched writing-point generation for core chapters"
```

### Task 12: format_references（GB/T 7714，失败罗列标题）

**Files:**
- Modify: `thesis_project/src/synthesizer.py`（追加）
- Test: `thesis_project/tests/test_synth_refs_table.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import synthesizer

CARDS = [{"title": "文A", "authors": ["李四"], "year": "2023", "source": "a.pdf"},
         {"title": "文B", "authors": [], "year": "", "source": "b.pdf"}]


def test_format_references_order_preserved(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "references": ["李四. 文A[J]. 某刊, 2023.",
                       "佚名. 文B[EB/OL]. «请补全»."]})
    refs = synthesizer.format_references(CARDS)
    assert len(refs) == 2
    assert refs[0].startswith("李四")


def test_format_references_count_mismatch_falls_back(monkeypatch, capsys):
    monkeypatch.setattr(synthesizer, "_chat_json", lambda s, u: {
        "references": ["只有一条"]})     # 数量对不上 -> 不可信，整体降级
    refs = synthesizer.format_references(CARDS)
    assert refs == ["文A. «请补全著录信息»（来源文件：a.pdf）",
                    "文B. «请补全著录信息»（来源文件：b.pdf）"]
    assert "参考文献" in capsys.readouterr().out


def test_format_references_llm_failure_falls_back(monkeypatch):
    def boom(system, user):
        raise RuntimeError("挂了")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    refs = synthesizer.format_references(CARDS)
    assert len(refs) == 2 and "a.pdf" in refs[0]


def test_format_references_empty_cards():
    assert synthesizer.format_references([]) == []
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_synth_refs_table.py -v`
Expected: FAIL，`has no attribute 'format_references'`

- [ ] **Step 3: 追加实现**

```python
# ---------------------------------------------------------------------------
#  ⑤ 参考文献表（GB/T 7714）
# ---------------------------------------------------------------------------
_GBT_SYS = (
    "你是参考文献格式化助手。把给定文献元数据逐条格式化为 GB/T 7714 著录条目，"
    "严格按输入顺序输出、数量一致。信息缺失处用«请补全»标注，"
    "禁止编造卷期页码等信息。")


def _raw_references(cards: list) -> list:
    return [f"{c['title']}. «请补全著录信息»（来源文件：{c['source']}）"
            for c in cards]


def format_references(cards: list) -> list:
    """摘要卡 -> GB/T 7714 条目列表（顺序即 [n] 引用编号）；失败罗列标题。"""
    if not cards:
        return []
    payload = [{"title": c["title"], "authors": c.get("authors", []),
                "year": c.get("year", ""), "source": c["source"]}
               for c in cards]
    try:
        data = _chat_json(
            _GBT_SYS,
            '请以 JSON 输出 {"references": ["条目1", "条目2"]}：\n'
            + json.dumps(payload, ensure_ascii=False))
        refs = [str(r).strip() for r in (data.get("references") or [])
                if str(r).strip()] if isinstance(data, dict) else []
        if len(refs) != len(cards):
            raise ValueError(f"条目数不符：{len(refs)} != {len(cards)}")
        return refs
    except Exception as e:  # noqa: BLE001
        _note_degrade("参考文献格式化", e)
        return _raw_references(cards)
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_synth_refs_table.py -v`
Expected: 4 PASS

- [ ] **Step 5: 提交**

```bash
python -m pytest tests/ -q
git add src/synthesizer.py tests/test_synth_refs_table.py
git commit -m "feat(synthesizer): GB/T 7714 reference list with raw fallback"
```

### Task 13: describe_images + attach_media（视觉理解与媒体挂载）

**Files:**
- Modify: `thesis_project/src/synthesizer.py`（追加）
- Test: `thesis_project/tests/test_synth_media.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import synthesizer, llm_vision
from config.format_spec import REFS_SPEC


def _img_doc(name):
    return {"source": f"input/{name}", "type": "image", "meta": {},
            "blocks": [{"kind": "image", "level": 0, "text": "", "rows": None,
                        "data": b"png-bytes", "ext": ".png"}]}


def _xlsx_doc(name):
    return {"source": f"input/{name}", "type": "xlsx", "meta": {},
            "blocks": [{"kind": "table", "level": 0, "text": "",
                        "rows": [["a", "1"]]}]}


def _ch(title, kind="core"):
    n = synthesizer._node(title, 1)
    n["kind"] = kind
    return n


def test_describe_images_skipped_when_unavailable(monkeypatch):
    monkeypatch.delenv("LLM_VISION_MODEL", raising=False)
    assert synthesizer.describe_images([_img_doc("架构图.png")]) == []


def test_describe_images_collects_notes(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("LLM_VISION_MODEL", "qwen-vl-plus")
    monkeypatch.setattr(llm_vision, "_chat_vision",
                        lambda p, b, m: '{"caption": "架构图", "summary": "三层"}')
    notes = synthesizer.describe_images([_img_doc("架构图.png")])
    assert notes == [{"source": "架构图.png", "caption": "架构图",
                      "summary": "三层"}]


def test_describe_images_failure_degrades(monkeypatch, capsys):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.setenv("LLM_VISION_MODEL", "qwen-vl-plus")

    def boom(p, b, m):
        raise RuntimeError("端点不支持")
    monkeypatch.setattr(llm_vision, "_chat_vision", boom)
    assert synthesizer.describe_images([_img_doc("图.png")]) == []
    assert "视觉理解" in capsys.readouterr().out


def test_attach_media_matches_by_filename_keyword():
    chapters = [_ch("系统架构设计"), _ch("总结", "conclusion")]
    synthesizer.attach_media(chapters, [_img_doc("架构.png")], [])
    assert len(chapters[0]["images"]) == 1        # 文件名"架构"命中章标题
    assert len(chapters) == 2                     # 命中即不建素材附录章


def test_attach_media_unmatched_goes_to_materials_chapter():
    chapters = [_ch("系统设计")]
    synthesizer.attach_media(chapters, [_xlsx_doc("问卷统计.xlsx")], [])
    assert chapters[-1]["title"] == REFS_SPEC["materials_chapter"]
    assert chapters[-1]["tables"] == [[["a", "1"]]]


def test_attach_media_vision_note_becomes_para():
    chapters = [_ch("系统架构设计")]
    notes = [{"source": "架构.png", "caption": "总体架构",
              "summary": "三层结构"}]
    synthesizer.attach_media(chapters, [_img_doc("架构.png")], notes)
    joined = "\n".join(chapters[0]["paras"])
    assert "总体架构" in joined and "三层结构" in joined


def test_attach_media_no_media_no_extra_chapter():
    chapters = [_ch("系统设计")]
    synthesizer.attach_media(chapters, [], [])
    assert len(chapters) == 1
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_synth_media.py -v`
Expected: FAIL，`has no attribute 'describe_images'`

- [ ] **Step 3: 追加实现**

```python
# ---------------------------------------------------------------------------
#  ⑥ 视觉理解 + 媒体挂载（纯规则挂载）
# ---------------------------------------------------------------------------
def describe_images(media_docs: list) -> list:
    """截图 -> [{"source", "caption", "summary"}]；未配置视觉模型返回 []。"""
    from src import llm_vision
    if not llm_vision.is_vision_available():
        return []
    notes = []
    for d in media_docs:
        if d["type"] != "image":
            continue
        name = os.path.basename(d["source"])
        b = d["blocks"][0]
        try:
            r = llm_vision.describe_image(b["data"], b.get("ext", ".png"))
            notes.append({"source": name, "caption": r["caption"],
                          "summary": r["summary"]})
        except Exception as e:  # noqa: BLE001
            _note_degrade(f"《{name}》视觉理解", e)
    return notes


def _match_chapter(chapters: list, filename: str):
    """文件名（去扩展名）与章标题做双向子串匹配；命中最先出现的章。"""
    stem = os.path.splitext(filename)[0]
    for ch in chapters:
        if ch.get("kind") == "review":
            continue          # 综述章是文献内容，不挂作者自己的素材
        if stem and (stem in ch["title"] or ch["title"] in stem or any(
                len(seg) >= 2 and seg in ch["title"]
                for seg in re.split(r"[\s_\-（）()]+", stem))):
            return ch
    return None


def attach_media(chapters: list, media_docs: list, img_notes: list) -> None:
    """xlsx/csv 表格与截图就地挂到语义匹配章；无匹配挂素材附录章。"""
    notes_by_source = {n["source"]: n for n in img_notes}
    materials = None

    def target_for(filename):
        nonlocal materials
        ch = _match_chapter(chapters, filename)
        if ch is not None:
            return ch
        if materials is None:
            materials = _node(REFS_SPEC["materials_chapter"], 1)
            materials["kind"] = "core"
            chapters.append(materials)
        return materials

    for d in media_docs:
        name = os.path.basename(d["source"])
        ch = target_for(name)
        for b in d["blocks"]:
            if b["kind"] == "table":
                ch["tables"].append(b["rows"])
                ch["paras"].append(f"（下表数据来自：{name}，表题请补全）")
            elif b["kind"] == "image":
                ch["images"].append({"data": b.get("data"),
                                     "ext": b.get("ext", ".png")})
                note = notes_by_source.get(name)
                if note:
                    ch["paras"].append(
                        f"（插图来自：{name}；图题建议：{note['caption']}；"
                        f"内容摘要：{note['summary']} {AI_MARK}）")
                else:
                    ch["paras"].append(f"（插图来自：{name}，图题请补全）")
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_synth_media.py -v`
Expected: 7 PASS

- [ ] **Step 5: 提交**

```bash
python -m pytest tests/ -q
git add src/synthesizer.py tests/test_synth_media.py
git commit -m "feat(synthesizer): optional vision notes and rule-based media attach"
```

### Task 14: synthesize() 总入口

**Files:**
- Modify: `thesis_project/src/synthesizer.py`（追加）
- Test: `thesis_project/tests/test_synth_entry.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
# -*- coding: utf-8 -*-
from src import synthesizer
from src.organizer import PLACEHOLDER
from src.llm_enhancer import AI_MARK
from config.format_spec import REFS_SPEC


def _topic_doc():
    return {"source": "input/topic.md", "type": "md",
            "meta": {"author": "张三"},
            "blocks": [{"kind": "heading", "level": 1,
                        "text": "基于X的Y系统", "rows": None}]}


def _ref_doc(name, text):
    return {"source": f"input/{name}", "type": "pdf", "meta": {},
            "blocks": [{"kind": "paragraph", "level": 0, "text": text,
                        "rows": None}]}


def _fake_chat_json_factory():
    """按 system 提示词分发的全流程假 LLM。"""
    def fake(system, user):
        if "文献调研助手" in system:
            return {"title": "文A", "authors": ["李四"], "year": "2023",
                    "topic": "tA", "method": "mA", "conclusion": "cA",
                    "quotes": ["观点A"]}
        if "论文结构顾问" in system:
            return {"chapters": [
                {"title": "绪论", "kind": "intro", "cards": []},
                {"title": "相关技术综述", "kind": "review", "cards": [0]},
                {"title": "系统设计", "kind": "core", "cards": [0]},
                {"title": "总结与展望", "kind": "conclusion", "cards": []}]}
        if "学术写作助手" in system:
            return {"paras": ["综述正文一[1]。", "综述正文二[1]。"]}
        if "写作教练" in system:
            return {"绪论": ["交代背景"], "系统设计": ["先画架构图，参考[1]"],
                    "总结与展望": ["概括工作"]}
        if "参考文献格式化" in system:
            return {"references": ["李四. 文A[J]. 某刊, 2023."]}
        raise AssertionError("未知提示词: " + system[:20])
    return fake


def test_synthesize_full_pipeline(monkeypatch):
    monkeypatch.setattr(synthesizer, "_chat_json", _fake_chat_json_factory())
    monkeypatch.delenv("LLM_VISION_MODEL", raising=False)
    thesis = synthesizer.synthesize(_topic_doc(), [_ref_doc("a.pdf", "正文")])

    # thesis dict 与 organizer 输出同构
    assert set(thesis) >= {"title", "author", "abstract", "abstract_en",
                           "keywords", "keywords_en", "chapters",
                           "auto_skeleton", "references"}
    assert thesis["title"] == "基于X的Y系统"
    assert thesis["author"] == "张三"
    assert thesis["abstract"] == PLACEHOLDER      # 不编造摘要
    assert thesis["auto_skeleton"] is False
    assert thesis["references"] == ["李四. 文A[J]. 某刊, 2023."]

    titles = [c["title"] for c in thesis["chapters"]]
    assert titles == ["绪论", "相关技术综述", "系统设计", "总结与展望"]

    review = thesis["chapters"][1]
    assert REFS_SPEC["review_notice"] in review["paras"][0]   # 醒目提示在首段
    assert review["paras"][1].endswith(AI_MARK)

    core = thesis["chapters"][2]
    assert core["paras"][0] == "【写作要点】"
    assert any("架构图" in p for p in core["paras"])
    assert core["paras"][-1] == PLACEHOLDER       # 核心章正文留给作者
    # 章节节点结构完整（docx_builder 兼容）
    for ch in thesis["chapters"]:
        assert set(ch) >= {"title", "level", "paras", "subs", "tables",
                           "images"}


def test_synthesize_prints_degrade_summary(monkeypatch, capsys):
    def boom(system, user):
        raise RuntimeError("全挂")
    monkeypatch.setattr(synthesizer, "_chat_json", boom)
    monkeypatch.delenv("LLM_VISION_MODEL", raising=False)
    thesis = synthesizer.synthesize(_topic_doc(), [_ref_doc("a.pdf", "正文")])
    # 全部降级仍产出完整骨架
    assert [c["title"] for c in thesis["chapters"]] == \
        [c["title"] for c in REFS_SPEC["default_outline"]]
    assert len(thesis["references"]) == 1
    out = capsys.readouterr().out
    assert "步降级" in out
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_synth_entry.py -v`
Expected: FAIL，`has no attribute 'synthesize'`

- [ ] **Step 3: 追加实现**

```python
# ---------------------------------------------------------------------------
#  总入口
# ---------------------------------------------------------------------------
def synthesize(topic_doc: dict, ref_docs: list) -> dict:
    """参考资料 -> thesis dict（与 organizer.organize 同构，仅 Word 用）。"""
    del _degraded[:]
    topic = parse_topic(topic_doc)
    text_docs = [d for d in ref_docs if d["type"] not in _MEDIA_TYPES]
    media_docs = [d for d in ref_docs if d["type"] in _MEDIA_TYPES]

    print(f"  文献 {len(text_docs)} 篇，数据/截图 {len(media_docs)} 个")
    cards = make_cards(text_docs)
    img_notes = describe_images(media_docs)
    outline = build_outline(topic, cards, img_notes)

    chapters = []
    for spec_ch in outline:
        ch = _node(spec_ch["title"], 1)
        ch["kind"] = spec_ch["kind"]
        if spec_ch["kind"] == "review":
            ch["paras"].append("【提示】" + REFS_SPEC["review_notice"])
            ch["paras"].extend(write_review(spec_ch, topic, cards))
        chapters.append(ch)

    plain = [spec_ch for spec_ch in outline if spec_ch["kind"] != "review"]
    points = write_points(plain, topic, cards)
    by_title = {ch["title"]: ch for ch in chapters}
    for spec_ch in plain:
        ch = by_title[spec_ch["title"]]
        pts = points.get(ch["title"])
        if pts:
            ch["paras"].append("【写作要点】")
            ch["paras"].extend(f"· {p}" for p in pts)
        quotes = [f"[{i + 1}] {q}" for i in spec_ch["cards"]
                  for q in cards[i].get("quotes", [])[:2]]
        if quotes:
            ch["paras"].append("素材摘录：")
            ch["paras"].extend(quotes)
        ch["paras"].append(PLACEHOLDER)

    attach_media(chapters, media_docs, img_notes)
    references = format_references(cards)

    if _degraded:
        print(f"  [提示] 本次共 {len(_degraded)} 步降级，请检查上方告警。")
    return {
        "title": topic["title"],
        "author": topic["author"],
        "abstract": PLACEHOLDER,
        "abstract_en": PLACEHOLDER,
        "keywords": [PLACEHOLDER],
        "keywords_en": [PLACEHOLDER],
        "chapters": chapters,
        "auto_skeleton": False,
        "references": references,
    }
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_synth_entry.py -v`
Expected: 2 PASS

- [ ] **Step 5: 提交**

```bash
python -m pytest tests/ -q
git add src/synthesizer.py tests/test_synth_entry.py
git commit -m "feat(synthesizer): assemble full thesis dict from reference materials"
```

## Phase 4 入口集成与收尾

### Task 15: main.py 模式分流（topic 检测 + --mode + 入口硬校验）

**Files:**
- Modify: `thesis_project/src/main.py`
- Test: `thesis_project/tests/test_main_refs_mode.py`（新建）

- [ ] **Step 1: 写失败测试**

```python
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
```

- [ ] **Step 2: 运行确认失败**

Run: `python -m pytest tests/test_main_refs_mode.py -v`
Expected: FAIL，`has no attribute '_split_topic'`

- [ ] **Step 3: 修改 main.py**

3a. 顶部 import 区（`from src import docx_builder, pptx_builder` 之后）加：

```python
from config.format_spec import REFS_SPEC
```

3b. 在 `_build_with_retry` 之后插入三个函数：

```python
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
    thesis = synthesizer.synthesize(topic_doc, ref_docs)
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
    return 0
```

3c. `main()` 的 argparse 增加（放在 `--only` 之后）：

```python
    ap.add_argument("--mode", choices=["auto", "refs", "draft"],
                    default="auto",
                    help="auto: input/ 有题目文件(topic.md等)时走参考资料模式；"
                         "refs/draft 强制指定")
```

3d. `main()` 中，在 `print("② 整理内容结构")` 之前插入分流：

```python
    topic_doc, ref_docs = _split_topic(docs)
    mode = args.mode
    if mode == "auto":
        mode = "refs" if topic_doc is not None else "draft"
    if mode == "refs":
        return _run_refs_mode(args, topic_doc, ref_docs)
    # draft 模式沿用原流程（docs 不剔除题目文件）
```

- [ ] **Step 4: 运行确认通过**

Run: `python -m pytest tests/test_main_refs_mode.py -v`
Expected: 7 PASS

- [ ] **Step 5: 全量回归 + 提交**

```bash
python -m pytest tests/ -q
git add src/main.py tests/test_main_refs_mode.py
git commit -m "feat(main): auto-detect refs mode via topic file with --mode override"
```

### Task 16: 端到端测试（打桩 LLM 全流程出 docx）

**Files:**
- Test: `thesis_project/tests/test_e2e_refs.py`（新建）

- [ ] **Step 1: 写测试（直接可运行——底层均已实现，本任务是集成验证）**

```python
# -*- coding: utf-8 -*-
"""参考资料模式端到端：真实文件 -> main() -> 校验生成的 docx。"""
import os
import sys

import docx as docx_lib
import pytest

from src import main as main_mod
from src import synthesizer


@pytest.fixture()
def refs_input(tmp_path):
    d = tmp_path / "input"
    d.mkdir()
    (d / "topic.md").write_text(
        "---\nauthor: 张三\n---\n# 基于X的Y系统\n\n研究内容：……\n",
        encoding="utf-8")
    (d / "文献A.md").write_text("# 文A\n\n目标检测综述正文。\n",
                                encoding="utf-8")
    (d / "数据.csv").write_text("组别,精度\nA,0.91\n", encoding="utf-8")
    return d


def _fake_chat_json(system, user):
    if "文献调研助手" in system:
        return {"title": "文A", "authors": ["李四"], "year": "2023",
                "topic": "目标检测", "method": "综述", "conclusion": "有效",
                "quotes": ["检测器分两类"]}
    if "论文结构顾问" in system:
        return {"chapters": [
            {"title": "绪论", "kind": "intro", "cards": []},
            {"title": "相关技术综述", "kind": "review", "cards": [0]},
            {"title": "系统设计", "kind": "core", "cards": [0]},
            {"title": "总结与展望", "kind": "conclusion", "cards": []}]}
    if "学术写作助手" in system:
        return {"paras": ["综述正文[1]。"]}
    if "写作教练" in system:
        return {"系统设计": ["先画架构图"]}
    if "参考文献格式化" in system:
        return {"references": ["李四. 文A[J]. 某刊, 2023."]}
    raise AssertionError("未知提示词")


def test_e2e_refs_mode_generates_word_only(refs_input, tmp_path, monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    monkeypatch.delenv("LLM_VISION_MODEL", raising=False)
    monkeypatch.setattr(synthesizer, "_chat_json", _fake_chat_json)
    out = tmp_path / "output"
    monkeypatch.setattr(sys, "argv",
                        ["main.py", "--input", str(refs_input),
                         "--output", str(out)])
    assert main_mod.main() == 0

    assert (out / "论文草案.docx").exists()
    assert not (out / "答辩PPT草案.pptx").exists()     # refs 模式不产 PPT

    doc = docx_lib.Document(str(out / "论文草案.docx"))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "基于X的Y系统" in text
    assert "相关技术综述" in text
    assert "<AI生成，请核对>" in text
    assert "李四. 文A[J]. 某刊, 2023." in text
    assert len(doc.tables) >= 1                        # csv 表格已插入


def test_e2e_draft_mode_untouched(tmp_path, monkeypatch):
    """无 topic 文件时走原流程，仍生成两份产物（回归保护）。"""
    d = tmp_path / "input"
    d.mkdir()
    (d / "草稿.md").write_text("# 我的论文\n\n## 绪论\n\n正文。\n",
                               encoding="utf-8")
    out = tmp_path / "output"
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    monkeypatch.setattr(sys, "argv",
                        ["main.py", "--input", str(d), "--output", str(out)])
    assert main_mod.main() == 0
    assert (out / "论文草案.docx").exists()
    assert (out / "答辩PPT草案.pptx").exists()
```

- [ ] **Step 2: 运行确认通过**

Run: `python -m pytest tests/test_e2e_refs.py -v`
Expected: 2 PASS（若失败按报错修正前序任务的集成缝隙——这是本任务的目的）

- [ ] **Step 3: 全量回归 + 提交**

```bash
python -m pytest tests/ -q
git add tests/test_e2e_refs.py
git commit -m "test(e2e): refs mode full pipeline with stubbed LLM"
```

### Task 17: run.bat / README / spec 修订 + 真实冒烟

**Files:**
- Modify: `thesis_project/run.bat:27,30,33`（依赖检查加 openpyxl）
- Modify: `thesis_project/README.md`（新增"参考资料模式"章节）
- Modify: `docs/superpowers/specs/2026-07-22-refs-to-draft-design.md`（llm_vision 偏差回写）

- [ ] **Step 1: run.bat 依赖行更新**

第 27 行改为：

```bat
%PY% -c "import docx, pptx, pdfplumber, openai, openpyxl" >nul 2>nul
```

第 30 行与 33 行的 pip install 清单同步加 `openpyxl`：

```bat
    %PY% -m pip install python-docx python-pptx pdfplumber openai openpyxl
```

- [ ] **Step 2: README 更新**

在"LLM 增强（可选）"章节之后新增：

```markdown
---

## 参考资料模式（题目 + 文献 → 论文初稿骨架）

`input\` 里放一个**题目文件**（`topic.md` / `题目.txt`，写论文题目、研究内容
简述、拟采用方法）+ 参考文献（PDF/Word/md/txt/json）、数据（xlsx/csv）、
截图（png/jpg），双击 `run.bat` 即自动进入本模式：

- LLM 生成：文献综述章节（带 [n] 引用）、全文大纲、核心章节【写作要点】
  与素材摘录、GB/T 7714 参考文献表；
- LLM 不生成：研究设计/实现/实验等核心章节正文——留 `<请填写>` 由你完成；
- 只生成 `论文草案.docx`，不生成 PPT（补完正文后用普通模式再生成）；
- xlsx/csv 自动插表、截图自动插图；设置 `LLM_VISION_MODEL`（如 qwen-vl-plus）
  后截图还会获得图题建议与内容摘要；
- **必须**设置 `LLM_API_KEY`（见上节），未设置会报错退出；
- `--mode refs|draft` 可强制指定模式，覆盖自动检测。

> ⚠ 学术诚信：综述正文全部带 `<AI生成，请核对>` 标记，属于初稿素材，
> 务必逐条核对原文献后改写为自己的表述再使用。
```

同时把 README 开头"读取你给出的 **Word / PDF / TXT / Markdown / JSON** 源文件"
一句更新为"**Word / PDF / TXT / Markdown / JSON / Excel(xlsx·csv) / 图片**"。

- [ ] **Step 3: spec 偏差回写**

编辑 `docs/superpowers/specs/2026-07-22-refs-to-draft-design.md` 第 3 节：
把 `llm_client.py` 两处描述改为 `llm_vision.py`（仅承载视觉调用），并注明
"文本调用继续经 `llm_enhancer._chat`（保持既有打桩链）；JSON 解析提为
`_parse_json` 复用"。第 4 节"网络出口"与第 6 节回归说明同步措辞。

- [ ] **Step 4: 真实冒烟（不打桩，需要真实 API Key；无 Key 则只验证报错路径）**

```bash
mkdir -p input_smoke && printf '# 基于X的Y系统\n\n研究内容测试。\n' > input_smoke/topic.md
printf '# 某文献\n\n文献正文。\n' > input_smoke/ref1.md
# 无 Key：应打印 [错误] 参考资料模式需要 LLM 并退出码 1
python src/main.py --input input_smoke --output output_smoke; echo "exit=$?"
# 有 Key 时再跑一次应生成 output_smoke/论文草案.docx
rm -rf input_smoke output_smoke
```

- [ ] **Step 5: 全量回归 + 提交**

```bash
python -m pytest tests/ -q
git add run.bat README.md ../docs/superpowers/specs/2026-07-22-refs-to-draft-design.md
git commit -m "docs: document refs mode; add openpyxl to run.bat deps"
```

---

## 验证清单（全部完成后）

1. `python -m pytest tests/ -q` —— 全量通过（原有 30+ 文件 + 新增 12 个文件）。
2. 冒烟 A（排版模式回归）：`python src/main.py --input sample_input --output output_smoke` 仍产出 docx+pptx。
3. 冒烟 B（refs 模式无 Key）：见 Task 17 Step 4，报错文案正确、退出码 1。
4. 冒烟 C（refs 模式有 Key，真实 LLM）：题目 + 2 篇 md 文献 + 1 个 csv，
   产出的 docx 应包含：综述章（带【提示】与 AI 标记与 [n]）、【写作要点】、
   csv 表格、GB/T 7714 参考文献；且无 PPT。
5. `run.bat` 双击（Windows）：input\ 放 topic.md + 文献时走 refs 模式全程无崩溃。

## 风险与注意

- **测试打桩链**：绝不把 `_chat/_chat_json/_client` 移出 llm_enhancer（见头部偏差说明）。
- **`_node` 复用**：synthesizer 直接 import `organizer._node`，若日后 organizer
  改节点结构，synthesizer 测试会第一时间红。
- **read_only 模式的 openpyxl**：合并单元格非锚点读 None；公式未缓存值读 None——
  均落为空串，文档字符串已注明。
- **`--only` 在 refs 模式下被忽略**（refs 只产 Word），不报错；README 已说明模式差异。





