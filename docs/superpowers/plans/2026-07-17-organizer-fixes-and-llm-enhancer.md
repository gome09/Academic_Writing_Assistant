# 论文草案生成器缺陷修复 + LLM 增强层 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 修复 thesis_project 内容整理管道的 9 处已验证缺陷，并新增可选的 LLM 增强层（OpenAI 兼容接口），提升论文/PPT 草稿质量。

**Architecture:** 保持 `readers → organizer → builders` 三段式不变。readers/builders 仍为纯规则（解析与排版是确定性工作）；organizer 修复语义缺陷并暴露 `classify_fn`/`bullets_fn` 两个注入点；新增 `src/llm_enhancer.py` 在 `organize()` 之后做增强 pass，无 `LLM_API_KEY` 或任一步失败时静默回退到规则结果，主流程永不因 LLM 崩溃。

**Tech Stack:** Python 3.13、python-docx、python-pptx、pdfplumber、openai SDK（≥1.0，OpenAI 兼容端点：DeepSeek/通义/Kimi/Ollama 均可）、pytest 8。

**项目根目录：** `D:/BackendDevelopment/Project/Project_Test-7`（已 git init，基线 commit `957465d`）。
**代码根目录：** `thesis_project/`（下文所有相对路径以此为准；pytest 均在 `thesis_project/` 目录下执行）。
**任务必须按序执行**：Task 3/4/5/6 都修改 `src/organizer.py` 的同一批函数，后面的任务代码假设前面的任务已完成。

**已确认的设计决策（不得擅自更改）：**
1. 范围 = 规则修复 + LLM 增强层（`--llm` 开关，可选）。
2. LLM 用 openai SDK + 可配置 base_url。环境变量：`LLM_API_KEY`（必填才启用）、`LLM_BASE_URL`（选填）、`LLM_MODEL`（选填，默认 `gpt-4o-mini`）。
3. 无标题文档的段落**按原文顺序全部放入"研究内容"一章**，其余骨架章节留 `<请填写>` 占位。
4. LLM 只做整理/提炼/翻译/分类，不扩写正文；AI 生成的摘要类内容追加标记 `<AI生成，请核对>`。
5. 测试中**绝不真实调用 LLM API**，一律 monkeypatch 打桩。

---

## 文件结构总览

| 文件 | 动作 | 职责 |
|---|---|---|
| `tests/conftest.py` | 新建 | 把项目根加入 sys.path |
| `tests/factories.py` | 新建 | 测试用 block/doc 工厂函数 |
| `requirements.txt` | 新建 | 依赖清单 |
| `src/readers.py` | 修改 | 编码回退；PDF 分段/标题识别/表格去重 |
| `src/organizer.py` | 修改 | 题目识别、单章顺序填充、编号剥离、特殊章节、三级标题、表格全量、PPT 分页、注入点 |
| `src/docx_builder.py` | 修改 | 渲染三级标题 |
| `src/pptx_builder.py` | 修改 | 删除页脚死代码 |
| `src/llm_enhancer.py` | 新建 | LLM 客户端 + 五个增强函数 + enhance 总入口 |
| `src/main.py` | 修改 | `--llm` 开关接入 |
| `README.md` | 修改 | LLM 用法文档 |
| `tests/test_*.py` | 新建 | 每个任务一个测试文件（共 15 个） |

---

### Task 0: 测试脚手架与依赖清单

**Files:**
- Create: `thesis_project/tests/__init__.py`（空文件）
- Create: `thesis_project/tests/conftest.py`
- Create: `thesis_project/tests/factories.py`
- Create: `thesis_project/requirements.txt`

- [ ] **Step 1: 创建测试目录与 conftest**

`thesis_project/tests/__init__.py` 为空文件。

`thesis_project/tests/conftest.py`：

```python
# -*- coding: utf-8 -*-
"""把项目根（thesis_project/）加入 sys.path，使 src / config 可导入。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
```

- [ ] **Step 2: 创建测试工厂**

`thesis_project/tests/factories.py`：

```python
# -*- coding: utf-8 -*-
"""构造 readers 中间结构（Block / Document）的测试工厂。"""


def h(level, text):
    """heading 块"""
    return {"kind": "heading", "level": level, "text": text, "rows": None}


def p(text):
    """paragraph 块"""
    return {"kind": "paragraph", "level": 0, "text": text, "rows": None}


def table(rows):
    """table 块"""
    return {"kind": "table", "level": 0, "text": "", "rows": rows}


def doc(blocks, meta=None, type_="md"):
    """Document"""
    return {"source": "test", "type": type_, "blocks": blocks, "meta": meta or {}}
```

- [ ] **Step 3: 创建 requirements.txt**

`thesis_project/requirements.txt`：

```
python-docx
python-pptx
pdfplumber
openai>=1.0
```

- [ ] **Step 4: 验证 pytest 能发现测试目录**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: `no tests ran`（收集成功、无报错即可）

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/tests thesis_project/requirements.txt
git commit -m "chore: add test scaffolding and requirements.txt" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 1: 文本编码回退（GBK 中文 txt 乱码修复）

**缺陷：** `readers.py:45,59` 强制 `utf-8 + errors="ignore"`，中文 Windows 常见的 GBK 编码 txt 会读成乱码且被静默吞掉。

**Files:**
- Modify: `thesis_project/src/readers.py`（`read_txt`、`read_md` 的打开方式）
- Test: `thesis_project/tests/test_readers_encoding.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_readers_encoding.py`：

```python
# -*- coding: utf-8 -*-
from src.readers import read_txt, read_md, _read_text


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
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_readers_encoding.py -v`
Expected: FAIL / ERROR，`cannot import name '_read_text'`

- [ ] **Step 3: 实现**

在 `src/readers.py` 的 `_clean` 函数之后新增：

```python
def _read_text(path: str) -> str:
    """读文本文件：utf-8(-sig) -> gb18030 依次严格尝试，最后 utf-8 宽容兜底。

    UTF-8 文本几乎不可能被 gb18030 抢先命中（utf-8 先试且严格），
    而 GBK/GB18030 文本解码 utf-8 必然报错落到第二档。
    """
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")
```

`read_txt` 中把：

```python
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        raw = f.read()
```

改为：

```python
    raw = _read_text(path)
```

`read_md` 中同样把：

```python
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        raw = f.read()
```

改为：

```python
    raw = _read_text(path)
```

- [ ] **Step 4: 运行确认通过**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_readers_encoding.py -v`
Expected: 4 passed

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/readers.py thesis_project/tests/test_readers_encoding.py
git commit -m "fix(readers): fall back to gb18030 when text is not utf-8" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 2: 论文题目误识别修复

**缺陷：** `organizer.py:59` 把第一个一级标题当论文题目，输入 `# 绪论` 时题目变成"绪论"。

**Files:**
- Modify: `thesis_project/src/organizer.py`（新增 `_GENERIC_HEADING`/`_looks_generic`，修改 `_extract_meta`）
- Test: `thesis_project/tests/test_organizer_meta.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_organizer_meta.py`：

```python
# -*- coding: utf-8 -*-
from src.organizer import organize, PLACEHOLDER
from tests.factories import h, p, doc


def test_generic_chapter_heading_not_title():
    """# 绪论 这类章节名不能被当成论文题目。"""
    docs = [doc([h(1, "绪论"), p("背景内容。")])]
    thesis, _ = organize(docs)
    assert thesis["title"] == PLACEHOLDER


def test_numbered_chapter_heading_not_title():
    docs = [doc([h(1, "第一章 绪论"), p("背景内容。")])]
    thesis, _ = organize(docs)
    assert thesis["title"] == PLACEHOLDER


def test_real_title_heading_recognized():
    docs = [doc([h(1, "基于深度学习的垃圾分类系统设计"), h(1, "绪论"), p("内容。")])]
    thesis, _ = organize(docs)
    assert thesis["title"] == "基于深度学习的垃圾分类系统设计"


def test_meta_title_wins_over_heading():
    docs = [doc([h(1, "基于深度学习的垃圾分类系统设计")],
                meta={"title": "来自frontmatter的题目"})]
    thesis, _ = organize(docs)
    assert thesis["title"] == "来自frontmatter的题目"
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_organizer_meta.py -v`
Expected: 前两个测试 FAIL（题目被误识别为"绪论"/"第一章 绪论"），后两个 PASS

- [ ] **Step 3: 实现**

在 `src/organizer.py` 的 `_SECTION_HINT` 定义之后新增：

```python
# 常见章节名 / 结构性标题 —— 不能当论文题目
_GENERIC_HEADING = re.compile(
    r"^(第\s*[一二三四五六七八九十百\d]+\s*[章节部分]"
    r"|\d+(\.\d+)*[\s、.．]"
    r"|[一二三四五六七八九十]+[、.．]"
    r"|绪论|引言|前言|导论|概述|摘\s*要|abstract|目\s*录|结论|总结|展望"
    r"|参考文献|致\s*谢|附\s*录|关键词"
    r"|相关(理论|技术|工作)|文献综述|研究(背景|现状|意义|内容|方法)"
    r"|国内外研究现状|需求分析|系统设计|系统实现|实验|测试)",
    re.IGNORECASE,
)


def _looks_generic(title: str) -> bool:
    return bool(_GENERIC_HEADING.match(title.strip()))
```

在 `_extract_meta` 中把：

```python
        if not title and b["kind"] == "heading" and b["level"] <= 1 and t:
            title = t
```

改为：

```python
        if (not title and b["kind"] == "heading" and b["level"] <= 1
                and t and not _looks_generic(t)):
            title = t
```

- [ ] **Step 4: 运行确认通过**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_organizer_meta.py -v`
Expected: 4 passed

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/organizer.py thesis_project/tests/test_organizer_meta.py
git commit -m "fix(organizer): do not mistake generic chapter headings for thesis title" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 3: 无标题文档段落顺序保持（单章方案）

**缺陷：** `organizer.py:130-133` 把段落按 `i % 2` 交替塞进"绪论"和"系统实现"，原文顺序被打乱。
**决策：** 全部段落按原序放入新增的"研究内容"一章（插在骨架第 3 位之后），其余骨架章节填 `<请填写>` 占位；同时给 thesis 打 `auto_skeleton` 标记供 LLM 层识别。

**Files:**
- Modify: `thesis_project/src/organizer.py`（`_build_chapters` 返回值改为元组、else 分支重写、`organize` 适配）
- Test: `thesis_project/tests/test_organizer_skeleton.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_organizer_skeleton.py`：

```python
# -*- coding: utf-8 -*-
from src.organizer import organize, PLACEHOLDER, DEFAULT_CHAPTERS
from tests.factories import h, p, doc


def _headingless_docs(n=6):
    return [doc([p(f"第{i}段。") for i in range(1, n + 1)], type_="txt")]


def test_paragraph_order_preserved():
    thesis, _ = organize(_headingless_docs())
    content = [c for c in thesis["chapters"] if c["title"] == "研究内容"]
    assert len(content) == 1
    assert content[0]["paras"] == [f"第{i}段。" for i in range(1, 7)]


def test_skeleton_chapters_have_placeholder():
    thesis, _ = organize(_headingless_docs())
    for ch in thesis["chapters"]:
        if ch["title"] in DEFAULT_CHAPTERS:
            assert ch["paras"] == [PLACEHOLDER]


def test_auto_skeleton_flag_true_for_headingless():
    thesis, _ = organize(_headingless_docs())
    assert thesis["auto_skeleton"] is True


def test_auto_skeleton_flag_false_with_headings():
    thesis, _ = organize([doc([h(1, "绪论"), p("内容。")])])
    assert thesis["auto_skeleton"] is False
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_organizer_skeleton.py -v`
Expected: 4 FAIL（无"研究内容"章、无 auto_skeleton 键）

- [ ] **Step 3: 实现**

`src/organizer.py` 中 `_build_chapters` 的 else 分支（原第 123-133 行）：

```python
    else:
        # 无标题：套骨架，段落顺序填入"研究内容"章节
        paras = [b["text"] for b in blocks if b["text"]]
        for idx, name in enumerate(DEFAULT_CHAPTERS):
            ch = {"title": name, "level": 1, "paras": [], "subs": []}
            chapters.append(ch)
        # 把已有段落均匀塞进中间几章（绪论/实现），避免全空
        target_indices = [0, 3]  # 绪论、系统实现
        for i, p in enumerate(paras):
            ch = chapters[target_indices[i % len(target_indices)]]
            ch["paras"].append(p)

    return chapters
```

整体替换为（注意函数返回值变为元组）：

```python
    else:
        # 无标题：套标准骨架并留占位符；全部段落按原文顺序放入
        # "研究内容"一章，保持叙述连贯（LLM 增强层可再做语义分章）。
        for name in DEFAULT_CHAPTERS:
            chapters.append({"title": name, "level": 1,
                             "paras": [PLACEHOLDER], "subs": []})
        paras = [b["text"] for b in blocks if b["text"]]
        chapters.insert(3, {"title": "研究内容", "level": 1,
                            "paras": paras, "subs": []})

    return chapters, not has_heading
```

同时 `has_heading` 分支末尾的 `return chapters` 也随之变为 `return chapters, not has_heading`（即函数只保留最后这一个 return）。

`organize()` 中把：

```python
    chapters = _build_chapters(docs)
```

改为：

```python
    chapters, auto_skeleton = _build_chapters(docs)
```

并在 thesis 字典中 `"chapters": chapters,` 一行之后新增：

```python
        "auto_skeleton": auto_skeleton,
```

- [ ] **Step 4: 运行确认通过（含回归）**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: 全部 passed（此前任务的测试不回归）

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/organizer.py thesis_project/tests/test_organizer_skeleton.py
git commit -m "fix(organizer): keep paragraph order for headingless docs in single chapter" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 4: 章节标题重复编号剥离

**缺陷：** `docx_builder.py:239` 无条件加 `第{N}章　` 前缀，源标题"第一章 绪论"会渲染成"第一章　第一章 绪论"。
**方案：** 在 organizer 建树时剥离标题自带的编号前缀（`第X章`、`一、`、`1.1` 等），builder 不动。

**Files:**
- Modify: `thesis_project/src/organizer.py`（新增 `_NUM_PREFIX`/`_strip_numbering`，`_build_chapters` 建章/节时调用）
- Test: `thesis_project/tests/test_organizer_numbering.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_organizer_numbering.py`：

```python
# -*- coding: utf-8 -*-
import pytest
from src.organizer import organize, _strip_numbering
from tests.factories import h, p, doc


@pytest.mark.parametrize("raw,expect", [
    ("第一章 绪论", "绪论"),
    ("第1章　绪论", "绪论"),
    ("一、绪论", "绪论"),
    ("1.1 研究背景", "研究背景"),
    ("1.1卷积网络", "卷积网络"),
    ("3 系统实现", "系统实现"),
    ("绪论", "绪论"),               # 无前缀不变
    ("2023年发展综述", "2023年发展综述"),  # 年份不是编号
    ("第一章", "第一章"),            # 剥完为空则保留原文
])
def test_strip_numbering(raw, expect):
    assert _strip_numbering(raw) == expect


def test_chapter_and_sub_titles_stripped():
    docs = [doc([h(1, "第一章 绪论"), h(2, "1.1 研究背景"), p("内容。")])]
    thesis, _ = organize(docs)
    ch = thesis["chapters"][0]
    assert ch["title"] == "绪论"
    assert ch["subs"][0]["title"] == "研究背景"
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_organizer_numbering.py -v`
Expected: ERROR，`cannot import name '_strip_numbering'`

- [ ] **Step 3: 实现**

在 `src/organizer.py` 的 `_looks_generic` 之后新增：

```python
# 标题自带的编号前缀（生成时会统一重新编号，避免"第一章 第一章 绪论"）
_NUM_PREFIX = re.compile(
    r"^(第\s*[一二三四五六七八九十百\d]+\s*[章节]\s*[、.．:：]?"
    r"|[一二三四五六七八九十]+\s*[、.．]"
    r"|\d+(\.\d+)*\s*[、.．]?\s+"
    r"|\d+(\.\d+)+\s*)"
)


def _strip_numbering(title: str) -> str:
    stripped = _NUM_PREFIX.sub("", title).strip()
    return stripped or title.strip()
```

`_build_chapters` 中三处标题赋值改为剥离后的文本：

```python
            if b["kind"] == "heading" and b["level"] <= 1:
                current_ch = {"title": _strip_numbering(b["text"]), "level": 1,
                              "paras": [], "subs": []}
```

```python
                current_sub = {"title": _strip_numbering(b["text"]), "level": 2,
                               "paras": []}
```

（level>=3 分支此时仍是并入段落文本，Task 6 再改。）

- [ ] **Step 4: 运行确认通过（含回归）**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: 全部 passed

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/organizer.py thesis_project/tests/test_organizer_numbering.py
git commit -m "fix(organizer): strip pre-existing numbering from chapter titles" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 5: 特殊章节处理（参考文献抽取，摘要/目录/致谢/附录剔除）

**缺陷：** 源文件里的 `# 参考文献` 变成"第X章 参考文献"且与内置参考文献节重复；`# 摘要` 抽进 meta 后又留在章节树。

**Files:**
- Modify: `thesis_project/src/organizer.py`（新增 `_split_special_chapters`，`organize` 接入）
- Test: `thesis_project/tests/test_organizer_special.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_organizer_special.py`：

```python
# -*- coding: utf-8 -*-
from src.organizer import organize
from tests.factories import h, p, doc


def test_reference_chapter_extracted_to_references():
    docs = [doc([h(1, "总结"), p("总结内容。"),
                 h(1, "参考文献"),
                 p("[1] 张三. 某研究[J]. 某期刊, 2024."),
                 p("[2] 李四. 某系统[D]. 某大学, 2023.")])]
    thesis, _ = organize(docs)
    titles = [c["title"] for c in thesis["chapters"]]
    assert "参考文献" not in titles
    assert thesis["references"] == [
        "张三. 某研究[J]. 某期刊, 2024.",
        "李四. 某系统[D]. 某大学, 2023.",
    ]


def test_abstract_and_thanks_chapters_dropped():
    docs = [doc([h(1, "摘要"), p("摘要正文内容，足够长的一段。"),
                 h(1, "绪论"), p("绪论内容。"),
                 h(1, "致谢"), p("感谢导师。")])]
    thesis, _ = organize(docs)
    titles = [c["title"] for c in thesis["chapters"]]
    assert titles == ["绪论"]
    # 摘要仍被 meta 抽取，没有丢
    assert thesis["abstract"] == "摘要正文内容，足够长的一段。"


def test_default_reference_kept_when_no_ref_chapter():
    docs = [doc([h(1, "绪论"), p("内容。")])]
    thesis, _ = organize(docs)
    assert len(thesis["references"]) == 1
    assert "GB/T 7714" in thesis["references"][0]
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_organizer_special.py -v`
Expected: 前两个 FAIL，第三个 PASS

- [ ] **Step 3: 实现**

在 `src/organizer.py` 的 `_strip_numbering` 之后新增：

```python
_REF_TITLE = re.compile(r"^(参\s*考\s*文\s*献|references?)\s*$", re.IGNORECASE)
_DROP_TITLE = re.compile(r"^(摘\s*要|abstract|目\s*录|致\s*谢|谢\s*辞|附\s*录)",
                         re.IGNORECASE)


def _split_special_chapters(chapters):
    """参考文献章 -> 抽成条目列表；摘要/目录/致谢/附录章 -> 移出正文。

    摘要正文已由 _extract_meta 抽取，这里剔除只是防止正文重复成章。
    """
    body, references = [], []
    for ch in chapters:
        t = ch["title"].strip()
        if _REF_TITLE.match(t):
            paras = list(ch["paras"])
            for sub in ch.get("subs", []):
                paras.extend(sub["paras"])
            for para in paras:
                for line in para.splitlines():
                    line = re.sub(r"^\[?\d+\]?[\.、]?\s*", "", line.strip())
                    if line:
                        references.append(line)
        elif _DROP_TITLE.match(t):
            continue
        else:
            body.append(ch)
    return body, references
```

`organize()` 中把：

```python
    chapters, auto_skeleton = _build_chapters(docs)
```

之后插入一行，并把 references 改用抽取结果：

```python
    chapters, auto_skeleton = _build_chapters(docs)
    chapters, references = _split_special_chapters(chapters)
```

thesis 字典中把：

```python
        "references": [
            "示例. GB/T 7714 著录格式. 出版地: 出版者, 年份.  <请替换为真实文献>",
        ],
```

改为：

```python
        "references": references or [
            "示例. GB/T 7714 著录格式. 出版地: 出版者, 年份.  <请替换为真实文献>",
        ],
```

- [ ] **Step 4: 运行确认通过（含回归）**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: 全部 passed

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/organizer.py thesis_project/tests/test_organizer_special.py
git commit -m "fix(organizer): extract reference chapter, drop abstract/toc/thanks chapters" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 6: 三级标题支持

**缺陷：** `organizer.py:105-108` 把 level>=3 的标题降级为普通段落文本，层级信息丢失；docx 从不使用 Heading 3 样式（规范里已定义）。

**Files:**
- Modify: `thesis_project/src/organizer.py`（`_build_chapters` 支持三级、`_build_deck` 汇集三级段落）
- Modify: `thesis_project/src/docx_builder.py`（渲染 Heading 3）
- Test: `thesis_project/tests/test_organizer_level3.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_organizer_level3.py`：

```python
# -*- coding: utf-8 -*-
from src.organizer import organize
from tests.factories import h, p, doc


def _docs():
    return [doc([h(1, "系统设计"),
                 h(2, "总体架构"),
                 h(3, "前端模块"), p("前端负责采集。"),
                 h(3, "后端模块"), p("后端负责推理。")])]


def test_level3_headings_build_tree():
    thesis, _ = organize(_docs())
    sub = thesis["chapters"][0]["subs"][0]
    assert sub["title"] == "总体架构"
    assert [s3["title"] for s3 in sub["subs"]] == ["前端模块", "后端模块"]
    assert sub["subs"][0]["paras"] == ["前端负责采集。"]
    assert sub["subs"][1]["paras"] == ["后端负责推理。"]


def test_level3_without_parent_sub_promoted():
    """没有二级小节时，三级标题提升为二级。"""
    docs = [doc([h(1, "系统设计"), h(3, "某模块"), p("内容。")])]
    thesis, _ = organize(docs)
    sub = thesis["chapters"][0]["subs"][0]
    assert sub["title"] == "某模块"
    assert sub["paras"] == ["内容。"]


def test_level3_paras_reach_ppt_bullets():
    _, deck = organize(_docs())
    all_bullets = [b for s in deck["slides"] if s["type"] == "content"
                   for b in s["bullets"]]
    assert any("前端负责采集" in b for b in all_bullets)


def test_docx_renders_level3(tmp_path):
    import docx as docx_lib
    from src import docx_builder
    thesis, _ = organize(_docs())
    out = docx_builder.build(thesis, str(tmp_path / "t.docx"))
    d = docx_lib.Document(out)
    h3 = [q.text for q in d.paragraphs if q.style.name == "Heading 3"]
    assert any("前端模块" in t for t in h3)
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_organizer_level3.py -v`
Expected: 4 FAIL（sub 没有 "subs" 键 / KeyError / Heading 3 不存在）

- [ ] **Step 3: 实现 organizer**

`src/organizer.py` 的 `_build_chapters` 中 `has_heading` 分支整体替换为如下（这是该分支在 Task 3/4/5 之后的完整最终形态）：

```python
    if has_heading:
        current_ch = None
        current_sub = None
        current_sub3 = None
        for b in blocks:
            if b["kind"] == "heading" and b["level"] <= 1:
                current_ch = {"title": _strip_numbering(b["text"]), "level": 1,
                              "paras": [], "subs": []}
                chapters.append(current_ch)
                current_sub = None
                current_sub3 = None
            elif b["kind"] == "heading" and b["level"] == 2:
                if current_ch is None:
                    current_ch = {"title": PLACEHOLDER, "level": 1,
                                  "paras": [], "subs": []}
                    chapters.append(current_ch)
                current_sub = {"title": _strip_numbering(b["text"]), "level": 2,
                               "paras": [], "subs": []}
                current_ch["subs"].append(current_sub)
                current_sub3 = None
            elif b["kind"] == "heading":  # level >= 3
                if current_sub is None:
                    # 无上级小节：提升为二级
                    if current_ch is None:
                        current_ch = {"title": PLACEHOLDER, "level": 1,
                                      "paras": [], "subs": []}
                        chapters.append(current_ch)
                    current_sub = {"title": _strip_numbering(b["text"]),
                                   "level": 2, "paras": [], "subs": []}
                    current_ch["subs"].append(current_sub)
                    current_sub3 = None
                else:
                    current_sub3 = {"title": _strip_numbering(b["text"]),
                                    "level": 3, "paras": []}
                    current_sub["subs"].append(current_sub3)
            else:  # 正文/列表/代码/表格文本
                text = b["text"] if b["kind"] != "table" else _table_to_text(b)
                if not text:
                    continue
                target = current_sub3 or current_sub or current_ch
                if target is not None:
                    target["paras"].append(text)
                else:
                    # 出现在任何标题之前的段落 -> 前言缓冲
                    if not chapters or chapters[0]["title"] != "前言":
                        chapters.insert(0, {"title": "前言", "level": 1,
                                            "paras": [], "subs": []})
                    chapters[0]["paras"].append(text)
```

`_build_deck` 中把段落汇集改为包含三级：

```python
        paras = list(ch["paras"])
        for sub in ch.get("subs", []):
            paras.extend(sub["paras"])
```

改为：

```python
        paras = list(ch["paras"])
        for sub in ch.get("subs", []):
            paras.extend(sub["paras"])
            for sub3 in sub.get("subs", []):
                paras.extend(sub3["paras"])
```

- [ ] **Step 4: 实现 docx_builder**

`src/docx_builder.py` 的 `build()` 中，二级小节循环体：

```python
        for si, sub in enumerate(ch.get("subs", []), 1):
            h2 = doc.add_heading(level=2)
            run = h2.add_run(f"{ci}.{si}　{sub['title']}")
            _set_run_font(run, W["headings"][2]["font_cn"],
                          W["headings"][2]["font_en"],
                          W["headings"][2]["size_pt"], True)
            for para in sub["paras"]:
                _add_para(doc, para, W["body"], indent_chars=2)
```

末尾追加三级渲染，整体替换为：

```python
        for si, sub in enumerate(ch.get("subs", []), 1):
            h2 = doc.add_heading(level=2)
            run = h2.add_run(f"{ci}.{si}　{sub['title']}")
            _set_run_font(run, W["headings"][2]["font_cn"],
                          W["headings"][2]["font_en"],
                          W["headings"][2]["size_pt"], True)
            for para in sub["paras"]:
                _add_para(doc, para, W["body"], indent_chars=2)
            for ti, sub3 in enumerate(sub.get("subs", []), 1):
                h3 = doc.add_heading(level=3)
                run = h3.add_run(f"{ci}.{si}.{ti}　{sub3['title']}")
                _set_run_font(run, W["headings"][3]["font_cn"],
                              W["headings"][3]["font_en"],
                              W["headings"][3]["size_pt"], True)
                for para in sub3["paras"]:
                    _add_para(doc, para, W["body"], indent_chars=2)
```

- [ ] **Step 5: 运行确认通过（含回归）**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: 全部 passed

- [ ] **Step 6: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/organizer.py thesis_project/src/docx_builder.py thesis_project/tests/test_organizer_level3.py
git commit -m "feat(organizer,docx): support level-3 headings in chapter tree and docx" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 7: 表格全量转文本

**缺陷：** `organizer.py:138-140` 的 `_table_to_text` 只取前 2 行，其余数据丢失。

**Files:**
- Modify: `thesis_project/src/organizer.py`（`_table_to_text`）
- Test: `thesis_project/tests/test_organizer_table.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_organizer_table.py`：

```python
# -*- coding: utf-8 -*-
from src.organizer import _table_to_text
from tests.factories import table


def test_all_rows_kept():
    blk = table([["指标", "数值"],
                 ["准确率", "94.2%"],
                 ["延迟", "85ms"],
                 ["体积", "4MB"]])
    text = _table_to_text(blk)
    assert "准确率 | 94.2%" in text
    assert "延迟 | 85ms" in text
    assert "体积 | 4MB" in text


def test_empty_cells_skipped():
    blk = table([["a", "", "b"], ["", "", ""]])
    assert _table_to_text(blk) == "a | b"
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_organizer_table.py -v`
Expected: FAIL（第 3、4 行数据不在输出里）

- [ ] **Step 3: 实现**

`src/organizer.py` 中把：

```python
def _table_to_text(block):
    rows = block.get("rows") or []
    return " | ".join(" ".join(r) for r in rows[:2])
```

替换为：

```python
def _table_to_text(block):
    """表格 -> 多行文本：每行一条、竖线分隔、跳过空单元格与空行。"""
    rows = block.get("rows") or []
    lines = []
    for r in rows:
        cells = [c.strip() for c in r if c and c.strip()]
        if cells:
            lines.append(" | ".join(cells))
    return "\n".join(lines)
```

- [ ] **Step 4: 运行确认通过（含回归）**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: 全部 passed

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/organizer.py thesis_project/tests/test_organizer_table.py
git commit -m "fix(organizer): keep all table rows when converting to text" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 8: PDF 读取改进（分段、标题识别、表格去重）

**缺陷：** `readers.py:226` 按 `\n\s*\n` 切 PDF 文本几乎总得到"每页一大段"；表格内容同时出现在 `extract_tables` 和 `extract_text` 里造成重复。

**方案：** 新增两个纯函数（可单测，无需构造真实 PDF）：`_join_lines`（中文行直连、英文词间补空格）和 `_pdf_lines_to_blocks`（按行合并成段，行尾终止标点断段；短编号行识别为标题）。pdfplumber 路径用 `page.find_tables()` 的 bbox 过滤正文，避免表格文字重复。

**Files:**
- Modify: `thesis_project/src/readers.py`（新增 `_join_lines`/`_pdf_heading_level`/`_pdf_lines_to_blocks`，重写 `read_pdf` 的 pdfplumber 路径）
- Test: `thesis_project/tests/test_readers_pdf.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_readers_pdf.py`：

```python
# -*- coding: utf-8 -*-
from src.readers import _join_lines, _pdf_lines_to_blocks


def test_join_lines_chinese_no_space():
    assert _join_lines(["深度学习在图像", "识别领域取得突破。"]) == \
        "深度学习在图像识别领域取得突破。"


def test_join_lines_english_with_space():
    assert _join_lines(["deep learning", "model"]) == "deep learning model"


def test_heading_detected():
    blocks = []
    _pdf_lines_to_blocks("第一章 绪论\n本文研究了垃圾分类\n问题的解决方案。", blocks)
    assert blocks[0]["kind"] == "heading"
    assert blocks[0]["level"] == 1
    assert blocks[0]["text"] == "第一章 绪论"
    assert blocks[1]["kind"] == "paragraph"
    assert blocks[1]["text"] == "本文研究了垃圾分类问题的解决方案。"


def test_numbered_subheading_level():
    blocks = []
    _pdf_lines_to_blocks("1.1 研究背景\n内容一。", blocks)
    assert blocks[0]["kind"] == "heading"
    assert blocks[0]["level"] == 2


def test_sentence_end_splits_paragraphs():
    blocks = []
    _pdf_lines_to_blocks("第一句话结束。\n第二句话开始，\n然后结束。", blocks)
    paras = [b["text"] for b in blocks if b["kind"] == "paragraph"]
    assert paras == ["第一句话结束。", "第二句话开始，然后结束。"]


def test_long_unpunctuated_line_not_heading():
    blocks = []
    _pdf_lines_to_blocks("2023 年以来国内外研究者在垃圾分类领域开展了大量卓有成效的工作\n并取得进展。", blocks)
    assert blocks[0]["kind"] == "paragraph"
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_readers_pdf.py -v`
Expected: ERROR，`cannot import name '_join_lines'`

- [ ] **Step 3: 实现纯函数**

在 `src/readers.py` 的 `read_pdf` 之前新增：

```python
def _join_lines(lines):
    """合并断行：ASCII 单词之间补空格，中文直接相连。"""
    out = ""
    for ln in lines:
        if (out and out[-1].isascii() and out[-1].isalnum()
                and ln[:1].isascii() and ln[:1].isalnum()):
            out += " "
        out += ln
    return out


def _pdf_heading_level(prefix: str) -> int:
    if "章" in prefix:
        return 1
    m = re.match(r"\d+(\.\d+)*", prefix)
    return min(m.group(0).count(".") + 1, 3) if m else 1


_PDF_HEADING = re.compile(r"^(第\s*[一二三四五六七八九十百\d]+\s*章|\d+(\.\d+)*[\s、.．])")


def _pdf_lines_to_blocks(txt: str, blocks: list) -> None:
    """PDF 文本按行重组：行尾终止标点断段；短编号行识别为标题。

    PDF 提取的文本几乎没有连续空行，不能按空行分段；
    这里以「句末标点在行尾」为段落边界，是对中文论文的合理近似。
    """
    buf = []

    def flush():
        if buf:
            blocks.append(_block("paragraph", _join_lines(buf)))
            buf.clear()

    for line in txt.splitlines():
        line = _clean(line)
        if not line:
            flush()
            continue
        m = _PDF_HEADING.match(line)
        if (m and len(line) <= 25
                and not re.search(r"[。！？；，,;:.]\s*$", line)):
            flush()
            blocks.append(_block("heading", line,
                                 level=_pdf_heading_level(m.group(1))))
            continue
        buf.append(line)
        if re.search(r"[。！？!?]$", line):
            flush()
    flush()
```

- [ ] **Step 4: 运行纯函数测试确认通过**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_readers_pdf.py -v`
Expected: 6 passed

- [ ] **Step 5: 重写 read_pdf 的 pdfplumber 路径**

`src/readers.py` 中把 `read_pdf` 的 pdfplumber 分支：

```python
    if pdfplumber is not None:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                # 表格
                for tbl in page.extract_tables() or []:
                    rows = [[(_clean(c) if c else "") for c in row] for row in tbl]
                    if rows:
                        blocks.append(_block("table", "", rows=rows))
                # 正文
                txt = page.extract_text() or ""
                for para in re.split(r"\n\s*\n", txt):
                    para = _clean(para.replace("\n", " "))
                    if para:
                        blocks.append(_block("paragraph", para))
```

替换为：

```python
    if pdfplumber is not None:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                # 表格（find_tables 以便拿到 bbox 用于正文去重）
                tables = page.find_tables()
                bboxes = []
                for tbl in tables:
                    rows = [[(_clean(c) if c else "") for c in row]
                            for row in (tbl.extract() or [])]
                    rows = [r for r in rows if any(r)]
                    if rows:
                        blocks.append(_block("table", "", rows=rows))
                        bboxes.append(tbl.bbox)

                # 正文：过滤掉落在表格 bbox 内的字符，避免内容重复
                def _outside(obj, _bx=tuple(bboxes)):
                    cx = (obj["x0"] + obj["x1"]) / 2
                    cy = (obj["top"] + obj["bottom"]) / 2
                    return not any(x0 <= cx <= x1 and y0 <= cy <= y1
                                   for (x0, y0, x1, y1) in _bx)

                target = page.filter(_outside) if bboxes else page
                txt = target.extract_text() or ""
                _pdf_lines_to_blocks(txt, blocks)
```

同时把 pypdf 退化路径中的：

```python
            for para in re.split(r"\n\s*\n", txt):
                para = _clean(para.replace("\n", " "))
                if para:
                    blocks.append(_block("paragraph", para))
```

替换为：

```python
            _pdf_lines_to_blocks(txt, blocks)
```

- [ ] **Step 6: 运行全部测试确认通过**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: 全部 passed

- [ ] **Step 7: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/readers.py thesis_project/tests/test_readers_pdf.py
git commit -m "fix(readers): line-based pdf paragraphing, heading detection, table dedup" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 9: PPT 要点溢出分页 + 截断优化 + 页脚死代码清理

**缺陷：** 超过 6 条的要点被 `pptx_builder.py:146` 的 `bullets[:6]` 静默丢弃；`_to_bullets` 硬截 40 字产生残句；`pptx_builder.py:91` 的 `p2 = tf.add_paragraph()` 是死代码。

**Files:**
- Modify: `thesis_project/src/organizer.py`（`_to_bullets` 截断优化、`_build_deck` 分页、顶部导入 PPT_SPEC）
- Modify: `thesis_project/src/pptx_builder.py`（删除 `_footer` 死代码）
- Test: `thesis_project/tests/test_deck_overflow.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_deck_overflow.py`：

```python
# -*- coding: utf-8 -*-
from src.organizer import organize, _to_bullets
from tests.factories import h, p, doc


def test_bullets_overflow_creates_continuation_slide():
    sentences = "".join(f"第{i}个要点内容够长可以入选。" for i in range(1, 11))
    docs = [doc([h(1, "系统实现"), p(sentences)])]
    _, deck = organize(docs)
    contents = [s for s in deck["slides"]
                if s["type"] == "content" and s["title"].startswith("系统实现")]
    assert len(contents) == 2
    assert len(contents[0]["bullets"]) == 6
    assert contents[1]["title"] == "系统实现（续）"
    assert len(contents[1]["bullets"]) == 4
    # 无静默丢弃：10 个要点全部保留
    assert sum(len(s["bullets"]) for s in contents) == 10


def test_no_continuation_when_few_bullets():
    docs = [doc([h(1, "系统实现"), p("只有一个要点。")])]
    _, deck = organize(docs)
    contents = [s for s in deck["slides"]
                if s["type"] == "content" and s["title"].startswith("系统实现")]
    assert len(contents) == 1


def test_truncation_cuts_at_comma():
    long = "本系统采用了轻量化设计方案与迁移学习方法，通过知识蒸馏进一步压缩模型体积并保持精度水平"
    bullets = _to_bullets([long + "。"])
    # 在逗号处截断而不是硬切 40 字
    assert bullets[0] == "本系统采用了轻量化设计方案与迁移学习方法…"
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_deck_overflow.py -v`
Expected: 第 1、3 个 FAIL（只有 1 页 content、截断带残句），第 2 个 PASS

- [ ] **Step 3: 实现 organizer**

`src/organizer.py` 顶部 `import re` 之后新增：

```python
from config.format_spec import PPT_SPEC

_MAX_BULLETS_PER_SLIDE = PPT_SPEC["layout"]["max_bullets_per_slide"]
```

`_to_bullets` 整体替换为（默认上限提到 12 = 两页量，截断优先落在逗号/顿号）：

```python
def _to_bullets(paras, max_bullets=12, max_len=40):
    bullets = []
    for p in paras:
        # 按句号/分号切，取较短句作要点
        for seg in re.split(r"[。；;\n]", p):
            seg = seg.strip()
            if len(seg) < 4:
                continue
            if len(seg) > max_len:
                cut = max(seg.rfind(c, 0, max_len + 1) for c in "，、,")
                seg = seg[:cut] if cut >= max_len // 2 else seg[:max_len]
                seg += "…"
            bullets.append(seg)
            if len(bullets) >= max_bullets:
                return bullets
    return bullets or ["<待补充要点>"]
```

`_build_deck` 中把内容页生成：

```python
        slides.append({"type": "section", "title": label[key]})
        for item in group:
            slides.append({"type": "content",
                           "title": item["title"],
                           "bullets": item["bullets"]})
```

替换为（每页 ≤6 条，溢出自动生成"（续）"页）：

```python
        slides.append({"type": "section", "title": label[key]})
        for item in group:
            bullets = item["bullets"]
            chunks = [bullets[i:i + _MAX_BULLETS_PER_SLIDE]
                      for i in range(0, len(bullets), _MAX_BULLETS_PER_SLIDE)]
            for pi, chunk in enumerate(chunks):
                title = item["title"] if pi == 0 else f"{item['title']}（续）"
                slides.append({"type": "content", "title": title,
                               "bullets": chunk})
```

- [ ] **Step 4: 清理 pptx_builder 死代码**

`src/pptx_builder.py` 的 `_footer` 中删除这两行：

```python
    p2 = tf.add_paragraph()  # 右侧页码用第二段模拟——改为同段右对齐
    # 用制表符不易控，直接放页码到右侧文本框
```

（`bullets[:P["layout"]["max_bullets_per_slide"]]` 保留不动，作为 builder 端的最后防线。）

- [ ] **Step 5: 运行确认通过（含回归）**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: 全部 passed

- [ ] **Step 6: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/organizer.py thesis_project/src/pptx_builder.py thesis_project/tests/test_deck_overflow.py
git commit -m "feat(deck): paginate overflowing bullets, comma-aware truncation, remove dead code" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 10: LLM 客户端基础（llm_enhancer.py）

**Files:**
- Create: `thesis_project/src/llm_enhancer.py`
- Test: `thesis_project/tests/test_llm_base.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_llm_base.py`：

```python
# -*- coding: utf-8 -*-
import pytest
from src import llm_enhancer


def test_not_available_without_key(monkeypatch):
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    assert llm_enhancer.is_available() is False


def test_available_with_key(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")
    assert llm_enhancer.is_available() is True


def test_chat_json_plain(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat", lambda s, u: '{"a": 1}')
    assert llm_enhancer._chat_json("s", "u") == {"a": 1}


def test_chat_json_fenced(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat",
                        lambda s, u: '```json\n{"a": [1, 2]}\n```')
    assert llm_enhancer._chat_json("s", "u") == {"a": [1, 2]}


def test_chat_json_with_surrounding_prose(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat",
                        lambda s, u: '好的，结果如下：{"a": 1} 以上就是结果。')
    assert llm_enhancer._chat_json("s", "u") == {"a": 1}


def test_chat_json_no_json_raises(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat", lambda s, u: "抱歉，我无法处理。")
    with pytest.raises(ValueError):
        llm_enhancer._chat_json("s", "u")
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_llm_base.py -v`
Expected: ERROR，`No module named 'src.llm_enhancer'`

- [ ] **Step 3: 实现**

新建 `thesis_project/src/llm_enhancer.py`：

```python
# -*- coding: utf-8 -*-
"""
LLM 增强层 —— 通过 OpenAI 兼容接口提升草案质量（可选，失败自动回退）。

配置（环境变量）：
    LLM_API_KEY    必填；未设置则增强层整体跳过
    LLM_BASE_URL   选填；默认官方 OpenAI；可指向 DeepSeek/通义/Kimi/Ollama 等兼容端点
    LLM_MODEL      选填；默认 gpt-4o-mini

原则：
  - LLM 只做「整理 / 提炼 / 翻译 / 分类」，不扩写正文、不编造事实。
  - AI 生成的摘要类内容带标记 <AI生成，请核对>，与 <请填写> 同一哲学。
  - 任意一步失败只打印告警并保留规则结果，绝不让主流程崩溃。
  - 测试中不真实调用 API：_chat 是唯一网络出口，打桩即可。
"""
from __future__ import annotations
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

AI_MARK = "<AI生成，请核对>"
_VALID_BUCKETS = {"background", "method", "result", "conclusion"}


def is_available() -> bool:
    return bool(os.environ.get("LLM_API_KEY"))


def _client():
    from openai import OpenAI
    return OpenAI(api_key=os.environ["LLM_API_KEY"],
                  base_url=os.environ.get("LLM_BASE_URL") or None)


def _chat(system: str, user: str) -> str:
    """单轮对话，返回原始文本。唯一网络出口，便于测试打桩。"""
    resp = _client().chat.completions.create(
        model=os.environ.get("LLM_MODEL", "gpt-4o-mini"),
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": user}],
        temperature=0.2,
    )
    return resp.choices[0].message.content or ""


def _chat_json(system: str, user: str):
    """要求模型输出 JSON 并解析；容忍代码块包裹与前后废话。"""
    text = _chat(system + "\n只输出 JSON，不要任何其它文字。", user)
    m = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL)
    if m:
        text = m.group(1)
    starts = [i for i in (text.find("{"), text.find("[")) if i >= 0]
    if not starts:
        raise ValueError(f"LLM 未返回 JSON：{text[:80]!r}")
    obj, _ = json.JSONDecoder().raw_decode(text[min(starts):].strip())
    return obj
```

- [ ] **Step 4: 运行确认通过**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_llm_base.py -v`
Expected: 6 passed

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/llm_enhancer.py thesis_project/tests/test_llm_base.py
git commit -m "feat(llm): add OpenAI-compatible client base with robust JSON parsing" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 11: LLM 元信息抽取 + 英文摘要/关键词翻译

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py`（新增 `refine_meta`、`translate_abstract`）
- Test: `thesis_project/tests/test_llm_meta.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_llm_meta.py`：

```python
# -*- coding: utf-8 -*-
from src import llm_enhancer
from src.llm_enhancer import AI_MARK
from src.organizer import PLACEHOLDER
from tests.factories import p, doc


def _thesis(**over):
    t = {"title": PLACEHOLDER, "author": PLACEHOLDER,
         "abstract": PLACEHOLDER, "abstract_en": PLACEHOLDER,
         "keywords": [PLACEHOLDER], "keywords_en": [PLACEHOLDER],
         "chapters": [], "references": [], "auto_skeleton": False}
    t.update(over)
    return t


def test_refine_meta_fills_placeholders(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "title": "某垃圾分类系统研究", "author": "张三",
        "abstract": "本文研究了垃圾分类。", "keywords": ["深度学习", "分类"]})
    t = _thesis()
    llm_enhancer.refine_meta(t, [doc([p("正文片段。")])])
    assert t["title"] == "某垃圾分类系统研究"
    assert t["author"] == "张三"
    assert t["abstract"] == f"本文研究了垃圾分类。 {AI_MARK}"
    assert t["keywords"] == ["深度学习", "分类"]


def test_refine_meta_keeps_existing_values(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "title": "LLM乱给的题目", "author": "LLM乱给的作者",
        "abstract": "x", "keywords": ["x"]})
    t = _thesis(title="真题目", author="真作者",
                abstract="真摘要", keywords=["真关键词"])
    llm_enhancer.refine_meta(t, [doc([p("正文。")])])
    assert t["title"] == "真题目"
    assert t["author"] == "真作者"
    assert t["abstract"] == "真摘要"
    assert t["keywords"] == ["真关键词"]


def test_refine_meta_skips_llm_when_nothing_needed(monkeypatch):
    def boom(s, u):
        raise AssertionError("不应调用 LLM")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    t = _thesis(title="a", author="b", abstract="c", keywords=["d"])
    llm_enhancer.refine_meta(t, [doc([p("正文。")])])


def test_translate_abstract(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "abstract_en": "This paper studies waste sorting.",
        "keywords_en": ["deep learning", "classification"]})
    t = _thesis(abstract="本文研究了垃圾分类。", keywords=["深度学习"])
    llm_enhancer.translate_abstract(t)
    assert t["abstract_en"] == f"This paper studies waste sorting. {AI_MARK}"
    assert t["keywords_en"] == ["deep learning", "classification"]


def test_translate_abstract_skipped_when_no_abstract(monkeypatch):
    def boom(s, u):
        raise AssertionError("不应调用 LLM")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    t = _thesis()  # abstract 仍是占位符
    llm_enhancer.translate_abstract(t)
    assert t["abstract_en"] == PLACEHOLDER
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_llm_meta.py -v`
Expected: ERROR，`has no attribute 'refine_meta'`

- [ ] **Step 3: 实现**

在 `src/llm_enhancer.py` 末尾追加：

```python
# ---------------------------------------------------------------------------
#  元信息抽取 / 英文摘要翻译
# ---------------------------------------------------------------------------
_META_SYS = ("你是论文排版助手。根据用户提供的论文草稿开头片段抽取元信息。"
             "只能从原文归纳，禁止编造；无法确定的字段输出空字符串或空列表。")

_EN_SYS = ("你是学术翻译。把给定的中文论文摘要与关键词翻译成规范的学术英文。"
           "忠实原文，不添加原文没有的信息。")


def refine_meta(thesis: dict, docs: list) -> None:
    """就地补全 title/author/abstract/keywords 中仍为占位符的字段。"""
    from src.organizer import PLACEHOLDER
    need = [k for k in ("title", "author", "abstract")
            if thesis[k] == PLACEHOLDER]
    if thesis["keywords"] == [PLACEHOLDER]:
        need.append("keywords")
    if not need:
        return
    head = "\n".join(b["text"] for d in docs
                     for b in d["blocks"][:40] if b["text"])[:3000]
    data = _chat_json(
        _META_SYS,
        '请以 JSON 输出 {"title": "...", "author": "...", '
        '"abstract": "...", "keywords": ["...", "..."]}：\n\n' + head)
    for k in need:
        v = data.get(k)
        if not v:
            continue
        if k == "keywords":
            thesis[k] = [str(x).strip() for x in v if str(x).strip()][:5]
        elif k == "abstract":
            thesis[k] = f"{str(v).strip()} {AI_MARK}"
        else:
            thesis[k] = str(v).strip()


def translate_abstract(thesis: dict) -> None:
    """中文摘要/关键词 -> 英文（就地写入 abstract_en / keywords_en）。"""
    from src.organizer import PLACEHOLDER
    ab = thesis["abstract"].replace(AI_MARK, "").strip()
    if not ab or ab == PLACEHOLDER:
        return
    kws = [k for k in thesis["keywords"] if k != PLACEHOLDER]
    data = _chat_json(
        _EN_SYS,
        '请以 JSON 输出 {"abstract_en": "...", "keywords_en": ["..."]}：\n\n'
        f"摘要：{ab}\n关键词：{'；'.join(kws)}")
    if data.get("abstract_en"):
        thesis["abstract_en"] = f"{str(data['abstract_en']).strip()} {AI_MARK}"
    if data.get("keywords_en"):
        thesis["keywords_en"] = [str(x).strip()
                                 for x in data["keywords_en"]
                                 if str(x).strip()][:5]
```

- [ ] **Step 4: 运行确认通过**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_llm_meta.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/llm_enhancer.py thesis_project/tests/test_llm_meta.py
git commit -m "feat(llm): meta extraction and english abstract translation" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 12: deck 生成注入点（classify_fn / bullets_fn）+ LLM 章节分类

**缺陷：** `organizer._classify` 关键词命不中一律归 method。
**方案：** `_build_deck` 增加可选 `classify_fn`、`bullets_fn` 参数（默认原规则函数，规则路径行为完全不变）；llm_enhancer 提供 `classify_chapters`。

**Files:**
- Modify: `thesis_project/src/organizer.py`（`_build_deck` 签名与内部）
- Modify: `thesis_project/src/llm_enhancer.py`（新增 `classify_chapters`）
- Test: `thesis_project/tests/test_llm_classify.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_llm_classify.py`：

```python
# -*- coding: utf-8 -*-
from src import llm_enhancer
from src.organizer import _build_deck


def _ch(title, paras=("内容一二三。",)):
    # 注意：要点须 >=4 字，否则被 _to_bullets 的短句过滤丢弃
    return {"title": title, "level": 1, "paras": list(paras), "subs": []}


def test_build_deck_accepts_custom_classifier():
    deck_meta = {"title": "T", "author": "A"}
    deck = _build_deck(deck_meta, [_ch("某个古怪标题")],
                       classify_fn=lambda t: "result")
    slides = deck["slides"]
    sec_idx = next(i for i, s in enumerate(slides)
                   if s["type"] == "section" and s["title"] == "研究成果")
    assert slides[sec_idx + 1]["type"] == "content"
    assert slides[sec_idx + 1]["title"] == "某个古怪标题"


def test_build_deck_accepts_custom_bullets_fn():
    deck = _build_deck({"title": "T", "author": "A"}, [_ch("绪论")],
                       bullets_fn=lambda paras: ["自定义要点"])
    contents = [s for s in deck["slides"] if s["type"] == "content"]
    assert contents[0]["bullets"] == ["自定义要点"]


def test_build_deck_default_unchanged():
    deck = _build_deck({"title": "T", "author": "A"}, [_ch("绪论")])
    contents = [s for s in deck["slides"] if s["type"] == "content"]
    assert contents[0]["bullets"] == ["内容一二三"]


def test_classify_chapters_filters_invalid(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "绪论": "background", "怪章": "nonsense", "实验": "result"})
    got = llm_enhancer.classify_chapters(["绪论", "怪章", "实验"])
    assert got == {"绪论": "background", "实验": "result"}
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_llm_classify.py -v`
Expected: FAIL/ERROR（`_build_deck` 不接受关键字参数、`classify_chapters` 不存在）

- [ ] **Step 3: 实现 organizer 注入点**

`src/organizer.py` 的 `_build_deck` 开头：

```python
def _build_deck(meta, chapters):
    """把章节映射到 PPT 7 段结构。"""
    buckets = {"background": [], "method": [], "result": [], "conclusion": []}
    for ch in chapters:
        key = _classify(ch["title"])
```

改为：

```python
def _build_deck(meta, chapters, classify_fn=None, bullets_fn=None):
    """把章节映射到 PPT 7 段结构。

    classify_fn(title)->bucket、bullets_fn(paras)->list[str] 为可选注入点，
    默认用规则实现（_classify / _to_bullets），供 LLM 增强层替换。
    """
    classify = classify_fn or _classify
    to_bullets = bullets_fn or _to_bullets
    buckets = {"background": [], "method": [], "result": [], "conclusion": []}
    for ch in chapters:
        key = classify(ch["title"])
```

同函数内把：

```python
        buckets[key].append({"title": ch["title"], "bullets": _to_bullets(paras)})
```

改为：

```python
        buckets[key].append({"title": ch["title"], "bullets": to_bullets(paras)})
```

- [ ] **Step 4: 实现 llm_enhancer.classify_chapters**

在 `src/llm_enhancer.py` 末尾追加：

```python
# ---------------------------------------------------------------------------
#  章节 -> PPT 分区 语义分类
# ---------------------------------------------------------------------------
_CLS_SYS = ("你是答辩PPT助手。把论文章节标题分类到四个部分之一："
            "background（背景/绪论/综述）、method（方法/设计/理论/需求）、"
            "result（实现/实验/测试/结果）、conclusion（总结/结论/展望）。")


def classify_chapters(titles) -> dict:
    """返回 {标题: bucket}，只保留合法 bucket；缺失的标题由调用方回退规则分类。"""
    data = _chat_json(
        _CLS_SYS,
        '请以 JSON 输出 {"章节标题": "background|method|result|conclusion", ...}：\n'
        + json.dumps(list(titles), ensure_ascii=False))
    if not isinstance(data, dict):
        raise ValueError("分类结果不是 JSON 对象")
    return {str(t): b for t, b in data.items() if b in _VALID_BUCKETS}
```

- [ ] **Step 5: 运行确认通过（含回归）**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: 全部 passed

- [ ] **Step 6: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/organizer.py thesis_project/src/llm_enhancer.py thesis_project/tests/test_llm_classify.py
git commit -m "feat(deck): classify_fn/bullets_fn injection points + llm chapter classifier" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 13: LLM 要点提炼 + deck 重建编排

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py`（新增 `_llm_bullets`、`_safe_bullets`、`rebuild_deck`）
- Test: `thesis_project/tests/test_llm_bullets.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_llm_bullets.py`：

```python
# -*- coding: utf-8 -*-
from src import llm_enhancer
from src.organizer import PLACEHOLDER


def _thesis():
    return {"title": "某系统研究", "author": "张三",
            "abstract": "摘要。", "abstract_en": PLACEHOLDER,
            "keywords": ["k"], "keywords_en": [PLACEHOLDER],
            "chapters": [
                {"title": "绪论", "level": 1,
                 "paras": ["研究背景很重要。研究意义也很大。"], "subs": []},
                {"title": "某个古怪标题", "level": 1,
                 "paras": ["做了实验，效果不错。"], "subs": []},
            ],
            "references": [], "auto_skeleton": False}


def test_llm_bullets_used(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json",
                        lambda s, u: {"bullets": ["要点一", "要点二"]})
    assert llm_enhancer._llm_bullets(["原文段落。"]) == ["要点一", "要点二"]


def test_llm_bullets_empty_paras_placeholder(monkeypatch):
    def boom(s, u):
        raise AssertionError("空内容不应调用 LLM")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    assert llm_enhancer._llm_bullets([]) == ["<待补充要点>"]
    assert llm_enhancer._llm_bullets([PLACEHOLDER]) == ["<待补充要点>"]


def test_safe_bullets_falls_back_on_error(monkeypatch):
    def boom(s, u):
        raise RuntimeError("api down")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    got = llm_enhancer._safe_bullets(["规则截取的句子够长可以入选。"])
    assert got == ["规则截取的句子够长可以入选"]


def test_rebuild_deck_uses_llm_classify_and_bullets(monkeypatch):
    def fake_chat_json(system, user):
        if "分类" in system or "background" in system:
            return {"某个古怪标题": "result"}
        return {"bullets": ["LLM要点"]}
    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    deck = llm_enhancer.rebuild_deck(_thesis())
    slides = deck["slides"]
    sec_idx = next(i for i, s in enumerate(slides)
                   if s["type"] == "section" and s["title"] == "研究成果")
    assert slides[sec_idx + 1]["title"] == "某个古怪标题"
    assert slides[sec_idx + 1]["bullets"] == ["LLM要点"]


def test_rebuild_deck_classify_failure_falls_back(monkeypatch):
    def fake_chat_json(system, user):
        if "background" in system:
            raise RuntimeError("api down")
        return {"bullets": ["LLM要点"]}
    monkeypatch.setattr(llm_enhancer, "_chat_json", fake_chat_json)
    deck = llm_enhancer.rebuild_deck(_thesis())
    # 分类失败 -> 规则分类："某个古怪标题" 命不中关键词 -> method
    slides = deck["slides"]
    sec_idx = next(i for i, s in enumerate(slides)
                   if s["type"] == "section" and s["title"] == "研究方法与过程")
    titles_after = [s["title"] for s in slides[sec_idx + 1:sec_idx + 3]]
    assert "某个古怪标题" in titles_after
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_llm_bullets.py -v`
Expected: ERROR，`has no attribute '_llm_bullets'`

- [ ] **Step 3: 实现**

在 `src/llm_enhancer.py` 末尾追加：

```python
# ---------------------------------------------------------------------------
#  PPT 要点提炼 + deck 重建
# ---------------------------------------------------------------------------
_BULLET_SYS = ("你是答辩PPT助手。把给定原文提炼成要点列表：最多6条、每条不超过40字、"
               "用名词短语或短句。只依据原文归纳，不得引入原文没有的内容。")


def _llm_bullets(paras) -> list:
    from src.organizer import PLACEHOLDER
    src = "\n".join(p for p in paras if p and p != PLACEHOLDER)[:2000]
    if not src.strip():
        return ["<待补充要点>"]
    data = _chat_json(_BULLET_SYS,
                      '请以 JSON 输出 {"bullets": ["...", "..."]}：\n\n' + src)
    bullets = [str(b).strip()[:40] for b in data.get("bullets", [])
               if str(b).strip()]
    return bullets[:6] or ["<待补充要点>"]


def _safe_bullets(paras) -> list:
    """LLM 提炼失败时回退规则截取。"""
    from src.organizer import _to_bullets
    try:
        return _llm_bullets(paras)
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] 要点提炼失败，改用规则截取：{e}")
        return _to_bullets(paras)


def rebuild_deck(thesis: dict) -> dict:
    """用 LLM 分类 + LLM 要点重建 PPT 大纲；分类失败整体回退规则分类。"""
    from src.organizer import _build_deck, _classify
    try:
        mapping = classify_chapters([c["title"] for c in thesis["chapters"]])
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] 章节分类失败，改用关键词规则：{e}")
        mapping = {}
    meta = {"title": thesis["title"], "author": thesis["author"]}
    return _build_deck(meta, thesis["chapters"],
                       classify_fn=lambda t: mapping.get(t) or _classify(t),
                       bullets_fn=_safe_bullets)
```

- [ ] **Step 4: 运行确认通过**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_llm_bullets.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/llm_enhancer.py thesis_project/tests/test_llm_bullets.py
git commit -m "feat(llm): bullet refinement with rule fallback and deck rebuild" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 14: LLM 语义分章（无标题文档）

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py`（新增 `rechapter`）
- Test: `thesis_project/tests/test_llm_rechapter.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_llm_rechapter.py`：

```python
# -*- coding: utf-8 -*-
from src import llm_enhancer
from src.organizer import organize, PLACEHOLDER, DEFAULT_CHAPTERS
from tests.factories import p, doc


def _skeleton_thesis():
    paras = ["A段背景。", "B段方法。", "C段其它。", "D段展望。"]
    thesis, _ = organize([doc([p(t) for t in paras], type_="txt")])
    return thesis


def test_rechapter_assigns_in_order(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "绪论": [1, 0],            # 故意乱序，实现应按原文顺序排回
        "总结与展望": [3],
    })
    t = _skeleton_thesis()
    llm_enhancer.rechapter(t)
    by_title = {c["title"]: c for c in t["chapters"]}
    assert by_title["绪论"]["paras"] == ["A段背景。", "B段方法。"]
    assert by_title["总结与展望"]["paras"] == ["D段展望。"]
    # 未分配段落留在"研究内容"
    assert by_title["研究内容"]["paras"] == ["C段其它。"]
    # 空章保留占位符
    assert by_title["系统实现"]["paras"] == [PLACEHOLDER]


def test_rechapter_invalid_indices_ignored(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {
        "绪论": [0, 99, -1, "x"],
        "系统实现": [0],           # 与绪论重复 -> 后者忽略
    })
    t = _skeleton_thesis()
    llm_enhancer.rechapter(t)
    by_title = {c["title"]: c for c in t["chapters"]}
    assert by_title["绪论"]["paras"] == ["A段背景。"]
    assert by_title["系统实现"]["paras"] == [PLACEHOLDER]


def test_rechapter_noop_without_flag(monkeypatch):
    def boom(s, u):
        raise AssertionError("非骨架文档不应调用 LLM")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    t = _skeleton_thesis()
    t["auto_skeleton"] = False
    before = [c["title"] for c in t["chapters"]]
    llm_enhancer.rechapter(t)
    assert [c["title"] for c in t["chapters"]] == before


def test_rechapter_noop_when_llm_assigns_nothing(monkeypatch):
    monkeypatch.setattr(llm_enhancer, "_chat_json", lambda s, u: {})
    t = _skeleton_thesis()
    llm_enhancer.rechapter(t)
    by_title = {c["title"]: c for c in t["chapters"]}
    assert by_title["研究内容"]["paras"] == \
        ["A段背景。", "B段方法。", "C段其它。", "D段展望。"]
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_llm_rechapter.py -v`
Expected: ERROR，`has no attribute 'rechapter'`

- [ ] **Step 3: 实现**

在 `src/llm_enhancer.py` 末尾追加：

```python
# ---------------------------------------------------------------------------
#  无标题文档的语义分章
# ---------------------------------------------------------------------------
_CHAP_SYS = ("你是论文结构助手。把编号段落分配到给定的章节骨架中，"
             "依据语义就近归类；不确定的段落可以不分配。"
             "不得改写段落内容，每个编号至多出现一次。")


def rechapter(thesis: dict) -> None:
    """无标题文档：把"研究内容"中的段落按语义分配到骨架各章。

    章内保持原文顺序；LLM 未分配的段落留在"研究内容"；全部失败则不动。
    """
    from src.organizer import PLACEHOLDER, DEFAULT_CHAPTERS
    if not thesis.get("auto_skeleton"):
        return
    src = next((c for c in thesis["chapters"] if c["title"] == "研究内容"), None)
    if src is None or not src["paras"]:
        return
    paras = src["paras"]
    numbered = "\n".join(f"[{i}] {p[:200]}" for i, p in enumerate(paras))
    data = _chat_json(
        _CHAP_SYS,
        "章节骨架：" + json.dumps(DEFAULT_CHAPTERS, ensure_ascii=False)
        + '\n请以 JSON 输出 {"章节名": [段落编号, ...], ...}：\n\n' + numbered)
    if not isinstance(data, dict):
        return
    used = set()
    assign = {}
    for name in DEFAULT_CHAPTERS:
        idxs = sorted(i for i in data.get(name, [])
                      if isinstance(i, int) and 0 <= i < len(paras)
                      and i not in used)
        used.update(idxs)
        assign[name] = idxs
    if not used:
        return
    new_chapters = []
    for name in DEFAULT_CHAPTERS:
        ps = [paras[i] for i in assign[name]] or [PLACEHOLDER]
        new_chapters.append({"title": name, "level": 1,
                             "paras": ps, "subs": []})
    leftovers = [p for i, p in enumerate(paras) if i not in used]
    if leftovers:
        new_chapters.insert(3, {"title": "研究内容", "level": 1,
                                "paras": leftovers, "subs": []})
    thesis["chapters"] = new_chapters
```

- [ ] **Step 4: 运行确认通过**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_llm_rechapter.py -v`
Expected: 4 passed

- [ ] **Step 5: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/llm_enhancer.py thesis_project/tests/test_llm_rechapter.py
git commit -m "feat(llm): semantic re-chaptering for headingless documents" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 15: enhance 总入口 + main.py --llm + 文档 + 端到端验证

**Files:**
- Modify: `thesis_project/src/llm_enhancer.py`（新增 `enhance`）
- Modify: `thesis_project/src/main.py`（`--llm` 参数）
- Modify: `thesis_project/README.md`（LLM 用法段落）
- Test: `thesis_project/tests/test_e2e.py`

- [ ] **Step 1: 写失败测试**

`thesis_project/tests/test_e2e.py`：

```python
# -*- coding: utf-8 -*-
import os

from src import llm_enhancer, docx_builder, pptx_builder
from src.readers import read_dir
from src.organizer import organize

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def test_enhance_skipped_without_key(monkeypatch):
    monkeypatch.delenv("LLM_API_KEY", raising=False)
    docs = read_dir(os.path.join(ROOT, "sample_input"))
    thesis, deck = organize(docs)
    t2, d2 = llm_enhancer.enhance(thesis, deck, docs)
    assert t2 is thesis and d2 is deck  # 原样返回


def test_enhance_survives_step_failures(monkeypatch):
    monkeypatch.setenv("LLM_API_KEY", "sk-test")

    def boom(s, u):
        raise RuntimeError("api down")
    monkeypatch.setattr(llm_enhancer, "_chat_json", boom)
    docs = read_dir(os.path.join(ROOT, "sample_input"))
    thesis, deck = organize(docs)
    t2, d2 = llm_enhancer.enhance(thesis, deck, docs)
    # 全部步骤失败也要拿回可用的草案数据
    assert t2["title"] == "基于深度学习的校园垃圾图像分类系统设计与实现"
    assert any(s["type"] == "content" for s in d2["slides"])


def test_full_pipeline_sample(tmp_path):
    docs = read_dir(os.path.join(ROOT, "sample_input"))
    thesis, deck = organize(docs)
    assert thesis["title"] == "基于深度学习的校园垃圾图像分类系统设计与实现"
    assert thesis["auto_skeleton"] is False
    wp = docx_builder.build(thesis, str(tmp_path / "论文草案.docx"))
    pp = pptx_builder.build(deck, str(tmp_path / "答辩PPT草案.pptx"))
    assert os.path.getsize(wp) > 0
    assert os.path.getsize(pp) > 0
```

- [ ] **Step 2: 运行确认失败**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests/test_e2e.py -v`
Expected: 前两个 ERROR（`has no attribute 'enhance'`），第三个 PASS

- [ ] **Step 3: 实现 enhance 总入口**

在 `src/llm_enhancer.py` 末尾追加：

```python
# ---------------------------------------------------------------------------
#  总入口
# ---------------------------------------------------------------------------
def enhance(thesis: dict, deck: dict, docs: list):
    """依次执行各增强步骤，每步独立容错；返回 (thesis, deck)。"""
    if not is_available():
        print("  [LLM] 未设置 LLM_API_KEY，跳过增强（使用纯规则结果）。")
        return thesis, deck
    steps = [
        ("元信息抽取", lambda: refine_meta(thesis, docs)),
        ("语义分章", lambda: rechapter(thesis)),
        ("英文摘要翻译", lambda: translate_abstract(thesis)),
    ]
    for name, fn in steps:
        try:
            fn()
            print(f"  [LLM] {name} 完成")
        except Exception as e:  # noqa: BLE001
            print(f"  [LLM告警] {name} 失败，保留规则结果：{e}")
    try:
        deck = rebuild_deck(thesis)
        print("  [LLM] PPT 大纲重建完成")
    except Exception as e:  # noqa: BLE001
        print(f"  [LLM告警] PPT 大纲重建失败，保留规则结果：{e}")
    return thesis, deck
```

- [ ] **Step 4: main.py 接入 --llm**

`src/main.py` 中，在：

```python
    ap.add_argument("--only", choices=["word", "ppt"], help="只生成其中一种")
```

之后新增：

```python
    ap.add_argument("--llm", action="store_true",
                    help="用 LLM 增强草案质量（需设置 LLM_API_KEY，"
                         "可选 LLM_BASE_URL / LLM_MODEL，OpenAI 兼容接口）")
```

把：

```python
    print("② 整理内容结构")
    thesis, deck = organize(docs)
    print(f"  论文：{len(thesis['chapters'])} 章；PPT：{len(deck['slides'])} 页。")
```

改为：

```python
    print("② 整理内容结构")
    thesis, deck = organize(docs)
    if args.llm:
        print("②+ LLM 增强")
        from src import llm_enhancer
        thesis, deck = llm_enhancer.enhance(thesis, deck, docs)
    print(f"  论文：{len(thesis['chapters'])} 章；PPT：{len(deck['slides'])} 页。")
```

- [ ] **Step 5: 更新 README**

`README.md` 的「方式二：命令行」代码块中 `python src/main.py --output 某输出目录` 一行之后追加：

```
python src/main.py --llm                       # 用 LLM 增强草稿质量（可选）
```

「生成时落实的规范」小节之前插入新小节：

```markdown
---

## LLM 增强（可选）

设置环境变量后加 `--llm` 即可启用，能显著提升草稿质量：

| 环境变量 | 必填 | 说明 |
|---|---|---|
| `LLM_API_KEY` | 是 | API 密钥；未设置时自动跳过增强 |
| `LLM_BASE_URL` | 否 | OpenAI 兼容端点，如 DeepSeek/通义/Kimi/本地 Ollama |
| `LLM_MODEL` | 否 | 模型名，默认 `gpt-4o-mini` |

增强内容：元信息抽取（题目/作者/摘要/关键词）、英文摘要与关键词翻译、
PPT 要点语义提炼、章节到 PPT 分区的语义分类、无标题文档的语义分章。

原则：LLM 只做整理/提炼/翻译/分类，不扩写正文；AI 生成的内容带
`<AI生成，请核对>` 标记；任一步失败自动回退纯规则结果，主流程不受影响。

```powershell
# PowerShell 示例（DeepSeek）
$env:LLM_API_KEY  = "sk-..."
$env:LLM_BASE_URL = "https://api.deepseek.com"
$env:LLM_MODEL    = "deepseek-chat"
python src/main.py --llm
```
```

依赖清单一行：

```
依赖：`python-docx`、`python-pptx`、`pdfplumber`（读 PDF）。
```

改为：

```
依赖：`python-docx`、`python-pptx`、`pdfplumber`（读 PDF）；`openai`（仅 --llm 需要）。
```

- [ ] **Step 6: 全量测试 + 真实样例运行**

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python -m pytest tests -v`
Expected: 全部 passed

Run: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python src/main.py --input sample_input --output output`
Expected: 正常生成两份草案，无异常输出；`论文：5 章`（sample 有 5 个一级标题，特殊章节剔除不影响它）

Run（无 key 时 --llm 应优雅跳过）: `cd "D:/BackendDevelopment/Project/Project_Test-7/thesis_project" && python src/main.py --input sample_input --output output --llm`
Expected: 打印 `[LLM] 未设置 LLM_API_KEY，跳过增强`，仍正常生成两份草案

- [ ] **Step 7: Commit**

```bash
cd "D:/BackendDevelopment/Project/Project_Test-7"
git add thesis_project/src/llm_enhancer.py thesis_project/src/main.py thesis_project/README.md thesis_project/tests/test_e2e.py
git commit -m "feat: wire --llm flag with graceful fallback, docs and e2e tests" -m "Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## 验收清单（全部完成后逐项核对）

- [ ] `python -m pytest tests -v` 全绿
- [ ] GBK 编码 txt 能正确读取（Task 1）
- [ ] `# 绪论` 不再被当成论文题目（Task 2）
- [ ] 纯 txt 输入段落顺序保持、单章存放、带 `auto_skeleton` 标记（Task 3）
- [ ] "第一章 绪论"不再渲染成"第一章　第一章 绪论"（Task 4）
- [ ] 源文件参考文献被抽取、摘要/致谢章不再进正文（Task 5）
- [ ] `###` 三级标题进入 docx Heading 3（Task 6）
- [ ] 表格所有行保留（Task 7）
- [ ] PDF 按行重组分段、编号行识别为标题、表格文字不重复（Task 8）
- [ ] PPT 要点超 6 条自动分"（续）"页、逗号处截断（Task 9）
- [ ] 无 `LLM_API_KEY` 时 `--llm` 优雅跳过；有 key 但 API 全挂时仍产出规则草案（Task 10-15）
- [ ] `run.bat` 双击流程不受影响（未改动该文件）
