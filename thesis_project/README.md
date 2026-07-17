# 论文 Word 草案 + 答辩 PPT 草案 生成器

读取你给出的 **Word / PDF / TXT / Markdown / JSON** 源文件，自动整理成结构化内容，
再按 **本科毕业论文 Word 规范** 和 **答辩 PPT 规范** 生成两份草案：
`论文草案.docx` 与 `答辩PPT草案.pptx`。

> 定位：**草案**。产物已套好格式与结构骨架，正文与 `<请填写>` 占位符需人工润色。

---

## 目录结构

```
thesis_project/
├── run.bat                # ★ 一键启动（Windows 双击即可）
├── config/
│   └── format_spec.py     # 格式规范"硬标准"：字号/页边距/行距/标题层级/PPT结构
├── src/
│   ├── readers.py         # 多格式读取器 -> 统一中间结构
│   ├── organizer.py       # 内容整理：重建章节树、映射 PPT 结构、提炼要点
│   ├── docx_builder.py    # 按规范生成 .docx
│   ├── pptx_builder.py    # 按规范生成 .pptx
│   └── main.py            # 主管道（命令行入口）
├── sample_input/          # 示例源文件（可直接试跑）
├── input/                 # 把你的源文件放这里
└── output/                # 生成的草案输出到这里
```

---

## 快速开始

### 方式一：一键启动（Windows，推荐）

双击 **`run.bat`** 即可。它会自动：
1. 定位 Python（优先 `py`，其次 `python`）；
2. 缺依赖时自动 `pip install python-docx python-pptx pdfplumber`；
3. `input\` 有文件就用 `input\`，为空则用 `sample_input\` 演示；
4. 生成完成后自动打开 `output\` 目录。

### 方式二：命令行

```bash
# 1) 先用示例试跑（会在 output/ 生成两份草案）
python src/main.py --input sample_input

# 2) 放入自己的文件后正式生成
#    把 Word/PDF/TXT/md/json 放进 input/，然后：
python src/main.py

# 其它用法
python src/main.py --input a.pdf b.md 某目录   # 指定若干文件/目录
python src/main.py --only word                 # 只生成 Word
python src/main.py --only ppt                  # 只生成 PPT
python src/main.py --output 某输出目录          # 自定义输出目录
python src/main.py --llm                       # 用 LLM 增强草稿质量（可选）
```

> **Windows 提示**：请用 `python`（不要用 `python3`，它可能是应用商店占位程序）。
> 脚本已自动把控制台切到 UTF-8，中文不会乱码。

依赖：`python-docx`、`python-pptx`、`pdfplumber`（读 PDF）；`openai`（仅 --llm 需要）。
```bash
pip install python-docx python-pptx pdfplumber
```

---

## 输入文件怎么写效果最好

整理器会**识别标题层级**来重建章节：

- **Markdown**：用 `#`/`##` 作章/节标题，YAML frontmatter 写 `title/author/keywords`（见 `sample_input/thesis_draft.md`）。
- **Word (.docx)**：正文里用「标题 1/标题 2」样式的段落会被识别为章/节。
- **JSON**：支持 `{title, content, children:[...]}` 递归结构；也支持任意数据（会平铺为段落）。
- **PDF / TXT**：按空行分段读取（PDF 无显式标题时，会套用标准章节骨架，把内容顺序填入）。

识别不到的字段（题目、作者、英文摘要等）会留 `<请填写>` 占位符，方便你搜索替换。

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

---

## 生成时落实的规范

### Word（本科毕业论文）
- A4；页边距四周 2.5cm + 装订线 1cm；页眉 1.5cm / 页脚 1.75cm
- 正文：宋体 / Times New Roman、小四(12pt)、固定行距 20 磅、首行缩进 2 字符、两端对齐
- 标题 1/2/3：黑体 三号(16)/四号(14)/13pt，套用内置样式 → **可自动生成目录**
- 摘要（小二加粗居中）、关键词、英文摘要、目录域（打开后按 **F9** 更新）、页码、参考文献(GB/T 7714)

### PPT（答辩）
- 16:9；字号层级 封面40 / 分节32 / 页标题30 / 正文24；微软雅黑
- 结构：封面 → 目录 → 研究背景 → 研究方法 → 研究成果 → 结论展望 → 致谢
- 白底 + 主色标题条；左对齐；每页要点 ≤ 6；遵循 10/20/30 原则

规范来源见 `config/format_spec.py` 中的 `SPEC_SOURCES`。
**各校要求不同，请以本校教务处/学院官方模板为准。**

---

## 自定义

改 `config/format_spec.py` 里的 `WORD_SPEC` / `PPT_SPEC` 常量即可全局调整字号、
页边距、行距、配色、PPT 结构段，无需改动生成逻辑。
