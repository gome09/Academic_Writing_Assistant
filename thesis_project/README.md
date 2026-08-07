# 论文 Word 草案 + 答辩 PPT 草案 生成器

读取你给出的 **Word / PDF / TXT(.txt·.text) / Markdown(.md·.markdown) / JSON / Excel(xlsx·csv) / 图片** 源文件并整理成结构化内容。
普通 `draft` 模式默认生成 `论文草案.docx` 与 `答辩PPT草案.pptx`；“题目 + 参考资料”的
`refs` 模式只生成 `论文草案.docx`。

> 定位：**草案**。产物已套好格式与结构骨架，正文与 `<请填写>` 占位符需人工润色。

> 当前架构、数据流与维护边界见 [架构说明](../docs/ARCHITECTURE.md)。


---

## 目录结构

```
thesis_project/
├── run.bat                # ★ 一键启动（Windows 双击即可）
├── config/
│   ├── format_spec.py     # 格式规范"硬标准"：字号/页边距/行距/标题层级/PPT结构
│   ├── template.py        # 外部 YAML 格式模板的加载与校验（--format-template）
│   └── format_template.example.yml  # YAML 格式模板示例
├── src/
│   ├── readers.py         # 多格式读取器 -> 统一中间结构
│   ├── organizer.py       # 内容整理：重建章节树、映射 PPT 结构、提炼要点
│   ├── docx_builder.py    # 按规范生成 .docx
│   ├── pptx_builder.py    # 按规范生成 .pptx
│   ├── synthesizer.py     # refs 模式：题目+文献 -> 综述/大纲/写作要点
│   ├── references.py      # 参考文献元数据抽取 + 多样式著录(GB/T 7714/APA/MLA/Chicago) + 引用校验
│   ├── llm_enhancer.py    # LLM 增强（可选，失败自动回退规则结果）
│   ├── llm_vision.py      # 图片理解（可选，LLM_VISION_MODEL）
│   ├── logging_setup.py   # 文件日志（output/运行日志.log，当前只记 PPT 告警）
│   ├── runtime_report.py  # 运行报告（output/运行报告.json）
│   ├── postprocess.py     # Word 域刷新 / PDF 导出（可选 win32com）
│   └── main.py            # 主管道（命令行入口）
├── tests/                 # pytest 用例
├── sample_input/          # 示例源文件（可直接试跑）
├── input/                 # 把你的源文件放这里（不入版本库，克隆后需自建）
├── output/                # 生成的草案输出到这里（不入版本库）
├── pyproject.toml         # 打包与 Ruff 配置
├── requirements.txt / .lock          # 核心依赖
├── requirements-llm.txt / .lock      # 可选：LLM 依赖
└── requirements-office.txt           # 可选：Word COM 依赖（无 .lock）
```

---

## 快速开始

### 方式一：一键启动（Windows，推荐）

双击 **`run.bat`** 即可。它会自动：
1. 定位 Python（优先 `py`，其次 `python`）；
2. 首次运行创建隔离虚拟环境 `.venv` 并在其中执行；
3. 缺核心依赖时自动 `pip install -r requirements.lock`；若设置了
   `LLM_API_KEY`，再自动补装 `requirements-llm.lock`；
4. `input\` 有文件就用 `input\`，为空则用 `sample_input\` 演示；
5. 生成完成后自动打开 `output\` 目录。

### 方式二：命令行

```bash
# 1) 先用示例试跑（会在 output/ 生成两份草案）
python src/main.py --input sample_input

# 2) 放入自己的文件后正式生成
#    把 Word/PDF/TXT/Markdown/JSON/Excel/图片 放进 input/，然后：
python src/main.py

# 其它用法
python src/main.py --input a.pdf b.md 某目录   # 指定若干文件/目录
python src/main.py --only word                 # draft 模式只生成 Word
python src/main.py --only ppt                  # draft 模式只生成 PPT
python src/main.py --output 某输出目录          # 自定义输出目录
python src/main.py --mode refs                 # 强制参考资料模式（draft 强制普通草案模式）
python src/main.py --llm                       # 用 LLM 增强草稿质量（可选）
python src/main.py --refresh-fields            # 生成后用本机 Word 刷新目录/页码域
python src/main.py --pdf                       # 生成 Word 时刷新域并导出 PDF（隐含 --refresh-fields）
python src/main.py --dry-run                   # 检查可读性/模式/外发清单并写报告，不调用外部服务或生成草案
python src/main.py --llm --yes                 # 非交互环境确认 LLM 外发
python src/main.py --format-template school.yml # 使用外部 YAML 格式模板
python src/main.py --lookup-metadata           # refs 模式显式查询 Crossref
python src/main.py --report 报告.json           # 自定义运行报告输出路径
```

> **Windows 提示**：请用 `python`（不要用 `python3`，它可能是应用商店占位程序）。
> 脚本已自动把控制台切到 UTF-8，中文不会乱码。
> 输出文件被 Word/WPS 打开占用时，会自动改存为 `论文草案(2).docx` 等（最多尝试 5 次）。

依赖：`python-docx`、`python-pptx`、`pdfplumber`、`openpyxl`、`PyYAML`；
`openai` 仅在 LLM 模式需要，`pywin32` 仅在 Word 域刷新/PDF 导出时需要。
```bash
pip install -r requirements.lock
pip install -r requirements-llm.lock       # 可选：LLM
pip install -r requirements-office.txt     # 可选：Word COM
```

---

## 输入文件怎么写效果最好

整理器会**识别标题层级**来重建章节：

- **Markdown**：用 `#`/`##` 作章/节标题，YAML frontmatter 写 `title/author/keywords`（见 `sample_input/thesis_draft.md`）。
- **Word (.docx)**：正文里用「标题 1/标题 2」样式的段落会被识别为章/节。
- **JSON**：支持 `{title, content, children:[...]}` 递归结构；也支持任意数据（会平铺为段落）。
- **PDF**：按页面文本行重组段落并识别编号标题；扫描版 PDF 没有文本层时会明确报错。**PDF 内嵌图片一律不导入**（只在控制台提示），需要配图请另存为图片文件放进 `input/`。
- **TXT（.txt / .text）**：按空行分段读取；无显式标题时会套用标准章节骨架并保持原文顺序。
- **Excel（.xlsx）**：每个非空工作表分别成为一个表格块。**CSV（.csv）**：整份文件只产出一个表格块。
- **图片（.png / .jpg / .jpeg / .bmp / .webp）**：读取为图片内容；可随 Word/PPT 草案一起排版。

识别不到的字段（题目、作者、英文摘要等）会留 `<请填写>` 占位符，方便你搜索替换。

---

## LLM 增强（可选）

设置环境变量后加 `--llm` 即可启用，能显著提升草稿质量：

| 环境变量 | 必填 | 说明 |
|---|---|---|
| `LLM_API_KEY` | 是 | API 密钥；未设置时 `draft` 模式自动跳过增强，`refs` 模式直接报错退出 |
| `LLM_BASE_URL` | 否 | OpenAI 兼容端点，如 DeepSeek/通义/Kimi/本地 Ollama |
| `LLM_MODEL` | 否 | 模型名，默认 `gpt-4o-mini` |
| `LLM_TIMEOUT` | 否 | 单步超时秒数，默认 60 |
| `LLM_VISION_MODEL` | 否 | 图片理解模型（如 qwen-vl-plus），refs 模式为截图生成图题与摘要 |
| `THESIS_LLM_CONSENT` | 否 | 设为 `1` 等效 `--yes`，非交互环境自动同意 LLM 外发 |
| `PYTEST_CURRENT_TEST` | — | 由 pytest 自动注入；**存在时会跳过外发确认**。仅供测试，请勿在生产环境手工设置 |

普通 `draft` 模式的增强内容：元信息抽取（题目/作者/摘要/关键词）、英文摘要与关键词翻译、
PPT 要点语义提炼、章节到 PPT 分区的语义分类、无标题文档的语义分章、
每页演讲备注（写入 PPT 备注区，带 AI 标记）。

原则：在普通 `draft` 模式中，LLM 只做整理/提炼/翻译/分类，不扩写正文；
摘要、英文摘要和演讲备注等生成文本会带 `<AI生成，请核对>` 标记，题目、作者、
关键词、PPT 要点和分类结果不会统一附加该标记。任一步失败会回退纯规则结果，
主流程不受影响。`refs` 模式会生成带核对标记的综述初稿，具体边界见下一节。

启用 LLM 时程序会先显示端点主机、文件数量、文本规模和图片数量。交互终端
需确认后才发送；CI/脚本等非交互环境必须显式传 `--yes`（或设
`THESIS_LLM_CONSENT=1`；pytest 运行时会因 `PYTEST_CURRENT_TEST` 自动跳过确认）。`--dry-run` 永不
调用外部服务。原始资料会作为不可信输入包裹，LLM 输出仍会经过本地结构、
长度和引用编号校验。`dry-run`、成功完成的运行以及进入 `draft` 构建后失败的运行会把详情写入
`output/运行报告.json`；无可读输入、拒绝外发、`refs` 前置检查失败或 `refs` 的 Word
写盘失败会提前退出，此时不保证生成运行报告。`output/运行日志.log` 当前只记录
**PPT 生成告警**，其余阶段仅输出到控制台——因此 `--only word` 和 `refs` 运行会得到空日志文件。

> ☆ **如何定位需重点复核的生成文本**：检索 `<AI生成，请核对>` 可找到摘要、英文摘要、
> 演讲备注和 `refs` 综述等带标记内容；题目、作者、关键词、PPT 要点和分类结果仍需结合
> `--llm` 运行报告与原始资料人工复核。


```powershell
# PowerShell 示例（DeepSeek）
$env:LLM_API_KEY  = "sk-..."
$env:LLM_BASE_URL = "https://api.deepseek.com"
$env:LLM_MODEL    = "deepseek-chat"
python src/main.py --llm
```

---

## 参考资料模式（题目 + 文献 → 论文初稿骨架）

`input\` 里放一个**题目文件**（`topic.md`、`topic.txt`、`题目.md` 或 `题目.txt`，
写论文题目、研究内容简述、拟采用方法）+ 参考文献（PDF/Word/md/txt/json）、数据（xlsx/csv）、
截图（png/jpg），双击 `run.bat` 即自动进入本模式：

- LLM 生成：文献综述章节（带 [n] 引用）、全文大纲、核心章节【写作要点】
  与素材摘录；参考文献的**著录格式**由本地代码确定性生成 GB/T 7714，不交给 LLM 排版。
  元数据来源有回退链：优先本地正则抽取（及 Crossref 回填），本地缺失时回退到摘要卡里
  由 LLM 抽取的题名/作者/年份，都没有才保留 `«请补全…»` 形式的占位符
  （`«请补全作者»`、`«请补全题名»`、`«请补全期刊»`、`«请补全著录信息»`）；
- LLM 不生成：研究设计/实现/实验等核心章节正文——留 `<请填写>` 由你完成；
- 只生成 `论文草案.docx`，不生成 PPT（补完正文后用普通模式再生成）；
- xlsx/csv 自动插表、截图自动插图；设置 `LLM_VISION_MODEL`（如 qwen-vl-plus）
  后截图还会获得图题建议与内容摘要；
- **必须**设置 `LLM_API_KEY`（见上节），未设置会报错退出；
- 默认不查询外部文献数据库；只有传 `--lookup-metadata` 才会访问 Crossref，
  查询结果缓存于 `.cache/reference_metadata.json`；
- `--mode refs|draft` 可强制指定模式，覆盖自动检测。

`--dry-run` 只检查文件可读性、识别出的模式和外发清单，不校验 `refs` 正式运行所需的
题目文件、参考资料与 API 密钥。`refs` 中单步失败会明确告警并记录降级步骤：摘要卡退为
原文片段卡，大纲退为默认骨架，综述退为素材摘录与占位符，写作要点或视觉理解失败时保留
占位内容；参考文献仍由本地元数据确定性格式化，不交给 LLM 猜测。

> ⚠ 学术诚信：综述正文全部带 `<AI生成，请核对>` 标记，属于初稿素材，
> 务必逐条核对原文献后改写为自己的表述再使用。

---

## 生成时落实的规范

### Word（本科毕业论文）
- A4；页边距四周 2.5cm + 装订线 1cm；页眉 1.5cm / 页脚 1.75cm
- 正文：宋体 / Times New Roman、小四(12pt)、固定行距 20 磅、首行缩进 2 字符、两端对齐
- 标题 1/2/3：黑体 三号(16)/四号(14)/13pt，套用内置样式 → **可自动生成目录**
- 摘要（小二加粗居中）、关键词、英文摘要、目录域（打开文档后按提示更新域，或按 **F9**）、页码
- 参考文献：普通草案模式**保留源条目原样**并编号，不做格式转换；源文件里识别不到任何参考文献时会插入一条示例兜底条目（含 `<请替换为真实文献>`）。参考资料模式按 `reference.standard`（默认 GB/T 7714，可选 APA/MLA/Chicago）确定性格式化，缺失字段保留占位符；文献类型依据字段启发式识别（J/M/C/D）
- 附录：标题以「附录」开头的章会被移到参考文献之后，并重编号为「附录A」「附录B」…

### PPT（答辩）
- 16:9；字号层级 封面40 / 分节32 / 页标题30 / 正文24；微软雅黑
- 结构：封面 → 目录 → 研究背景 → 研究方法 → 研究成果 → 结论展望 → 致谢
- 白底 + 主色标题条；左对齐；每页要点 ≤ 6；总页数与各分区页数超出配置范围时给出告警
- 表格上 PPT 时会被截断为**最多 8 行 6 列**并给出告警；超长表格会先由整理器拆成多页

规范来源见 `config/format_spec.py` 中的 `SPEC_SOURCES`。
**各校要求不同，请以本校教务处/学院官方模板为准。**

---

## 自定义

改 `config/format_spec.py` 里的 `WORD_SPEC` / `PPT_SPEC` 常量可调整已接入生成器的字号、
页边距、行距、配色等字段；也可用 YAML 深度覆盖：

```yaml
word:
  page:
    orientation: landscape
ppt:
  sizes:
    body_pt: 22
```

运行：`python src/main.py --format-template format.yml`。未知字段和类型错误会在
生成前直接报错；尺寸与范围仍应按示例和本校规范填写有效值。
`format_spec.py` 中标注 `# 暂未落实` 的字段（包括 `10/20/30` 原则、`figure.caption_position`、
`figure/table.number_by_chapter`、`table.style` 等）仅作文档参考，修改不会影响产物。
与之相对，`table.caption_position`、`figure/table.caption_align` 以及页面方向、
PPT 内容比例、媒体分页等字段是真实生效的，可通过 YAML 模板或配置文件调整。
`reference.standard` 现支持 `GB/T 7714`（默认）、`APA`、`MLA`、`Chicago` 四种著录样式，
切换后参考资料模式的参考文献条目随之变化（普通草案模式保留源条目原样）。
`PPT_SPEC["structure"]` 现驱动 PPT 分区生成（T2-1）——内容段的标题与顺序取自配置，
可通过 YAML 自定义增减分段（如去掉「结论与展望」或新增「未来工作」），默认仍为 7 段。
`PPT_SPEC["theme"]["preset"]` 支持 5 套内置主题（T2-2）：`academic_blue`（默认深蓝）、
`minimal_gray`（简约灰）、`campus_red`（院校红）、`dark`（深色）、`forest_green`（森林绿），
切换后 PPT 配色随之变化；也可单独覆盖 `primary_rgb`/`accent_rgb` 等字段与预设共存。
`PPT_SPEC["structure"]` 的页数范围仍只用于校验与告警。
