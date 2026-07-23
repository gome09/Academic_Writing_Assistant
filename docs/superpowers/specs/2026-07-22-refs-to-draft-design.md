# 参考资料 → 论文初稿（研究写作辅助）设计文档

日期：2026-07-22
状态：已与用户逐节确认

> **当前实现更新（2026-07-23）**：参考资料模式已实现并在原设计上增加了有序章节
> `blocks`、附录保留、确定性参考文献格式化、引用校验、运行报告、`--dry-run`/
> `--yes` 外发确认、Crossref 显式查询和 YAML 格式模板。本文保留为原始设计记录；
> 与当前实现不一致之处以 [README](../../../thesis_project/README.md) 和源码为准。

---

## 1. 背景与目标

现有 thesis_project 是"草稿排版器"：输入自己写的论文草稿，输出套好格式的
`论文草案.docx` + `答辩PPT草案.pptx`。当 `input\` 里放的是**参考文献、领域文章、
Excel 数据、截图**时，现有管道会把多篇文献机械拼贴成"论文"（第一篇文献的标题被
当成论文题目），Excel 与独立图片则被静默跳过。

本设计新增**参考资料模式**：读取参考资料后，由 LLM 生成一份"带综述和素材的
半成品论文骨架"，帮助作者启动写作。

**定位（用户已确认）：研究写作辅助，不是论文代写。**

- LLM 生成：文献综述章节（相关理论与技术 / 国内外研究现状）的真实正文（带引用
  编号）、全文大纲、核心研究章节的【写作要点】与素材摘录、GB/T 7714 参考文献表。
- LLM 不生成：研究设计 / 系统实现 / 实验结果等核心章节的正文——这些留
  `<请填写>` 占位符，由作者本人完成。
- 所有 AI 生成正文沿用现有 `<AI生成，请核对>` 标记；综述章开头插入醒目提示段：
  "本章由 AI 基于所给文献生成，请逐条核对原文后改写为自己的表述"。

## 2. 用户可见行为

### 2.1 触发方式（自动检测 + 显式覆盖）

- `input\` 下存在约定名题目文件（`topic.md` / `topic.txt` / `题目.md` /
  `题目.txt`，不区分大小写）→ 自动进入参考资料模式；否则走原有排版模式。
- `--mode refs|draft` 命令行参数可显式覆盖自动检测。
- 参考资料模式**必须**有 LLM：未设置 `LLM_API_KEY` 时直接报错退出并打印配置
  示例，绝不静默产出拼贴结果。topic 文件之外没有任何参考资料文件时同样报错。
- 双击 `run.bat` 即可用（分流逻辑在 main.py 内部，bat 无需感知模式）。

### 2.2 题目文件内容

自由格式的 Markdown/文本：论文题目（首个一级标题或首行）、研究内容简述、
拟采用方法等。全文作为背景喂给 LLM；题目行同时写入 thesis["title"]。
frontmatter 里的 `author` 写入 thesis["author"]（无则占位符）。

### 2.3 新支持的输入格式

| 格式 | 处理方式 |
|---|---|
| `.xlsx` | openpyxl 读取；每个非空工作表 → 一个表格块（表题占位符），插入 Word 草案；同时文本化作为 LLM 综合素材 |
| `.csv` | csv 模块读取（复用现有 `_read_text` 编码回退）→ 单表格块，同上 |
| `.png/.jpg/.jpeg/.bmp/.webp` | 作为插图导入草案（图题占位符）；若设置 `LLM_VISION_MODEL` 且端点支持多模态，则额外生成图题建议与内容摘要，摘要并入综合素材；未设置则仅插图 |

两种模式共享读取器：原排版模式也因此获得 xlsx/csv/截图导入能力（纯规则部分）。

### 2.4 产出

参考资料模式**只生成 `论文草案.docx`，不生成 PPT**（此阶段没有研究成果，
答辩 PPT 无意义；作者补完正文后用原模式再生成 PPT）。

## 3. 架构

方案（用户已确认）：**扩展现有管道 + 新增 synthesizer 模块**，与 organizer
平级；thesis 保持旧字段兼容，同时章节新增规范化 `blocks`，Word/PPT 构建器按
兼容适配层消费。

```
thesis_project/
├── src/
│   ├── readers.py        # 扩展：+read_xlsx +read_csv +read_image
│   ├── llm_vision.py      # 新建：仅承载视觉调用（_chat_vision / describe_image）；
│   │                     #   文本调用继续经 llm_enhancer._chat（保持既有打桩链）
│   ├── llm_enhancer.py   # 扩展：抽出 _parse_json 供 llm_vision 复用；对外行为与打桩点不变
│   ├── synthesizer.py    # 新建：参考资料模式核心（organizer 的平级替代品）
│   └── main.py           # 改动：topic 检测分流；--mode 参数；refs 模式禁 PPT
└── config/format_spec.py # 扩展：+REFS_SPEC（默认大纲骨架、综述章名、提示语）
```

### 3.1 数据流

```
input\（topic 文件 + PDF/docx/md/txt/json/xlsx/csv/图片）
  → readers 读成统一 Document（xlsx→table 块；截图→image 块）
  → main.py 检测 topic → synthesizer.synthesize(topic_doc, docs)
      ① 逐篇文献 → "摘要卡"（LLM）：{标题, 作者, 年份, 主题, 方法, 结论,
         可引用观点[], 来源文件名}
      ② 题目 + 全部摘要卡 → 论文大纲（LLM），并标注各章关联的摘要卡
      ③ 综述章节撰写（LLM，逐章）：基于关联摘要卡生成正文，引用处标 [n]
      ④ 核心研究章节（LLM，批量）：生成【写作要点】+ 素材摘录 + <请填写>
      ⑤ 参考文献表（本地确定性 formatter）：本地元数据 → GB/T 7714 条目；缺失字段
         保留 `«请补全…»`，不让 LLM 猜测
      ⑥ 媒体挂载（纯规则）：xlsx 表格、截图按语义就近挂到大纲章节，
         无法判断时挂到素材附录章
  → thesis dict（与 organizer 输出同构）→ 现有 docx_builder.build
```

### 3.2 thesis 字段映射（refs 模式）

| 字段 | 来源 |
|---|---|
| title / author | topic 文件（缺失则占位符） |
| abstract / abstract_en / keywords* | 占位符（尚无研究成果，不编造摘要） |
| chapters | ② 的大纲；综述章有正文，核心章为写作要点+素材+占位 |
| references | ⑤ 的 GB/T 7714 条目 |
| auto_skeleton | False |

## 4. LLM 调用设计

| 步骤 | 次数 | 输入规模控制 |
|---|---|---|
| ① 摘要卡 | 每篇 1 次 `_chat_json` | 每篇截取前约 6000 字符 |
| ② 大纲 | 1 次 | 题目全文 + 各摘要卡标题/主题行 |
| ③ 综述 | 每综述章 1 次 | 仅该章关联的摘要卡 |
| ④ 写作要点 | 1 次批量 | 大纲 + 摘要卡主题 |
| ⑤ 参考文献 | 0 次（本地 formatter） | 本地元数据；可选 Crossref 查询 |
| 视觉理解 | 每张截图 1 次（仅当 `LLM_VISION_MODEL` 已设置） | 图片 base64 |

- 10 篇文献 + 5 张截图约 17 次调用；参考文献格式化不调用 LLM，temperature 沿用 0.2。
- 所有文本调用经 `llm_enhancer._chat`（既有网络出口，测试打桩点，synthesizer 复用
  `_chat_json`）；视觉调用经 `llm_vision._chat_vision`（第二个网络出口，同样可打桩）。
  JSON 解析统一由 `llm_enhancer._parse_json` 承担并被 `llm_vision` 复用。
- 新增环境变量：`LLM_VISION_MODEL`（选填；未设置则跳过视觉理解）。

## 5. 错误处理（分级降级）

- **入口硬校验**（报错退出，退出码 1）：refs 模式无 `LLM_API_KEY`；
  topic 之外无任何参考资料。
- **逐篇容错**：单篇摘要卡失败 → 该篇退化为首段截断文本卡；xlsx 损坏/加密、
  图片格式不支持 → 跳过并告警（沿用 `read_dir` 惯例）。
- **步骤降级**：大纲失败 → `DEFAULT_CHAPTERS` 标准骨架；单个综述章失败 →
  该章留素材摘录+占位符；本地参考文献字段缺失 → 输出 `«请补全…»`。
- **视觉失败** → 仅插图，不阻塞。
- 每次降级打印 `[LLM告警]`，结束时汇总（"N 步降级，请检查"）。全部降级仍
  产出结构完整的 Word（骨架+素材摘录），退出码 0。

## 6. 测试策略

沿用现有惯例（pytest；`_chat` / `chat_vision` 打桩；`from src import xxx`）：

- 读取器：xlsx（多 sheet/空 sheet/合并单元格）、csv（编码回退）、image。
- synthesizer：摘要卡容错、大纲降级、综述引用编号与 references 对齐、
  写作要点带 AI 标记、媒体挂载、无 Key 入口报错。
- main：topic 检测分流、`--mode` 覆盖、refs 模式不产 PPT。
- 端到端：打桩 LLM 跑全流程，校验 docx 章节结构与引用编号。
- 回归：现有全量测试必须通过（`_parse_json` 抽取、`llm_vision` 新增后 llm_enhancer 行为不变）。

## 7. 依赖变更

- 新增必装：`openpyxl`（纯 Python，轻量）。
- `run.bat` 的依赖检查行加入 `openpyxl`。
- 不引入本地 OCR 库（视觉理解走 LLM 多模态端点，可选）。

## 8. 明确不做（YAGNI）

- 不写核心研究章节正文（学术诚信边界，用户已确认）。
- refs 模式不生成 PPT（含开题 PPT）。
- 不做本地 OCR。
- 默认不做期刊元数据联网检索；传 `--lookup-metadata` 时才允许 Crossref 查询，
  查询结果缓存到 `.cache/reference_metadata.json`。

## 9. 当前数据模型补充

- Document 仍以 `blocks` 保持读取顺序；章节节点新增规范化 `blocks`，每块携带
  `kind/source/source_index` 和文本、表格或图片载荷。
- `paras/tables/images` 作为兼容视图保留，旧测试夹具和手工构造 thesis 仍可被构建器消费。
- 附录章节使用 `section_role: appendix`，在参考文献后输出，不计入正文章号。
