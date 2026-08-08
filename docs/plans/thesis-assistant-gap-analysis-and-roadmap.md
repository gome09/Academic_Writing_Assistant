# 论文草案生成器 · 市场对标差距分析与优化实施路线图

> 编制日期：2026-08-07
> 依据：代码库逐文件审计（带行号证据） + 三类对标产品联网调研（论文写作类 / PPT 生成类 / AI 通用写作学术能力，2025-2026 公开资料）
> 范围：`thesis_project/` 业务代码、配置与测试
> 性质：**实施路线图**（分阶段任务拆解），非运行时事实来源；文档与代码冲突时以代码与测试为准

---

## 0. 结论速览

本项目在「**本地优先 / 隐私优先 / 格式确定性 / 学术诚信约束**」四点上构成真实差异化护城河，市面竞品普遍不具备这种"可离线、可自托管 LLM、参考文献不让 AI 编造、生成文本带核对标记"的组合。但在以下七个方向**全面落后于市场**，且其中若干是结构性缺陷而非功能补丁：

1. **文献检索能力近乎为零**（只有 Crossref 元数据回填，且默认关闭；无任何"找文献"能力）
2. **AI 用法过保守以至于落后**（draft 不扩写正文是诚信设计，但连多轮润色、校对、长上下文整合都没有）
3. **PPT 能力停留在"固定骨架"**（结构硬编码、无模板/主题、无文档转 PPT 智能、无配图/图表、无 PDF 直出）
4. **参考文献仅 GB/T 7714 且类型检测失效**（几乎全落 `[J]`）
5. **无 OCR、PDF 内嵌图片被丢弃**（学术场景高频痛点）
6. **工程地基有两条 P0 隐患**（`PYTEST_CURRENT_TEST` 可绕过外发确认；`format_spec` 全局可变 dict 被 YAML 原地改写）
7. **形态单一**（纯本地 CLI，无增量、无流式、无 Web 形态，阻碍协作与变现）

**核心策略**：守住四条红线（见 §3），分 6 个阶段推进——先清地基（阶段 0），再补参考文献与 PPT 两块高价值短板（阶段 1-2），随后在"受限开放"原则下深化 AI（阶段 3），扩展输入与插件化（阶段 4），最后视需要上形态（阶段 5）。

---

## 1. 项目能力画像（代码事实）

| 维度 | 现状 | 代码证据 |
|---|---|---|
| 入口 | 单一 CLI，`argparse`，无 Web/库调用入口 | `src/main.py:190-308` |
| 输入格式 | Word/PDF/TXT/MD/JSON/Excel/图片，无 OCR | `src/readers.py:449-464` `_READERS` |
| 模式 | draft（Word+PPT）/ refs（题目+文献→Word） | `src/main.py:125,265` |
| LLM | OpenAI 兼容，可选，需确认；draft 只整理/翻译/分类不扩写正文；refs 生成综述+大纲+写作要点 | `src/llm_enhancer.py`, `src/synthesizer.py` |
| 参考文献 | 本地正则→Crossref(默认关)→LLM 回退；GB/T 7714 确定性；类型默认 `[J]` | `src/references.py:57,69-96,116-146` |
| Word 格式 | 字体/页边距/目录域/页码/附录重编号，按本科论文规范 | `src/docx_builder.py` |
| PPT | 固定 7 段，16:9，硬编码坐标，config 只校验不驱动 | `src/organizer.py:411-482`, `src/pptx_builder.py` |
| 配置 | `format_spec.py` + YAML 覆盖；**14 个字段不被构建器消费** | `config/format_spec.py`, `config/template.py:62-67` |
| 隐私 | Crossref 默认关，LLM 需确认，dry-run 不触网 | `src/main.py:149-175,247-256` |
| 测试 | 53 个测试文件，多数有内容级断言；`lookup_crossref` 无单测 | `tests/` |

---

## 2. 差距矩阵（代码现状 × 市场标杆）

差距评级：🔴 显著落后 ｜ 🟡 部分落后 ｜ 🟢 持平/领先

### 2.1 文献检索与输入

| 能力 | 我方 | 市场标杆 | 评级 |
|---|---|---|---|
| 找文献（语义检索/相关推荐） | 无 | 橙篇(百度学术2亿)、Consensus(2亿篇)、Elicit(1.38亿+54.5万临床试验)、ResearchRabbit(2.7亿) | 🔴 |
| 文献元数据回填 | Crossref，默认关 | 橙篇自动生成引用格式；NoteExpress/知网研学引文格式智能生成 | 🟡 |
| 扫描件 OCR | 无，直接报错 | 主流工具普遍支持 | 🔴 |
| PDF 内嵌图片 | 一律丢弃 | 多数支持提取 | 🔴 |
| 多文件批量 | 支持（最多 N 文件） | 橙篇 100 文件/次、NotebookLM 50 源(Pro 300) | 🟢 |
| 格式扩展难度 | 改源码注册字典 | 主流为插件化 | 🟡 |

### 2.2 AI 增强

| 能力 | 我方 | 市场标杆 | 评级 |
|---|---|---|---|
| 元信息抽取/翻译/分类 | 有 | 普遍有 | 🟢 |
| 多轮润色/改写 | 无 | 秘塔(保守/标准/强力三档降重)、WPS 实时润色 | 🔴 |
| 校对（错别字/语法/术语） | 无 | 秘塔(纠错准确率98%)、WPS 智能校对(30+项) | 🔴 |
| 长上下文整篇处理 | 无（散落截断 4000/6000/500） | Kimi(200万字)、NotebookLM(50源深度引用)、DeepSeek(128K) | 🔴 |
| 源头引用/RAG | 无 | NotebookLM(深链到 PDF 段落)、Consensus(可溯源) | 🔴 |
| AIGC 检测/降重 | 无（仅标注） | PaperYY(双引擎降AIGC率)、笔杆(AIGC优化器) | 🟡 |
| 流式输出 | 无 | 普遍流式 | 🔴 |

### 2.3 参考文献

| 能力 | 我方 | 市场标杆 | 评级 |
|---|---|---|---|
| 著录样式 | 仅 GB/T 7714 | 笔灵(APA/MLA)、百度学术(多格式)、NoteExpress | 🔴 |
| 类型识别 | 默认 `[J]`，LLM 未返回 type | 主流区分 J/M/C/D 等 | 🔴 |
| 引用校验 | 有（越界/无关） | Scite(Smart Citation 上下文) | 🟢(校验) / 🔴(上下文) |
| 确定性（不让 AI 编造） | ✅ 强项 | 多数竞品 AI 参与排版， ours 更可信 | 🟢 |

### 2.4 PPT 生成

| 能力 | 我方 | 市场标杆 | 评级 |
|---|---|---|---|
| 生成入口 | 仅源文件→固定7段 | Gamma/AiPPT/WPS：主题/文档/链接/大纲/Excel 多入口 | 🔴 |
| 文档转 PPT | 无（结构固定） | Beautiful.ai/讯飞/AiPPT 核心能力 | 🔴 |
| 模板/主题 | 极有限（config 常量） | Slidesgo 1万+、AiPPT 20万+、讯飞 5000+ | 🔴 |
| AI 配图/图表 | 无 | Gamma/讯飞/轻竹(Excel 可视化) | 🔴 |
| 演讲备注 | 有（弱） | 讯飞(写-练-演全链路)、轻竹(差异化卖点) | 🟡 |
| 导出格式 | 仅 .pptx | 普遍 PPTX+PDF+PNG+链接 | 🔴 |
| 联网检索补数据 | 无 | 讯飞/WPS 联网搜索+标注出处 | 🔴 |
| 数字人/视频 | 无 | 腾讯智影/讯飞/AiPPT 数字人讲解 | 🔴(战略可选) |

### 2.5 工程与地基

| 能力 | 我方 | 状态 | 评级 |
|---|---|---|---|
| 外发确认可被绕过 | `PYTEST_CURRENT_TEST` 环境变量 | `main.py:161` | 🔴 P0 |
| 配置全局可变 | YAML 原地改 `WORD_SPEC` | `template.py:64-67` | 🔴 P0(并发) |
| 日志 | 60+ 处 `print()`，仅 pptx 写文件 | 全局 | 🟡 |
| 全局可变状态 | `_degraded`/`LAST_WARNINGS` | `synthesizer.py:41`,`pptx_builder.py:36` | 🟡 |
| 死代码 | `_render_media` 别名、`read_dir` | `docx_builder.py:457`,`readers.py:495` | 🟢(清理) |
| 14 个惰性配置字段 | 不被消费，4 个无标注 | `format_spec.py` | 🟡 |
| 文档/代码契约 | 隐式 dict，无 schema | `readers.py:6-20` | 🟡 |
| 测试缺口 | crossref 无单测、template 类型/负数分支、confirm 交互 | `tests/` | 🟡 |

### 2.6 形态与体验

| 能力 | 我方 | 市场标杆 | 评级 |
|---|---|---|---|
| 形态 | 本地 CLI | 普遍 Web SaaS + 客户端 + 协作 | 🔴(战略) |
| 增量/缓存 | 无（每次全量） | 普遍增量 | 🟡 |
| 协作 | 无 | Beautiful.ai/博思 百人实时协作 | 🔴(战略) |
| 模板市场 | 无 | Slidesgo/AiPPT 模板市场 | 🔴(战略) |
| 商业模式 | 无 | Freemium+订阅+私有化 | — |

---

## 3. 红线（路线图不得破坏）

下列四条是项目差异化护城河，**任何阶段任务若与之冲突，以红线为准**：

1. **学术诚信**：draft 模式 LLM 不扩写正文；refs 综述/写作要点带 `<AI生成，请核对>`；新增任何 AI 生成文本须保留可检索标记。**不得**让 LLM 编造参考文献元数据。
2. **格式确定性**：GB/T 7714（及未来新增样式）著录由本地代码生成；新增样式不得交给 LLM 排版。
3. **隐私优先**：Crossref 及任何新增网络出口**默认关闭**，需显式启用；LLM 外发保留确认机制；`--dry-run` 不触网。
4. **本地优先**：核心管道可在无网络/自托管 LLM（Ollama）下运行；新增能力不得强制依赖云服务。

> 守红线 ≠ 不进化。红线约束的是"如何做"（可选、可降级、可离线、带标记），而非"是否做"。路线图在"受限开放"原则下推进 AI 深化。

---

## 4. 实施路线图（分阶段）

每阶段标注：**目标 / 任务（含文件·验收）/ 红线检查 / 依赖**。
任务编号 `T{阶段}-{序号}`。优先级 `P0/P1/P2`。

### 阶段 0 · 地基清理与 P0 修复（必做）

**目标**：消除两条 P0 隐患，统一日志，清理死代码，补关键测试。此阶段不新增功能，为后续扩展打地基。

| 编号 | 任务 | 改动文件 | 验收标准 | 优先级 |
|---|---|---|---|---|
| T0-1 | 修复 `PYTEST_CURRENT_TEST` 绕过：运行时检测当前是否真在 pytest 下（如 `sys.modules.get("pytest")` 且 `PYTEST_CURRENT_TEST` 由 pytest 注入），仅在确属测试运行时豁免 | `src/main.py:160-161` | 单测覆盖"生产环境误设该变量仍要求确认"；新增 `test_llm_consent.py` | P0 |
| T0-2 | 配置不可变化：`apply_template` 不再原地 `clear/update` 全局 `WORD_SPEC`，改为返回新 dict 并由 `reload_spec` 刷新缓存；或深拷贝后合并 | `config/template.py:54-67`, `src/organizer.py:24`, `src/pptx_builder.py:30-43`, `src/main.py:220-225` | 同进程多次 `apply_template` 不互相污染；`test_format_template.py` 移除 `finally` 还原 hack；并发场景单测 | P0 |
| T0-3 | 日志统一：各模块 `print()` 改为 `logging`，复用 `logging_setup`；关键事件（LLM 降级、读取失败、写盘重试、Crossref）落盘 | `src/main.py`(30处), `src/llm_enhancer.py`(13处), 其余模块 | `--only word` 与 refs 运行后日志非空；日志含降级步骤 | P1 |
| T0-4 | 全局可变状态上下文化：`_degraded` 改为 `synthesize()` 内局部列表并随返回值传出；`LAST_WARNINGS` 改为 build 返回值的一部分 | `src/synthesizer.py:41,370`, `src/pptx_builder.py:36` | 模块级无可变全局；现有测试通过 | P1 |
| T0-5 | 死代码清理：删除 `docx_builder._render_media` 别名；迁移 `test_e2e.py` 用 `read_dir_detailed` 后删除 `readers.read_dir` | `src/docx_builder.py:457-459`, `src/readers.py:495-498`, `tests/test_e2e.py` | 无引用残留；测试通过 | P2 |
| T0-6 | 补测试：`lookup_crossref`（缓存命中/miss/DOI 过滤/超时，monkeypatch urlopen）、`template._validate` 类型错误与负数分支、`_confirm_llm_transfer` y/N 交互分支 | `tests/` 新增 3 文件 | 分支覆盖；不触真实网络 | P1 |
| T0-7 | 惰性字段治理：给 4 个无标注惰性字段（`figure.caption_position`/`figure.number_by_chapter`/`table.number_by_chapter`/`table.style`）补 `# 暂未落实` 标注；或在阶段 2/4 落实后移除标注 | `config/format_spec.py:96,98,104,106` | 标注与代码消费状态一致；同步更新 README/ARCHITECTURE | P2 |

**红线检查**：T0-1/T0-2 不改变外发与配置语义，仅加固；T0-3 日志不得写入密钥（已有，保持）。

**依赖**：无。此阶段先行。

---

### 阶段 1 · 参考文献与文献检索增强（高价值）

**目标**：把参考文献从"仅 GB/T 7714 且类型失效"升级为"多样式 + 类型识别 + 可选语义检索"，补齐最大功能短板之一。

| 编号 | 任务 | 改动文件 | 验收标准 | 优先级 |
|---|---|---|---|---|
| T1-1 | 文献类型识别：`entry_from_card` 依据字段启发式（有出版社→M、有会议名→C、有学位授予单位→D）+ LLM 提示词要求返回 `type`；默认仍 `J` 但不再恒等 | `src/references.py:51-66`, `src/synthesizer.py:108-110` | 学位论文落 `[D]`、专著落 `[M]`；新增 `test_references_type.py` | P1 |
| T1-2 | 多样式支持：新增 `format_apa`/`format_mla`/`format_chicago`；`WORD_SPEC["reference"]["standard"]` 改为被消费字段，按其分发；移除"暂未落实"标注 | `src/references.py`, `config/format_spec.py:118`, `src/synthesizer.py:433` | YAML 可切换样式；样式间单测；GB/T 7714 行为不变 | P1 |
| T1-3 | Crossref 健壮性：`lookup_crossref` 包 `try/except`（超时/网络/JSON 错），指数退避 1 次；超时取自 `LLM_TIMEOUT` 或新常量 | `src/references.py:116-146` | 网络异常不中断流程，降级为跳过；T0-6 单测覆盖 | P1 |
| T1-4 | 可选语义检索（找文献）：新增 `--search-literature`（默认关），接入 OpenAlex/Semantic Scholar（免费 API，无密钥）按题目返回相关文献列表，可作为 refs 的"补充参考资料" | 新增 `src/literature_search.py`, `src/main.py`, `src/synthesizer.py` | 默认关；启用需确认外发；结果带来源 URL；dry-run 不触网；新增网络出口同步更新 README/ARCHITECTURE/外发清单/测试 | P2 |
| T1-5 | 引用上下文校验（可选）：综述引用校验升级，检查引用编号与文献列表一致且每个文献至少被引一次，给出告警 | `src/references.py:99-113`, `src/synthesizer.py:238-251` | 孤立文献告警；单测 | P2 |

**红线检查**：T1-4 是新增网络出口，**必须**默认关闭 + 外发确认 + 同步文档/测试（遵循 AGENTS.md「新增网络出口」条款）；任何样式格式化**不得**交给 LLM。

**依赖**：T1-1/T1-2 独立；T1-3 依赖 T0-6；T1-4 依赖 T0-2（配置不可变后易扩展）。

---

### 阶段 2 · PPT 能力跃升（差异化重点）

**目标**：把 PPT 从"固定骨架"升级为"结构可配 + 主题模板 + 文档转 PPT 智能 + 多导出"，这是市面差距最集中处，也是本地工具可差异化发力的点（不强求云端协作，但求生成质量与可配置性）。

| 编号 | 任务 | 改动文件 | 验收标准 | 优先级 |
|---|---|---|---|---|
| T2-1 | 结构配置驱动生成：`PPT_SPEC["structure"]` 不再只校验，改为驱动 `organizer._build_deck` 生成分区/顺序/标题；支持 YAML 自定义答辩结构（如增减"研究背景"） | `src/organizer.py:411-482`, `src/pptx_builder.py:300`, `config/format_spec.py:162-170` | YAML 改结构后产物随之变化；默认仍为 7 段；结构校验测试更新 | P1 |
| T2-2 | 主题包系统：抽象"主题"=配色+字体+布局坐标；内置 3-5 套主题（学术蓝/简约灰/院校红/深色），YAML 可选 `ppt.theme`；去除 `_slide_*` 硬编码坐标，改为读主题 | `src/pptx_builder.py:129-260`, `config/format_spec.py` | 切主题后配色/布局变化；`reload_spec` 刷新；主题单测 | P1 |
| T2-3 | 布局参数化：提取魔法数字（`pptx_builder.py:210` 表格 `rows[:8]`/列 `6`、`356` 行宽估算、`docx_builder.py:431` 图片 `Cm(14)`、`418` 表格字号 `10.5`）为 `PPT_SPEC`/`WORD_SPEC` 字段 | `src/pptx_builder.py`, `src/docx_builder.py`, `config/format_spec.py` | YAML 可调；现有截断告警行为可配；单测 | P2 |
| T2-4 | PPT→PDF 导出：复用 `postprocess` COM 机制扩展到 PPT，`--pdf` 时 PPT 也导出 PDF（仅 Windows+Office） | `src/postprocess.py`, `src/main.py` | PPT 生成 PDF；无 Office 时优雅降级告警；单测 monkeypatch | P2 |
| T2-5 | 演讲备注系统化：现有备注增强为"分页备注 + 预计时长（对标 `principle.talk_minutes` 落实）+ 关键过渡提示" | `src/llm_enhancer.py`, `src/pptx_builder.py`, `config/format_spec.py:176` | 备注含时长估算；带 AI 标记；单测 | P2 |
| T2-6 | AI 配图占位与图表（可选）：refs/draft 中表格数据可生成 `python-pptx` 原生图表（柱/折线/饼）替代纯文本表格；AI 配图以"占位框+图题建议"形式（不强制依赖文生图，保持本地可用） | `src/pptx_builder.py`, `src/organizer.py` | 表格可一键转图表；无 LLM 也能用；单测 | P2 |

**红线检查**：T2-1/T2-2 不破坏默认 7 段（向后兼容）；T2-6 AI 配图不强制云依赖，占位优先。

**依赖**：T2-1/T2-2 依赖 T0-2（配置刷新机制健壮）；T2-4 依赖现有 postprocess。

---

### 阶段 3 · AI 增强深化（受限开放）

**目标**：在守住"不扩写正文/带标记/可降级"前提下，补齐多轮润色、校对、长上下文整合等市面标配能力。**所有新增 AI 能力必须可选、可降级、带核对标记。**

| 编号 | 任务 | 改动文件 | 验收标准 | 优先级 |
|---|---|---|---|---|
| T3-1 | 多轮润色/改写（可选 `--polish`）：draft 模式对正文段落提供"保守/标准/强力"三档改写，输出带 `<AI润色，请核对>` 标记；失败回退原文 | `src/llm_enhancer.py`, `src/main.py` | 默认关；启用需确认；带标记；可降级；单测 monkeypatch | P1 |
| T3-2 | AI 校对（可选 `--proofread`）：错别字/语法/术语一致性检查，输出修订建议列表（不自动改原文，仅建议） | `src/llm_enhancer.py`, `src/main.py` | 仅建议不改写；带标记；可降级；单测 | P2 |
| T3-3 | 长上下文整合：refs 模式对接长上下文模型（Kimi 200 万字/DeepSeek 128K），统一散落的 `_doc_text` 截断阈值（4000/6000/500）为可配常量，支持整篇文献深度摘要 | `src/synthesizer.py:49,85,133`, `src/llm_enhancer.py` | 阈值可配；长上下文模型下综述更深；失败降级；单测 | P2 |
| T3-4 | 源头引用/RAG（可选，较重）：本地嵌入（如 `sentence-transformers`）对参考资料建索引，综述生成时检索相关片段并标注来源段落（对标 NotebookLM 深链）；默认关，需 `--rag` | 新增 `src/rag.py`, `src/synthesizer.py` | 默认关；本地嵌入可离线；引用可溯源；dry-run 不建索引；单测 | P2 |
| T3-5 | AIGC 率提示（非检测）：对 AI 生成文本统计占比并写入运行报告，提示用户可能触发 AIGC 检测（不做降重，仅告知） | `src/runtime_report.py`, `src/main.py` | 报告含 AI 文本占比；不做改写；单测 | P2 |
| T3-6 | LLM 网络韧性：`_client` `max_retries=1` 改为可配指数退避（2-3 次） | `src/llm_enhancer.py:39`, `src/llm_vision.py` | 弱网下降级减少；单测 | P2 |

**红线检查**：T3-1/T3-2/T3-4 默认关闭 + 外发确认 + 标记 + 可降级；T3-4 RAG 必须本地嵌入可离线，不得强制云；T3-5 不做降重（守住诚信）。

**依赖**：T3-3 依赖 T0-3（日志）；T3-4 较重，可延后。

---

### 阶段 4 · 输入扩展与插件化

**目标**：补 OCR 与 PDF 图片提取两个高频痛点，并把 readers/builders 做成可扩展插件。

| 编号 | 任务 | 改动文件 | 验收标准 | 优先级 |
|---|---|---|---|---|
| T4-1 | OCR（可选 `--ocr`）：扫描件 PDF 用 PaddleOCR/Tesseract 识别文本层；无 OCR 依赖时回退现有报错 | `src/readers.py:321-329` | 可选；无依赖时优雅报错；单测 monkeypatch | P1 |
| T4-2 | PDF 内嵌图片提取：提取 PDF 图片字节为 image 块（可选 `--extract-pdf-images`，默认关，因可能量大） | `src/readers.py:341-368` | 可选；图片进入 blocks；单测 | P2 |
| T4-3 | readers 插件化：`_READERS` 改为注册器（装饰器/entry_points），新增格式无需改核心 | `src/readers.py:449-464` | 第三方可注册读取器；现有测试通过 | P2 |
| T4-4 | builder 插件化：抽象 builder 基类（消费 thesis/deck dict 契约），`--only` choices 动态；为未来新输出目标（如 HTML/Markdown）铺路 | `src/main.py:195,272-294`, 新增 `src/builders/base.py` | 新 builder 可独立实现；契约文档化；单测 | P2 |
| T4-5 | Document/blocks schema：用 `dataclass` 或 `TypedDict` 显式定义 Block/Document/Node 契约，替换隐式 dict；保留向后兼容 | `src/readers.py:6-20`, `src/organizer.py:209-253` | 类型可检查；`iter_node_blocks` 兼容层可逐步移除；单测 | P2 |

**红线检查**：T4-1 OCR 依赖可选（不强装）；T4-3/T4-4 插件化不破坏现有 `_READERS`/`_DISPATCH` 行为。

**依赖**：T4-5 依赖 T0-2；T4-3/T4-4 可独立。

---

### 阶段 5 · 形态与体验（战略，可选）

**目标**：在本地优先前提下提升易用性与可协作性。此阶段投入大，视产品定位决定是否推进。

| 编号 | 任务 | 改动 | 验收 | 优先级 |
|---|---|---|---|---|
| T5-1 | 增量/缓存：基于输入文件哈希缓存中间 `Document`/thesis 结构，未变文件跳过重读重整 | `src/main.py`, 新增 `src/cache.py` | 二次运行加速；缓存失效正确；单测 | P2 |
| T5-2 | 流式输出：LLM 调用改流式，长文本生成实时回显 | `src/llm_enhancer.py`, `src/synthesizer.py` | 流式回显；可降级非流式；单测 | P2 |
| T5-3 | 本地 Web UI（可选）：FastAPI/Gradio 本地服务，复用核心管道，提供上传/参数/预览；仍本地运行，不强制云 | 新增 `src/web/` | 本地起服务；不破坏 CLI；dry-run 可用 | P3 |
| T5-4 | 模板市场（可选）：format YAML 模板与 PPT 主题包的社区共享格式规范 | 新增规范文档 + 校验 | 模板可导入校验；单测 | P3 |

**红线检查**：T5-3 Web UI 仍本地优先；不引入强制云依赖；外发确认机制在 Web 形态下等价保留。

**依赖**：T5-1/T5-2 依赖阶段 0；T5-3 依赖阶段 2/4 插件化。

---

## 5. 优先级与依赖总览

```
阶段0(地基) ──┬─→ 阶段1(参考文献)  ──→ T1-4(语义检索,较重)
              ├─→ 阶段2(PPT)       ──→ T2-6(图表,可选)
              ├─→ 阶段3(AI深化)    ──→ T3-4(RAG,较重)
              └─→ 阶段4(输入/插件) ──→ 阶段5(形态,战略)

P0：T0-1, T0-2                      ← 立即修
P1：T0-3/T0-4/T0-6, T1-1/T1-2/T1-3, T2-1/T2-2, T3-1, T4-1   ← 高价值
P2：其余                             ← 按需
P3：T5-3/T5-4                       ← 战略可选
```

**建议执行顺序**：阶段 0 全做完 → 阶段 1(T1-1/T1-2/T1-3) 与 阶段 2(T2-1/T2-2) 并行 → 阶段 3(T3-1) → 阶段 4(T4-1) → 其余按需。

---

## 6. 风险与红线守护清单

| 风险 | 守护措施 |
|---|---|
| AI 深化侵蚀学术诚信 | 所有 AI 生成文本带可检索标记；draft 不扩写正文原则保留；T3-1 润色单独标记；不做自动降重 |
| 新增网络出口扩大攻击面 | 一律默认关 + 外发确认 + dry-run 不触网 + 同步 README/ARCHITECTURE/外发清单/测试 |
| 配置可变化引发并发 bug | T0-2 根治；阶段 5 Web UI 前必须完成 |
| 插件化破坏现有契约 | T4-3/T4-4 保留旧注册字典行为；T4-5 schema 向后兼容 |
| OCR/RAG 引入重依赖 | 均设为可选依赖，无依赖时优雅降级，不强装 |
| 文档与代码漂移 | 每阶段交付同步更新 `thesis_project/README.md`、`docs/ARCHITECTURE.md`、相关测试 |

---

## 7. 验证要求（遵循 AGENTS.md）

每个任务交付前：
1. `cd thesis_project && python -m pytest` 全量通过；
2. `python -m ruff check src tests config` 通过（行宽 100，目标 3.11）；
3. 修复缺陷类任务先加复现测试再改实现；
4. LLM/Crossref/网络类测试用 monkeypatch，不触真实网络；
5. 新增网络出口同步外发确认、README、架构说明、测试；
6. 配置变更同步检查 `template.py`、构建器缓存常量、YAML 模板测试；
7. 不宣称"暂未落实"字段已生效；交付时报告修改文件、行为变化、已执行/未执行验证。

---

## 附录 A · 调研来源（精选）

**论文写作类**：橙篇(cp.baidu.com，百度系非腾讯)、秘塔写作猫(xiezuocat.com)、笔杆网、PaperYY、笔灵AI、知网研学、WPS AI 论文助手。
**PPT 生成类**：Gamma、Tome、Beautiful.ai、Slidesgo、讯飞智文、博思AIPPT、轻竹办公、闪击PPT、AiPPT.cn、WPS 灵犀、腾讯智影。
**AI 学术**：Consensus(2亿篇)、Elicit(1.38亿+54.5万临床试验)、Scite(Smart Citation)、NotebookLM(50源/深链)、ResearchRabbit、Kimi(200万字)、DeepSeek(128K/STEM)、ChatGPT。
> 注：橙篇归属经多源核实为百度系（用户原标"腾讯"有误）；部分产品价格/版本以官网实时为准，标注"待核实"项需上线前复核。

## 附录 B · 代码审计关键证据索引

- P0 安全：`src/main.py:160-161`
- P0 配置可变：`config/template.py:62-67`，缓存常量 `src/pptx_builder.py:30-43`、`src/organizer.py:24`
- 14 个惰性字段：`config/format_spec.py:52,59,66,90,91,96,98,104,106,118,131,140,158,176-178`
- 无 OCR：`src/readers.py:321-329`
- PDF 图片丢弃：`src/readers.py:341-368`
- PPT 结构硬编码：`src/organizer.py:411-482`
- 参考文献类型默认 J：`src/references.py:51-66`
- 死代码：`src/docx_builder.py:457-459`、`src/readers.py:495-498`
- 测试缺口：无 `test_references_crossref.py`、`test_format_template.py` 仅 2 测试

---

_本路线图为实施规划文档，非运行时事实来源。各任务落地时以代码与测试为最终准绳。_
