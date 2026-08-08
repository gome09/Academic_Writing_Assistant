# 项目架构说明

本文描述当前源码实际实现。用户用法以 `thesis_project/README.md` 为准；字段、参数和行为发生变化时，应同时更新对应测试与文档。

## 入口与模式

命令行入口是 `thesis_project/src/main.py`。程序先通过 `readers.py` 将输入统一读取为 `Document`，再按模式进入不同管道：

```text
输入文件
  -> readers.py：统一 Document / blocks（可选 cache.py：增量缓存未变文件）
  -> main.py：自动检测或 --mode 指定
       -> draft：organizer.py -> 可选 llm_enhancer.py -> docx_builder.py / pptx_builder.py
       -> refs： synthesizer.py -> references.py / 可选 llm_vision.py -> docx_builder.py
  -> 可选 postprocess.py：--refresh-fields / --pdf 时用本机 Word 刷新域并导出 PDF
  -> runtime_report.py：dry-run、成功完成及 draft 构建结果的运行报告
  -> logging_setup.py：output/运行日志.log（当前仅 pptx_builder.py 写入告警）
```

- `auto`：存在 `topic.md`、`topic.txt`、`题目.md` 或 `题目.txt` 时进入 `refs`，否则进入 `draft`。
- `draft`：生成 Word 与 PPT，可用 `--only` 限制产物；`--llm` 是可选增强。
- `refs`：要求题目文件、至少一份参考资料和 `LLM_API_KEY`，只生成 Word。
- `--dry-run`：检查文件可读性、模式与外发清单并写运行报告，不调用 LLM/Crossref，也不生成 Word/PPT；它不执行 `refs` 的题目、参考资料和 API 密钥前置校验，也不读写读取缓存。
- `--cache`/`--no-cache`（T5-1）：读取增量缓存默认开启。`read_file` 以 `(路径, 内容 SHA256, ocr, extract_images)` 为键缓存 `Document`，未变文件二次运行直接命中、跳过重读；含 image 块的文档不缓存（避免大字节存储）；缓存持久化为 `thesis_project/.cache/reads.pkl`（版本号失效旧缓存）。`--dry-run` 不读写缓存。

## 输入模型

`readers.py` 当前支持：

- 文本：`.txt`、`.text`、`.md`、`.markdown`、`.json`
- 文档：`.docx`、`.pdf`
- 表格：`.xlsx`、`.csv`
- 图片：`.png`、`.jpg`、`.jpeg`、`.bmp`、`.webp`

统一内容块的 `kind` 取值为 `heading`、`paragraph`、`list_item`、`table`、`code`、`image`。`Block` / `Document` 以 `TypedDict` 显式定义契约（T4-5），运行时仍为普通 dict：`kind`/`level`/`text` 必填，`rows`/`data`/`ext` 按需出现。Word 内嵌图片会被提取；PDF 依赖文本层，其内嵌图片默认不导入（只打印提示），加 `--extract-pdf-images` 可用 pypdf 提取为 image 块（默认关，可选依赖，未安装时优雅降级）；扫描版 PDF 默认报错，加 `--ocr` 可启用 OCR（T4-1，可选依赖）。`.xlsx` 的每个非空工作表各成一个表格块，`.csv` 整份文件只产出一个表格块。

读取器已插件化（T4-3）：`_READERS` 是扩展名→读取函数的注册表，各 `read_*` 在定义处用 `@register_reader` 装饰器注册，第三方可注册新格式而无需改核心分发。

## 普通草案管道

`organizer.py` 负责元信息抽取、章节树重建、特殊章节处理、无标题文档骨架回退和 PPT 分区映射。附录章在此被标记为 `section_role="appendix"`，由 `docx_builder.py` 移到参考文献之后并重编号为「附录A/B…」。`draft` 模式的 LLM 增强步骤只在显式传入 `--llm` 时执行，并在单步失败时保留规则结果；但 `llm_enhancer.py` 模块本身两种模式都会加载——`refs` 用它做前置可用性检查，并复用 `_chat_json` 作为唯一的 LLM 网络出口。

`docx_builder.py` 与 `pptx_builder.py` 消费整理后的结构数据。格式以 `config/format_spec.py` 为默认值，可通过 `--format-template` 加载 YAML 深度覆盖；`config/template.py` 负责深度合并并对未知字段、类型错误和负数报错。`PPT_SPEC["structure"]` 现驱动 `organizer.py` 生成分区/顺序/标题（T2-1），默认仍为 7 段，可用 YAML 自定义增减；其页数范围仍只用于校验与告警。PPT 侧另有硬性截断：表格超过 8 行会由 `organizer.py` 拆页，`pptx_builder.py` 最终只渲染前 8 行 6 列并告警。`PPT_SPEC["layout"]["chart_from_table"]`（默认关，T2-6）开启后，可图表化表格（首列为类别、至少一列全数值）渲染为 python-pptx 原生柱状图，不可图表化或失败时回退表格；`image_placeholder`（默认关，T2-6）开启后无媒体内容页插入配图占位框与基于标题的图题建议，不依赖文生图。

构建器已插件化（T4-4）：`src/builders/base.py` 定义 `Builder` 抽象基类与 `BUILDERS` 注册表，`docx_builder` / `pptx_builder` 通过适配器（`WordBuilder` / `PptBuilder`）注册为 `word` / `ppt`。`main.py` 的 `--only` choices 与 draft 生成循环取自注册表，第三方可注册新输出目标（如 HTML/Markdown）而无需改分发逻辑；refs 模式仍显式只用 `word` 构建器。

`format_spec.py` 中有一批字段当前不被任何构建器消费，修改它们不会影响产物：除标注 `暂未落实` 者外，还包括 `figure.caption_position`（图题恒在图下）、`figure.number_by_chapter` 与 `table.number_by_chapter`（编号恒为 `章-序`）、`table.style`（恒用三线表）。同层的 `table.caption_position` 与 `figure/table.caption_align` 则是真实生效的。

## 参考资料管道

`synthesizer.py` 是 `refs` 模式的整理入口：

1. 将参考资料逐篇整理为摘要卡；
2. 仅当传入 `--lookup-metadata` 时，逐张摘要卡查询 Crossref 并回填期刊、卷期、页码、DOI；
3. 生成或回退到章节大纲；
4. 为综述章节生成带引用标记的初稿；
5. 为核心章节生成写作要点，不生成核心研究正文；
6. 将表格与图片挂载到匹配章节，无法匹配时放入素材附录；
7. 由 `references.py` 确定性生成参考文献并校验引用。

只有 `xlsx`、`csv` 和图片三类文档会作为媒体素材参与第 6 步；`docx`、`pdf` 文献只被取用文本，其内嵌图片不会进入 `refs` 产物。

`refs` 的 AI 综述正文和视觉摘要保留 `<AI生成，请核对>` 标记；写作要点不统一附加该标记，仍须人工核对。摘要卡、大纲、综述、写作要点或视觉理解失败时会分别告警并降级，降级步骤写入成功运行的报告。

参考文献的**格式化**过程完全确定性，不交给 LLM。但**元数据来源**有回退链：优先用本地正则抽取的题名/作者/年份（以及 Crossref 回填结果），本地缺失时回退到摘要卡中由 LLM 抽取的同名字段，两者都没有才保留 `«请补全作者»`、`«请补全题名»` 等占位符。

Crossref 查询默认关闭，仅 `--lookup-metadata` 显式启用，并缓存到 `thesis_project/.cache/reference_metadata.json`。

## 外部依赖与网络边界

- 核心依赖：`requirements.lock`
- LLM 依赖：`requirements-llm.lock`
- Word 域刷新/PDF 导出：`requirements-office.txt`，仅 Windows + Microsoft Word
- LLM 网络请求：`llm_enhancer.py` 和 `llm_vision.py`
- Crossref 网络请求：`references.py`，仅显式启用元数据查询时发生
- 文献检索网络请求：`literature_search.py`（OpenAlex / Semantic Scholar），仅 `--search-literature` 显式启用时发生，免费无密钥

启用 LLM 时，交互环境需确认外发；非交互环境需 `--yes` 或 `THESIS_LLM_CONSENT=1`。`PYTEST_CURRENT_TEST` 存在且 pytest 已加载时跳过外发确认（T0-1 加固：生产环境误设不再绕过）。`--dry-run` 不调用任何外部服务。`--search-literature` 同样需外发确认。无可读输入、拒绝外发、`refs` 前置检查失败或 `refs` 的 Word 写盘失败会在运行报告写入前退出。

产物写盘时若目标文件被 Word/WPS 占用，`main.py` 的 `_build_with_retry` 会依次改名重试（`论文草案(2).docx` 等，最多 5 个候选），全部失败才判定构建失败。`draft` 运行报告额外记录逐文件读取失败列表 `read_errors` 与 `pptx_builder` 的告警。报告默认写入 `output/运行报告.json`，可用 `--report` 改路径。

## 事实来源与验证

维护时按以下优先级判断真实行为：

1. 可执行源码与配置；
2. 自动化测试；
3. `thesis_project/README.md` 与本文。

历史实施计划和设计草稿不作为运行时事实来源。项目回归命令：

```powershell
cd thesis_project
python -m pytest
python -m ruff check src tests config
python src/main.py --help
```

以上三条与 `.github/workflows/ci.yml` 的门禁一致；`AGENTS.md` 是同一组命令的权威表述，此处只作引用，不要各自演化出第二套。
