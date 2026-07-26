# 项目架构说明

本文描述当前源码实际实现。用户用法以 `thesis_project/README.md` 为准；字段、参数和行为发生变化时，应同时更新对应测试与文档。

## 入口与模式

命令行入口是 `thesis_project/src/main.py`。程序先通过 `readers.py` 将输入统一读取为 `Document`，再按模式进入不同管道：

```text
输入文件
  -> readers.py：统一 Document / blocks
  -> main.py：自动检测或 --mode 指定
       -> draft：organizer.py -> 可选 llm_enhancer.py -> docx_builder.py / pptx_builder.py
       -> refs： synthesizer.py -> references.py / 可选 llm_vision.py -> docx_builder.py
  -> runtime_report.py：运行报告
```

- `auto`：存在 `topic.md`、`topic.txt`、`题目.md` 或 `题目.txt` 时进入 `refs`，否则进入 `draft`。
- `draft`：生成 Word 与 PPT，可用 `--only` 限制产物；`--llm` 是可选增强。
- `refs`：要求题目文件、至少一份参考资料和 `LLM_API_KEY`，只生成 Word。
- `--dry-run`：读取和检查输入，必要时打印 LLM 外发清单并写运行报告，不调用 LLM、不生成 Word/PPT。

## 输入模型

`readers.py` 当前支持：

- 文本：`.txt`、`.text`、`.md`、`.markdown`、`.json`
- 文档：`.docx`、`.pdf`
- 表格：`.xlsx`、`.csv`
- 图片：`.png`、`.jpg`、`.jpeg`、`.bmp`、`.webp`

统一内容块以 `kind` 区分段落、标题、表格和图片。Word 内嵌图片会被提取；PDF 依赖文本层，扫描版 PDF 不做 OCR。

## 普通草案管道

`organizer.py` 负责元信息抽取、章节树重建、特殊章节处理、无标题文档骨架回退和 PPT 分区映射。`llm_enhancer.py` 只在显式传入 `--llm` 时参与，并在单步失败时保留规则结果。

`docx_builder.py` 与 `pptx_builder.py` 消费整理后的结构数据。格式以 `config/format_spec.py` 为默认值，可通过 `--format-template` 加载 YAML 深度覆盖。源码中标为 `暂未落实` 的配置字段不影响产物。

## 参考资料管道

`synthesizer.py` 是 `refs` 模式的整理入口：

1. 将参考资料逐篇整理为摘要卡；
2. 生成或回退到章节大纲；
3. 为综述章节生成带引用标记的初稿；
4. 为核心章节生成写作要点，不生成核心研究正文；
5. 由 `references.py` 确定性生成参考文献并校验引用；
6. 将表格与图片挂载到匹配章节，无法匹配时放入素材附录。

所有 AI 生成正文必须保留 `<AI生成，请核对>` 标记。Crossref 查询默认关闭，仅 `--lookup-metadata` 显式启用，并缓存到 `thesis_project/.cache/reference_metadata.json`。

## 外部依赖与网络边界

- 核心依赖：`requirements.lock`
- LLM 依赖：`requirements-llm.lock`
- Word 域刷新/PDF 导出：`requirements-office.txt`，仅 Windows + Microsoft Word
- LLM 网络请求：`llm_enhancer.py` 和 `llm_vision.py`
- Crossref 网络请求：`references.py`，仅显式启用元数据查询时发生

启用 LLM 时，交互环境需确认外发；非交互环境需 `--yes` 或 `THESIS_LLM_CONSENT=1`。`--dry-run` 不调用任何外部服务。

## 事实来源与验证

维护时按以下优先级判断真实行为：

1. 可执行源码与配置；
2. 自动化测试；
3. `thesis_project/README.md` 与本文。

历史实施计划和设计草稿不作为运行时事实来源。项目回归命令：

```powershell
cd thesis_project
python -m pytest
python src/main.py --help
```
