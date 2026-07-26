# AGENTS.md

本文件适用于仓库根目录及其所有子目录。任何自动化编码代理在修改项目前都必须先阅读本文件；更深层目录若存在新的 `AGENTS.md`，以更深层规则补充或覆盖本文件。

## 项目定位

本项目的业务代码位于 `thesis_project/`，用于把论文素材整理为 Word 论文草案和答辩 PPT 草案，并提供可选的 LLM 增强与“题目 + 参考资料”写作辅助模式。

项目只生成草案和研究辅助材料，不承诺直接产出可提交的最终论文。不得移除 `<请填写>`、`<AI生成，请核对>` 等用于人工复核的标记，也不得把未经来源验证的 LLM 内容描述为事实。

## 事实来源

判断项目真实行为时按以下顺序取证：

1. `thesis_project/src/` 与 `thesis_project/config/` 中的可执行代码；
2. `thesis_project/tests/` 中的自动化测试；
3. `thesis_project/README.md` 与 `docs/ARCHITECTURE.md`；
4. Git 历史。

若 Markdown 与代码矛盾，以代码和测试为准，并在同一变更中更新文档。不要重新创建已经删除的 `docs/superpowers/` 历史实施计划。

## 主要代码路径

- CLI 入口：`thesis_project/src/main.py`
- 输入读取与统一块模型：`thesis_project/src/readers.py`
- 普通草案整理：`thesis_project/src/organizer.py`
- 参考资料模式：`thesis_project/src/synthesizer.py`
- 参考文献与 Crossref：`thesis_project/src/references.py`
- LLM 文本与视觉调用：`thesis_project/src/llm_enhancer.py`、`llm_vision.py`
- Word/PPT 构建：`thesis_project/src/docx_builder.py`、`pptx_builder.py`
- 格式配置：`thesis_project/config/format_spec.py`
- YAML 覆盖：`thesis_project/config/template.py`

完整数据流见 `docs/ARCHITECTURE.md`。

## 开发环境与命令

- 操作系统基线：Windows；命令示例优先使用 PowerShell。
- Python：`>=3.11,<3.14`。
- 工作目录：运行 Python、pytest 和 ruff 时通常进入 `thesis_project/`。
- 核心依赖：`requirements.lock`。
- LLM 依赖：`requirements-llm.lock`。
- Microsoft Word COM：`requirements-office.txt`。

常用验证命令：

```powershell
Set-Location thesis_project
python -m pytest
python -m ruff check src tests config
python src/main.py --help
```

修改范围较小时可先运行相关测试，但交付前应尽量运行全量 `python -m pytest`。当前 Ruff 配置的行宽为 100，目标版本为 Python 3.11。

## 实现约束

- 保持 `readers.py` 的统一 `Document` / `blocks` 数据契约；新增格式必须补读取器和测试。
- 普通 `draft` 模式默认生成 Word 与 PPT；`refs` 模式只生成 Word。不要让提示文字固定声称两种模式产物相同。
- `refs` 自动检测文件名来自 `REFS_SPEC["topic_filenames"]`，不要在多个文件中复制另一套名单。
- LLM 增强必须可选且可降级。普通模式在 LLM 失败时应保留规则结果；参考资料模式必须明确报告失败或降级。
- 不得让 LLM 编造参考文献元数据。缺失字段保留占位符；GB/T 7714 确定性格式化仅由参考文献模块负责。
- Crossref 默认关闭，只能在用户显式传入 `--lookup-metadata` 时访问。
- `--dry-run` 不得调用 LLM、Crossref 或生成 Word/PPT；允许写运行报告和日志。
- 新增网络出口、外发字段或环境变量时，必须同步更新外发确认、README、架构说明和测试。
- `format_spec.py` 中标为“暂未落实”的字段不得在文档中宣称已经生效。
- 配置行为变更必须同步检查 `config/template.py`、构建器的缓存常量和 YAML 模板测试。

## 测试要求

- 修复缺陷时先添加能复现问题的测试，再修改实现。
- 输入读取变更应覆盖正常文件、损坏文件、空内容和编码边界。
- Word/PPT 变更优先验证生成文件中的 XML、结构和边界，不只断言文件存在。
- LLM 测试必须使用 monkeypatch/fake，不访问真实网络，不依赖真实 API 密钥。
- 涉及 CLI 的变更应覆盖退出码、模式选择、dry-run、外发确认和运行报告。
- 不要为了通过测试删除已有断言或降低关键校验强度，除非业务规则确实改变并有明确依据。

## 文件与仓库卫生

- 不提交 `.venv/`、缓存、测试临时目录、`thesis_project/input/`、`output/`、`.cache/` 或冒烟产物。
- `sample_input/` 是测试和演示输入，不是普通文档；除非测试目标要求，不要删除或改成产品说明。
- 不覆盖用户未提交的修改，不清理与当前任务无关的文件。
- 删除文件前确认其确属冗余、生成物或已被替代，并检查仓库内引用。已跟踪文件应可通过 Git 恢复。
- 不执行 `git reset --hard`、强制检出、历史改写或自动提交，除非用户明确要求。
- 依赖变更需同步维护 `pyproject.toml`、相应 `.txt` 文件和锁定文件；不要只修改其中一个来源。

## 文档维护

- 用户操作、CLI 参数、依赖或产物变化时更新 `thesis_project/README.md`。
- 架构、数据流、网络边界或模式职责变化时更新 `docs/ARCHITECTURE.md`。
- 文档必须区分“已实现”“仅告警”“配置存在但暂未落实”。
- 示例命令必须能在 Windows PowerShell 和项目当前目录结构下成立。
- 文档链接使用相对路径，并在交付前检查本地链接没有失效。

## 安全与隐私

- 不读取、打印或提交真实 API 密钥、论文私密素材和用户个人信息。
- 不在日志、测试 fixture、截图或示例中写入真实凭据。
- 发送资料到 LLM 前必须保留现有确认机制；非交互环境需要 `--yes` 或 `THESIS_LLM_CONSENT=1`。
- 不扩大用户授权范围，不自行上传文件、查询外部服务或安装未请求的软件。

## 交付说明

完成任务时应报告：修改文件、行为变化、执行的验证、未执行或失败的检查，以及任何保留的用户未提交内容。不要声称未实际运行的测试已经通过。

<!-- project-upgrade-maintainer:start -->
## Upgrade Workflow Guardrails

When the user explicitly asks to initialize the project upgrade workflow, the agent may:

1. Create and maintain the `.upgrade/` workspace.
2. Create missing base files inside `.upgrade/`.
3. Update supported AI instruction files, `.gitignore`, and `.dockerignore` only through the approved marker blocks.

When the user asks to collect project artifacts, advance a phase, analyze cleanup, organize files, or organize `.upgrade/`, the agent must:

1. Restrict destructive cleanup scope to `.upgrade/` only.
2. Review before moving project files into `.upgrade/`.
3. Do not move or copy project files into `.upgrade/` without explicit user confirmation.
4. Use active-mode classification from the referenced policies.
5. Update `.upgrade/MANIFEST.md` after initialize, sync, collect-apply, phase-update, or cleanup-apply.
6. Output a maintenance report or trace after initialize, sync, collect-review, collect-apply, phase-update, cleanup-review, cleanup-apply, or health-check.
7. Require two confirmations for cleanup suggestions: first to inspect details, then to execute cleanup.

Hard prohibitions:

1. Do not delete any file outside `.upgrade/`.
2. Do not overwrite user content outside managed marker blocks.
3. Do not use unscoped bulk deletion.
4. Do not treat draft documents as final plans unless the user confirms.
5. Do not use `git add .`.
6. Do not move high-risk or referenced files unless the user explicitly approves the exact file.

<!-- project-upgrade-maintainer:end -->
