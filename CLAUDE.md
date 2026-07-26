# CLAUDE.md

Claude 在本仓库工作时必须先读取并遵守根目录 `AGENTS.md`。本文件提供 Claude 专用补充规则；与 `AGENTS.md` 冲突时，以更严格且更贴近用户当前请求的规则为准。

## 开始工作前

1. 查看 `git status --short`，识别并保留用户已有修改。
2. 阅读 `docs/ARCHITECTURE.md`、任务涉及的源码、测试和 README 对应章节。
3. 先用搜索定位真实调用链，不根据历史计划或文件名猜测实现。
4. 任务涉及删除、合并或重命名时，先搜索引用并确认替代文件。

不要把 `.claude/settings.local.json` 中的本机权限配置当作项目业务规范，也不要未经用户要求修改该文件。

## 项目事实摘要

- 业务根目录：`thesis_project/`
- CLI：`src/main.py`
- 普通模式：`readers -> organizer -> 可选 llm_enhancer -> docx/pptx builder`
- 参考资料模式：`readers -> synthesizer -> references/可选 vision -> docx builder`
- `refs` 模式只生成 Word；普通模式默认生成 Word 和 PPT。
- Crossref 默认关闭；LLM 外发需要确认；dry-run 不允许网络调用。
- 扫描版 PDF 不做 OCR。
- AI 正文保留 `<AI生成，请核对>`；未知内容保留 `<请填写>` 或参考文献补全占位符。

## 修改原则

- 以最小、可验证的改动解决问题，避免无关重构。
- 修改业务行为时同步修改或新增 pytest 测试。
- 修改 CLI、输入格式、产物、依赖、环境变量或网络行为时同步更新 README。
- 修改模块职责或数据流时同步更新 `docs/ARCHITECTURE.md`。
- 不恢复 `docs/superpowers/` 下已经删除的历史计划；需要记录新决策时写简短的当前文档，不复制大段实现代码。
- 不把 `format_spec.py` 中未被构建器消费的字段描述为已落实。
- 不让普通 draft 模式错误宣称参考文献已经统一转换为 GB/T 7714。

## 命令与验证

在 `thesis_project/` 下执行：

```powershell
python -m pytest
python -m ruff check src tests config
python src/main.py --help
```

若全量 Ruff 暴露与当前任务无关的既有问题，应只修复本次修改引入的问题，并在交付说明中如实记录。测试或命令失败时先判断是代码问题、环境问题还是权限问题，不得通过删除测试、跳过校验或访问真实网络来规避。

## Claude 工具使用边界

- 优先使用只读搜索和小范围编辑；保留用户工作区中的未提交内容。
- 不自动执行 `git add`、`git commit`、推送、强制检出或历史改写。
- 不使用真实 LLM API、Crossref 或其他网络服务进行测试。
- 不删除 `thesis_project/input/`、`output/`、用户文档或 `.claude/` 内容。
- 缓存和冒烟产物只有在用户要求清理且目标路径已精确确认时才可删除。
- 不在输出中展示环境变量值、密钥、私密论文内容或本机敏感路径中的数据。

## 回复要求

交付时简洁列出：

- 修改了什么；
- 为什么与当前源码一致；
- 运行了哪些测试或检查及结果；
- 哪些内容因权限、环境或任务范围未处理。

不得把推测写成已验证事实，也不得声称生成了实际不存在的 Word、PPT、PDF 或提交记录。

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
