# Academic Writing Assistant

论文 Word 草案 + 答辩 PPT 草案生成器。

读取 Word / PDF / TXT / Markdown / JSON / Excel / 图片 源文件，整理成结构化内容并按格式规范生成论文与答辩材料草案。

> 定位是**草案**：产物已套好格式与结构骨架，正文与 `<请填写>` 占位符需人工核对润色。

## 快速开始

业务代码位于 `thesis_project/`：

```bash
cd thesis_project
pip install -r requirements.lock
python src/main.py --input sample_input
```

Windows 用户可直接双击 `thesis_project/run.bat`。

## 两种模式

- **draft**：`readers -> organizer -> 可选 llm_enhancer -> docx/pptx builder`，默认生成 Word 和 PPT。
- **refs**：`readers -> synthesizer -> references/可选 vision -> docx builder`，题目 + 参考资料，只生成 Word。

两种模式都可接 `--refresh-fields` / `--pdf` 做 Word 域刷新与 PDF 导出（依赖本机 Word）。

## 网络与隐私

- Crossref 元数据查询默认关闭，需显式 `--lookup-metadata`。
- 向 LLM 端点外发资料前需要确认；`--dry-run` 不允许任何网络调用。
- 扫描版 PDF 默认不做 OCR，PDF 内嵌图片默认不导入；需分别显式 `--ocr` / `--extract-pdf-images` 启用。

## 文档

- [使用说明](thesis_project/README.md) — 完整 CLI 参数、输入格式、环境变量
- [架构说明](docs/ARCHITECTURE.md) — 模块职责与数据流
- [AGENTS.md](AGENTS.md) — 协作与验证规范

## 开发

```bash
cd thesis_project
python -m pytest
python -m ruff check src tests config
```

## 许可证

[MIT](LICENSE)
