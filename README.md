## md2docx

将 Markdown 转换为 Office 文档（.docx 或 .doc）。

### 安装

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

可选（推荐）：安装系统 Pandoc 以获得最佳转换质量。

- macOS（使用 Homebrew）：
```bash
brew install pandoc
```

可选支持 .doc 导出：安装 LibreOffice（提供 `soffice`）。

### 使用

- 转换为 DOCX（默认）
```bash
python md2doc.py input.md -o output.docx
```

- 转换为旧版 DOC（需要 LibreOffice）
```bash
python md2doc.py input.md -o output.doc
```

若未提供 `-o`，将生成与输入同名的 `.docx`。

- 设置统一页眉（宋体，小五）
```bash
python md2doc.py input.md -o output.docx --header "复旦大学硕士学位论文"
```

- 使用配置文件（支持 YAML/JSON）
```bash
python md2doc.py input.md -o output.docx --config configs/fudan.yml
```

### 说明

- 优先使用 Pandoc（需系统安装 `pandoc`）。
- 若无 Pandoc，则使用 `python-docx` 基础渲染（支持标题、段落、粗体/斜体文本的基本输出、行内代码、代码块、列表、引用、水平线与链接文本）。
- 输出 `.doc` 时：先生成 `.docx`，再尝试用 LibreOffice 无头转换，失败则保留 `.docx` 并提示。

### 版式与样式（默认）

- 纸张：A4（210mm × 297mm）
- 页边距：上 30mm，下 25mm，左 30mm，右 25mm
- 正文：中文宋体（SimSun），英文与数字 Times New Roman；小四（12pt）；行距固定 20 磅
- 一级标题：黑体，三号（16pt），居中；段前 1 行，段后 1 行
- 二级标题：黑体，四号（14pt），左对齐；段前 1 行，段后 1 行
- 三级标题：黑体，小四（12pt），左对齐；段前 1 行，段后 1 行
- 页眉页脚：宋体，小五（9pt）；可通过 `--header` 设置页眉文本
  - 也可通过配置文件的 `header.text` 指定（优先级：命令行 `--header` > 配置文件）

### 已知限制（回退渲染）

- 复杂表格、公式、脚注、目录等在回退模式下支持有限；建议安装 Pandoc。
- 图片与超链接在回退模式下仅输出为文本（可定制扩展）。
 - 奇偶页不同页眉在当前版本未单独区分（统一按 `--header` 应用）。



