# Docx Flattern Tools

文档格式转换工具，支持 docx 与 Markdown 之间的双向转换。

## 使用方法

1. 安装依赖：
   ```bash
   uv sync
   ```

2. 激活虚拟环境：
   ```bash
   source .venv/bin/activate
   ```

3. Docx 转 Markdown：
   ```bash
   python src/docx_flattern_tools/docx2md.py your_document.docx
   ```
   这将生成同名的 Markdown 文件（例如：your_document.md）

4. Markdown 转 Docx：
   ```bash
   python src/docx_flattern_tools/md2docx.py your_document.md
   ```
   这将生成同名的 Docx 文件（例如：your_document.docx）