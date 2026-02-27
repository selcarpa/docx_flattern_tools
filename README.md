# Docx Flattern Tools

将 docx 文档转换为 Markdown 格式的工具。

## 使用方法

1. 安装依赖：
   ```bash
   uv sync
   ```

2. 激活虚拟环境：
   ```bash
   source .venv/bin/activate
   ```

3. 转换单个文档：
   ```bash
   python src/docx_flattern_tools/docx2md.py your_document.docx
   ```
   这将生成同名的 Markdown 文件（例如：your_document.md）