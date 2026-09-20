# 华通智排AI

通信工程设计文件智能排版工具，基于 Python + customtkinter + Word COM。

## 功能

- 8 套排版模板：华通出版、通用报告、工程合同、学术论文、招投标、党政公文、试卷、范文复刻
- 一至四级标题智能识别（Word 样式名 + 文字正则）
- 批量排版、自动目录、三段式页码分节、表格统一规范
- 多模型 AI 接入（OpenAI 兼容协议：通义千问 / DeepSeek / 智谱 / Ollama / vLLM）
- 原地套样式（保留图片公式）、PDF 导出、水印、修订清理
- 孤行控制、图注绑定、自动更新目录域

## 技术栈

- Python 3.10
- customtkinter（UI）
- python-docx（文档读写）
- pywin32 / Word COM（原地修改）
- OpenAI SDK（大模型接入）

## 打包

```
pyinstaller --noconfirm --clean --windowed --onefile \
  --icon app.ico --name "华通智排AI" \
  --collect-data customtkinter --hidden-import inplace main.py
```

## 开源协议

MIT
