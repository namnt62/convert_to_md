# Batch_Doc_To_MD
一个批量处理脚本，用于自动将（一个文件夹下所有的、包括子文件夹内的） .doc、.docx 和 .pdf 文件快速批量转换为 Markdown (.md) 格式。

## 项目描述

本脚本帮助用户高效地将大量 Word 和 PDF 文档自动化地转换为 Markdown 文件，适合大模型RAG知识库管理、文档整理和内容迁移等场景。

## 功能特性

- 自动批量转换 `.doc` 文件为 `.docx` 格式。
- 自动将 `.docx` 文件转换为 Markdown 格式。
- 自动提取 PDF 文件文本并转换为 Markdown 格式。
- 支持通用文本文件转换。
- 转换过程自动清理中间文件，保持目录清洁。

## 环境要求

- Python 3.x  

## Cài đặt thư viện

```bash
pip install pdfplumber python-docx pywin32 markitdown
```

## Cài đặt bổ sung
```https://github.com/jgm/pandoc/releases```

## 使用方法

1. 克隆本项目到本地：

```bash
git clone https://github.com/yourusername/BatchDocToMarkdown.git
```

2. 修改脚本中源文件夹和目标文件夹的路径，请注意不要互相包含：

```python
source_folder = r"你的源文件夹路径"
target_folder = r"你的目标文件夹路径"
```

3. 运行脚本：

```bash
python process.py
```

## 项目结构

```
BatchDocToMarkdown/
├── process.py          # 主处理脚本
├── README.md           # 项目说明文档
├── LICENSE             # 项目许可证
└── requirements.txt    # Python 依赖包列表
```

## 开源许可证

本项目采用 MIT 许可证。

## 贡献与反馈

欢迎提交 Pull Request 进行代码优化和功能扩展，也可以在 Issues 中提出功能请求或报告 bug。

---

Enjoy your batch processing!
