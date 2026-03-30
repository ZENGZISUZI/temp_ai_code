---
name: pdf-extractor
description: PDF内容提取工具。从PDF文件中提取文本、图片、表格，输出为Word文档或独立文件。当用户需要：提取PDF文字、PDF转Word、PDF提取图片、PDF提取表格、PDF内容导出时使用此技能。
---

# PDF提取工具

从PDF文件中提取文本、图片、表格内容。

## 功能

| 模式 | 说明 | 输出 |
|------|------|------|
| `all` | 完整提取 | Word文档(.docx) |
| `text` | 仅文本 | 文本文件(.txt) |
| `images` | 仅图片 | 图片文件夹(.png) |
| `tables` | 仅表格 | CSV文件 |

## 使用方法

运行脚本 `scripts/pdf_extractor.py`，配置以下参数：

```python
if __name__ == '__main__':
    # PDF文件路径
    PDF_PATH = r"D:\AI\test.pdf"
    
    # 输出路径（None表示与PDF同目录）
    OUTPUT_PATH = None
    
    # 提取模式: "all" / "text" / "images" / "tables"
    MODE = "all"
    
    # 是否提取图片（仅MODE="all"时有效）
    EXTRACT_IMAGES = True
```

## 依赖

脚本会自动安装以下依赖：
- pdfplumber - PDF解析
- python-docx - Word生成
- Pillow - 图片处理

## 示例

**提取PDF转Word：**
```python
PDF_PATH = r"D:\AI\report.pdf"
MODE = "all"
```

**仅提取文本：**
```python
PDF_PATH = r"D:\AI\report.pdf"
MODE = "text"
```

**仅提取图片：**
```python
PDF_PATH = r"D:\AI\report.pdf"
MODE = "images"
```

**仅提取表格：**
```python
PDF_PATH = r"D:\AI\report.pdf"
MODE = "tables"
```
