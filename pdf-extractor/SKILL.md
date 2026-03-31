---
name: pdf-extractor
description: PDF内容提取工具。从PDF文件中提取文本、图片、表格，输出为Word文档或独立文件。支持扫描版PDF的OCR识别。当用户需要：提取PDF文字、PDF转Word、PDF提取图片、PDF提取表格、PDF内容导出、扫描PDF识别时使用此技能。
---

# PDF提取工具

从PDF文件中提取文本、图片、表格内容，支持扫描版PDF的OCR识别。

## 功能

| 模式 | 说明 | 输出 | 适用场景 |
|------|------|------|----------|
| `all` | 完整提取 | Word文档(.docx) | 正常PDF，提取全部内容 |
| `text` | 仅文本 | 文本文件(.txt) | 只需要文字内容 |
| `images` | 仅图片 | 图片文件夹(.png) | 提取PDF中的图片 |
| `tables` | 仅表格 | CSV文件 | 提取表格数据 |
| `ocr` | OCR识别 | Word文档(.docx) | **扫描版PDF/纯图片PDF** |

## 使用方法

运行脚本 `scripts/pdf_extractor.py`，配置以下参数：

```python
if __name__ == '__main__':
    # PDF文件路径
    PDF_PATH = r"D:\AI\test.pdf"
    
    # 输出路径（None表示与PDF同目录）
    OUTPUT_PATH = None
    
    # 提取模式: "all" / "text" / "images" / "tables" / "ocr"
    MODE = "all"
    
    # 是否提取图片（仅MODE="all"时有效）
    EXTRACT_IMAGES = True
    
    # OCR语言（仅MODE="ocr"时有效）: 'ch'=中文, 'en'=英文
    OCR_LANG = "ch"
```

## 依赖

脚本会自动安装以下依赖：

**普通模式 (all/text/images/tables)：**
- pdfplumber - PDF解析
- python-docx - Word生成
- Pillow - 图片处理

**OCR模式 (ocr)：**
- pdf2image - PDF转图片
- paddleocr - OCR识别
- paddlepaddle - 深度学习框架
- Pillow - 图片处理

> ⚠️ OCR模式首次运行会自动下载PaddleOCR模型（约100MB），需要等待。

## 示例

**提取PDF转Word：**
```python
PDF_PATH = r"D:\AI\report.pdf"
MODE = "all"
```

**扫描版PDF识别（纯图片PDF）：**
```python
PDF_PATH = r"D:\AI\scan.pdf"
MODE = "ocr"
OCR_LANG = "ch"  # 中文识别
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

## 如何判断用哪个模式？

- PDF可以选中文字 → 用 `all` 或 `text`
- PDF全是图片，无法选中文字 → 用 `ocr`
- 只需要图片 → 用 `images`
- 只需要表格数据 → 用 `tables`
