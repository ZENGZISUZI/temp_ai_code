# 动态报告生成器

从 PDF、Excel、图片提取数据，自动生成 Word 报告。

## 安装依赖

```bash
pip install python-docx pdfplumber pandas openpyxl pillow
```

## 快速使用

```python
from report_generator import ReportBuilder, ReportConfig, ReportSection

# 1. 配置报告
config = ReportConfig(
    title="项目测试报告",
    author="张三",
    company="XX公司"
)

# 2. 创建构建器
builder = ReportBuilder(config)

# 3. 加载数据
builder.load_from_pdf("data.pdf", "PDF 数据")
builder.load_from_excel("data.xlsx", "Excel 数据")
builder.add_images(["img1.png", "img2.png"], "图片")

# 4. 生成报告
builder.build("output.docx")
```

## 支持的数据源

| 数据源 | 方法 | 提取内容 |
|--------|------|----------|
| PDF | `load_from_pdf()` | 文本、表格 |
| Excel | `load_from_excel()` | 所有工作表数据 |
| 图片 | `add_images()` | 插入报告 |

## 自定义章节

```python
section = ReportSection(
    title="章节标题",
    content="正文内容...",
    tables=[
        ["列1", "列2", "列3"],
        ["数据1", "数据2", "数据3"]
    ],
    images=["chart.png"],
    subsections=[...]  # 子章节
)

builder.add_custom_section(section)
```

## 目录结构

```
report_generator/
├── report_generator.py   # 核心代码
├── README.md             # 本文件
├── input/                # 输入文件
│   ├── *.pdf
│   ├── *.xlsx
│   └── *.png
└── output/               # 输出报告
    └── report.docx
```
