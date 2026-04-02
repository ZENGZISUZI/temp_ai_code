# -*- coding: utf-8 -*-
"""
Excel测试报告转Word文档工具
自动提取Excel数据生成标准Word报告模板
依赖: pip install pandas openpyxl python-docx xlrd
      - openpyxl: 支持 .xlsx 格式
      - xlrd: 支持 .xls 格式（旧版Excel）
"""

import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import re


# ==================== 智能字段映射配置 ====================

# 概述部分字段映射（Excel列名 -> Word报告字段）
OVERVIEW_FIELD_MAPPING = {
    # 产品信息相关
    '产品信息': ['零部件名称', '产品名称', '产品型号', '设备名称', '名称'],
    # 试验信息相关
    '试验信息': ['试验名称', '试验类型', '测试类型', '试验目的'],
    # 工作模式相关
    '工作模式': ['工作模式', '运行模式', '工作状态', '模式'],
    # 测试仪器设备相关
    '测试仪器设备': ['测试仪器', '仪器设备', '试验设备', '设备', '使用仪器'],
}

# 小用例表格字段映射（Word字段 -> Excel可能列名）
TESTCASE_FIELD_MAPPING = {
    '开始日期': ['开始日期', '开始时间', '起始日期', '开始'],
    '结束日期': ['结束日期', '结束时间', '终止日期', '结束'],
    '样机数量': ['样机数量', '数量', '样品数量', '台数'],
    '样机编号': ['样机编号', '编号', '样品编号', '机号'],
    '试验机构': ['试验机构', '检测机构', '测试机构', '机构'],
    '试验环境': ['试验环境', '环境条件', '环境', '测试环境'],
    '试验标准': ['试验标准', '标准', '测试标准', '参考标准'],
    '试验条件': ['试验条件', '条件', '测试条件'],
    '规格要求': ['规格要求', '要求', '技术要求', '规格'],
    '试验数据': ['试验数据', '数据', '测试数据', '结果数据'],
    '试验结论': ['试验结论', '结论', '测试结论', '结果'],
}


def clean_case_number(name):
    """
    清理用例名字中的数字前缀
    
    支持格式:
    - "1、xxx" -> "xxx"
    - "一、xxx" -> "xxx"
    - "1.xxx" -> "xxx"
    - "1 xxx" -> "xxx"
    - "（1）xxx" -> "xxx"
    - "(1)xxx" -> "xxx"
    
    参数:
        name: 用例名字
        
    返回:
        清理后的名字
    """
    if not name:
        return name
    
    # 中文数字映射
    chinese_nums = '一二三四五六七八九十'
    
    # 匹配模式：数字/中文数字 + 标点符号
    patterns = [
        r'^[\d]+\s*[、.．。:：]\s*',  # 1、 1. 1．
        r'^[一二三四五六七八九十]+\s*[、.．。:：]\s*',  # 一、 一.
        r'^[\(（][\d]+[）\)]\s*',  # (1) （1）
        r'^[\d]+\s+',  # 1 开头加空格
    ]
    
    result = name.strip()
    for pattern in patterns:
        result = re.sub(pattern, '', result)
    
    return result.strip()


def extract_case_number(name):
    """
    提取用例名字中的数字前缀
    
    参数:
        name: 用例名字
        
    返回:
        数字（阿拉伯数字），如果没有则返回None
    """
    if not name:
        return None
    
    # 中文数字映射
    chinese_to_num = {
        '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
        '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
    }
    
    # 尝试匹配中文数字
    match = re.match(r'^([一二三四五六七八九十]+)\s*[、.．。:：]', name)
    if match:
        chinese_num = match.group(1)
        return chinese_to_num.get(chinese_num)
    
    # 尝试匹配阿拉伯数字
    match = re.match(r'^[\(（]?(\d+)[）\)]?\s*[、.．。:：]?', name)
    if match:
        return int(match.group(1))
    
    return None


def validate_chapter_position(big_cases):
    """
    验证大用例是否放在正确的章节位置
    
    参数:
        big_cases: 大用例列表
        
    返回:
        验证结果列表 [{'name': 'xxx', 'expected': 1, 'actual': 1, 'valid': True}, ...]
    """
    results = []
    
    for idx, big_case in enumerate(big_cases, 1):
        name = big_case.get('name', '')
        expected_num = extract_case_number(name)
        
        result = {
            'name': name,
            'expected': expected_num,
            'actual': idx,
            'valid': expected_num is None or expected_num == idx
        }
        
        if not result['valid']:
            print(f"⚠️ 警告: 大用例 '{name}' 位置不匹配，期望在第{expected_num}章节，实际在第{idx}章节")
        
        results.append(result)
    
    return results


def find_best_match(target_field, excel_columns, mapping_dict):
    """
    智能匹配：根据目标字段在Excel列名中找最佳匹配

    参数:
        target_field: 目标字段名（如"产品信息"）
        excel_columns: Excel实际列名列表
        mapping_dict: 字段映射字典

    返回:
        匹配到的Excel列名，未匹配返回None
    """
    if target_field not in mapping_dict:
        return None

    keywords = mapping_dict[target_field]

    for keyword in keywords:
        for col in excel_columns:
            if col and keyword in str(col):
                return col

    return None


def smart_match_field(field_name, excel_columns):
    """
    智能匹配单个字段

    参数:
        field_name: 字段名
        excel_columns: Excel列名列表

    返回:
        匹配到的列名
    """
    # 先尝试精确匹配
    if field_name in excel_columns:
        return field_name

    # 再尝试包含匹配
    for col in excel_columns:
        if col and field_name in str(col):
            return col

    # 尝试反向匹配（Excel列名包含在字段名中）
    for col in excel_columns:
        if col and str(col) in field_name:
            return col

    return None


def set_cell_font(cell, font_name='宋体', font_size=10.5, bold=False):
    """设置单元格字体"""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def set_table_border(table):
    """设置表格边框"""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')

    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)

    tblPr.append(tblBorders)
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def add_heading_with_number(doc, text, level=1):
    """添加带编号的标题"""
    heading = doc.add_heading(text, level=level)
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in heading.runs:
        run.font.name = '黑体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    return heading


def create_testcase_table(doc, data_dict):
    """
    创建测试用例表格

    参数:
        doc: Word文档对象
        data_dict: 数据字典 {字段名: 值}
    """
    # 表格字段顺序
    fields = ['开始日期', '结束日期', '样机数量', '样机编号', '试验机构',
              '试验环境', '试验标准', '试验条件', '规格要求', '试验数据', '试验结论']

    # 创建表格（11行2列：字段名 | 值）
    table = doc.add_table(rows=len(fields), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_border(table)

    for i, field in enumerate(fields):
        # 第一列：字段名
        cell0 = table.cell(i, 0)
        cell0.text = field
        set_cell_font(cell0, bold=True)

        # 第二列：值
        cell1 = table.cell(i, 1)
        value = data_dict.get(field, '')
        cell1.text = str(value) if value else ''
        set_cell_font(cell1)

    doc.add_paragraph()  # 空行
    return table


class ExcelToWordReport:
    """Excel转Word报告主类"""

    def __init__(self, excel_path, word_path=None):
        """
        初始化

        参数:
            excel_path: Excel文件路径
            word_path: Word输出路径，默认同名
        """
        self.excel_path = excel_path
        self.word_path = word_path or os.path.splitext(excel_path)[0] + '_报告.docx'

        # 读取Excel
        self.df = None
        self.excel_columns = []

        # 解析后的数据
        self.overview_data = {}  # 概述数据
        self.big_cases = []  # 大用例列表 [{'name': 'aaaa', 'small_cases': [...]}]
        self.summary_data = []  # 汇总数据
        self.col_name_to_idx = {}  # 列名到索引的映射

    def load_excel(self, sheet_name=0):
        """
        加载Excel文件（自动识别xlsx/xls格式）

        参数:
            sheet_name: sheet名称或索引
        """
        # 先检测实际格式
        actual_format = detect_excel_format(self.excel_path)
        file_ext = os.path.splitext(self.excel_path)[1].lower()
        
        print(f"文件扩展名: {file_ext}, 实际格式: {actual_format}")
        
        # 根据实际格式选择引擎
        if actual_format == 'xls':
            engine = 'xlrd'
        else:
            engine = 'openpyxl'

        print(f"使用引擎: {engine}")

        try:
            self.df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None, engine=engine)
        except Exception as e:
            # 如果默认引擎失败，尝试另一个
            alt_engine = 'xlrd' if engine == 'openpyxl' else 'openpyxl'
            print(f"引擎 {engine} 失败: {e}")
            print(f"尝试 {alt_engine}...")
            self.df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None, engine=alt_engine)

        self.excel_columns = [str(col) for col in self.df.iloc[0].tolist() if pd.notna(col)]
        print(f"加载Excel成功，共 {len(self.df)} 行")
        print(f"检测到列名: {self.excel_columns[:10]}...")  # 显示前10个

    def find_test_project_column(self):
        """找到"试验项目"列（在所有行中搜索）"""
        # 在所有行中搜索"试验项目"列名
        for row_idx in range(len(self.df)):
            for col_idx, cell in enumerate(self.df.iloc[row_idx]):
                if pd.notna(cell) and '试验项目' in str(cell):
                    print(f"找到'试验项目'列: 第{row_idx + 1}行, 第{col_idx + 1}列")
                    return col_idx, row_idx  # 返回列索引和标题行索引
        return None, None

    def find_merged_cells_info(self):
        """
        检测合并单元格（大用例）
        返回: [(起始行, 结束行, 大用例名), ...]
        """
        # 根据文件格式选择库
        file_ext = os.path.splitext(self.excel_path)[1].lower()

        if file_ext == '.xls':
            # xls格式用xlrd，不支持合并单元格检测
            print("警告: .xls格式不支持合并单元格检测，将使用备用解析方式")
            return [], None, None

        # xlsx格式使用openpyxl读取合并单元格信息
        from openpyxl import load_workbook

        wb = load_workbook(self.excel_path)
        ws = wb.active

        # 找到试验项目列（在所有行中搜索）
        test_col = None
        header_row = None
        for row_idx in range(1, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value and '试验项目' in str(cell_value):
                    test_col = col_idx
                    header_row = row_idx
                    print(f"找到'试验项目'列: 第{row_idx}行, 第{col_idx}列")
                    break
            if test_col:
                break

        if not test_col:
            print("未找到'试验项目'列")
            return [], None, None

        merged_ranges = []
        header_merge_end = header_row  # 标题行合并的结束行
        
        # 先找到标题行的合并范围（可能是行合并或列合并）
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row == header_row and merged_range.min_col <= test_col <= merged_range.max_col:
                header_merge_end = merged_range.max_row
                print(f"标题行合并范围: 第{header_row}-{header_merge_end}行")
                break
        
        # 检测大用例：查找标题行下方的行合并单元格
        # 大用例名字在试验项目列，通过行合并（多行合并成一格）来标识
        for merged_range in ws.merged_cells.ranges:
            # 行合并：min_row != max_row，且在试验项目列
            if merged_range.min_col == test_col and merged_range.min_row > header_merge_end:
                start_row = merged_range.min_row
                end_row = merged_range.max_row
                cell_value = ws.cell(row=start_row, column=test_col).value
                if cell_value and str(cell_value).strip():
                    merged_ranges.append((start_row, end_row, str(cell_value).strip()))
                    print(f"找到大用例(行合并): 第{start_row}-{end_row}行, 名字: {cell_value}")
        
        # 如果没找到行合并，尝试检测列合并（横向合并）
        if not merged_ranges:
            print("未检测到行合并，尝试检测列合并...")
            for row_idx in range(header_merge_end + 1, ws.max_row + 1):
                for merged_range in ws.merged_cells.ranges:
                    # 列合并：min_col != max_col，且包含试验项目列
                    if (merged_range.min_row == row_idx and 
                        merged_range.min_col <= test_col <= merged_range.max_col and
                        merged_range.min_col != merged_range.max_col):
                        cell_value = ws.cell(row=row_idx, column=merged_range.min_col).value
                        if cell_value and str(cell_value).strip():
                            # 用行号作为标识，合并范围结束行作为小用例开始
                            merged_ranges.append((row_idx, row_idx, str(cell_value).strip()))
                            print(f"找到大用例(列合并): 第{row_idx}行, 列{merged_range.min_col}-{merged_range.max_col}, 名字: {cell_value}")
                        break

        # 按行号排序
        merged_ranges.sort(key=lambda x: x[0])
        return merged_ranges, test_col, header_row

    def parse_overview_data(self, big_case_start_row):
        """
        解析概述数据（大用例之前的内容）

        参数:
            big_case_start_row: 第一个大用例的起始行
        """
        # 概述数据在大用例之前，通常是键值对形式
        for row_idx in range(1, big_case_start_row):
            row_data = self.df.iloc[row_idx]

            # 查找键值对（假设在D、E列或相邻列）
            for col_idx in range(len(row_data) - 1):
                key = row_data.iloc[col_idx]
                value = row_data.iloc[col_idx + 1]

                if pd.notna(key) and pd.notna(value):
                    key_str = str(key).strip()
                    value_str = str(value).strip()

                    # 智能匹配到概述字段
                    for field, keywords in OVERVIEW_FIELD_MAPPING.items():
                        if any(kw in key_str for kw in keywords):
                            self.overview_data[field] = value_str
                            break

        print(f"解析概述数据: {self.overview_data}")

    def parse_test_cases(self):
        """解析测试用例（大用例和小用例）"""
        merged_ranges, test_col, header_row = self.find_merged_cells_info()

        if not merged_ranges:
            print("未检测到合并单元格，尝试其他方式解析...")
            # 备用解析逻辑
            self.parse_without_merge()
            return

        # 解析概述数据（第一个大用例之前）
        first_big_case_row = merged_ranges[0][0]
        self.parse_overview_data(first_big_case_row)

        # 建立列名到索引的映射（使用标题行）
        self.build_column_mapping(header_row)

        # 解析大用例和小用例
        for i, (start_row, end_row, big_case_name) in enumerate(merged_ranges):
            big_case = {
                'name': big_case_name,
                'small_cases': []
            }

            # 小用例在合并单元格下方（从end_row+1到下一个大用例之前）
            next_start = merged_ranges[i + 1][0] if i + 1 < len(merged_ranges) else len(self.df) + 1

            for row_idx in range(end_row + 1, next_start):
                if row_idx <= len(self.df):
                    row_data = self.df.iloc[row_idx - 1]  # pandas索引从0开始
                    small_case_name = row_data.iloc[test_col - 1] if test_col else None

                    if pd.notna(small_case_name) and str(small_case_name).strip():
                        small_case = {
                            'name': str(small_case_name),
                            'data': self.extract_testcase_data(row_data)
                        }
                        big_case['small_cases'].append(small_case)

                        # 添加到汇总数据
                        self.summary_data.append({
                            '序号': len(self.summary_data) + 1,
                            '试验分类': clean_case_number(big_case_name),
                            '试验项目': clean_case_number(str(small_case_name)),
                            '测试结论': small_case['data'].get('试验结论', '')
                        })

            self.big_cases.append(big_case)

        print(f"解析完成: {len(self.big_cases)} 个大用例, 共 {len(self.summary_data)} 个小用例")

    def build_column_mapping(self, header_row):
        """
        建立列名到索引的映射

        参数:
            header_row: 标题行号（openpyxl格式，从1开始）
        """
        self.col_name_to_idx = {}
        
        # 根据文件格式选择方式
        file_ext = os.path.splitext(self.excel_path)[1].lower()
        
        if file_ext == '.xls':
            # xls格式直接从DataFrame读取
            if header_row and header_row < len(self.df):
                for col_idx, col_name in enumerate(self.df.iloc[header_row]):
                    if pd.notna(col_name):
                        self.col_name_to_idx[str(col_name).strip()] = col_idx
        else:
            # xlsx格式使用openpyxl读取标题行
            from openpyxl import load_workbook
            wb = load_workbook(self.excel_path)
            ws = wb.active
            
            for col_idx in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=header_row, column=col_idx).value
                if cell_value:
                    self.col_name_to_idx[str(cell_value).strip()] = col_idx - 1  # 转为pandas索引（从0开始）
        
        print(f"列名映射: {list(self.col_name_to_idx.keys())[:15]}...")

    def parse_without_merge(self):
        """无合并单元格时的备用解析"""
        test_col, header_row = self.find_test_project_column()
        if test_col is None:
            print("无法找到试验项目列")
            return

        # 建立列名映射
        self.build_column_mapping(header_row + 1 if header_row else 1)

        current_big_case = None

        for row_idx in range(1, len(self.df)):
            row_data = self.df.iloc[row_idx]
            test_project = row_data.iloc[test_col]

            if pd.notna(test_project):
                # 判断是否是大用例（简单规则：看是否缩进或特殊标记）
                # 这里需要根据实际情况调整
                pass

    def extract_testcase_data(self, row_data):
        """
        从行数据中提取测试用例数据

        参数:
            row_data: DataFrame的一行

        返回:
            字典 {字段名: 值}
        """
        data = {}

        # 如果没有列名映射，尝试建立
        if not hasattr(self, 'col_name_to_idx') or not self.col_name_to_idx:
            # 默认使用第一行作为标题
            self.col_name_to_idx = {}
            for idx, col_name in enumerate(self.df.iloc[0]):
                if pd.notna(col_name):
                    self.col_name_to_idx[str(col_name)] = idx

        # 智能匹配每个字段
        for field, keywords in TESTCASE_FIELD_MAPPING.items():
            for keyword in keywords:
                for col_name, idx in self.col_name_to_idx.items():
                    if keyword in col_name:
                        value = row_data.iloc[idx] if idx < len(row_data) else None
                        if pd.notna(value):
                            data[field] = str(value)
                            break
                if field in data:
                    break

        return data

    def generate_word_report(self):
        """生成Word报告"""
        doc = Document()

        # 设置默认字体
        doc.styles['Normal'].font.name = '宋体'
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        # ===== 1. 概述 =====
        add_heading_with_number(doc, '1 概述', level=1)

        # 1.1 产品信息
        add_heading_with_number(doc, '1.1 产品信息', level=2)
        doc.add_paragraph(self.overview_data.get('产品信息', '（待填写）'))

        # 1.2 试验信息
        add_heading_with_number(doc, '1.2 试验信息', level=2)
        doc.add_paragraph(self.overview_data.get('试验信息', '（待填写）'))

        # 1.3 工作模式
        add_heading_with_number(doc, '1.3 工作模式', level=2)
        doc.add_paragraph(self.overview_data.get('工作模式', '（待填写）'))

        # 1.4 测试仪器设备
        add_heading_with_number(doc, '1.4 测试仪器设备', level=2)
        doc.add_paragraph(self.overview_data.get('测试仪器设备', '（待填写）'))

        # ===== 2. 试验结果汇总 =====
        add_heading_with_number(doc, '2 试验结果汇总', level=1)

        # 创建汇总表格
        summary_table = doc.add_table(rows=len(self.summary_data) + 1, cols=4)
        summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_border(summary_table)

        # 表头
        headers = ['序号', '试验分类', '试验项目', '测试结论']
        for i, header in enumerate(headers):
            cell = summary_table.cell(0, i)
            cell.text = header
            set_cell_font(cell, bold=True)

        # 数据行
        for row_idx, item in enumerate(self.summary_data):
            for col_idx, header in enumerate(headers):
                cell = summary_table.cell(row_idx + 1, col_idx)
                cell.text = str(item.get(header, ''))
                set_cell_font(cell)
        
        # 合并相同试验分类的单元格
        if self.summary_data:
            current_category = None
            merge_start = 1
            
            for row_idx, item in enumerate(self.summary_data, 1):
                category = item.get('试验分类', '')
                
                if category != current_category:
                    # 如果是新分类，合并前一个分类的单元格
                    if current_category is not None and row_idx > merge_start + 1:
                        # 合并试验分类列（第2列，索引1）
                        summary_table.cell(merge_start, 1).merge(summary_table.cell(row_idx - 1, 1))
                    current_category = category
                    merge_start = row_idx
            
            # 合并最后一个分类
            if len(self.summary_data) > 1 and merge_start < len(self.summary_data):
                summary_table.cell(merge_start, 1).merge(summary_table.cell(len(self.summary_data), 1))

        doc.add_paragraph()

        # ===== 3. 测试数据 =====
        add_heading_with_number(doc, '3 测试数据', level=1)
        
        # 验证大用例章节位置
        print("\n验证大用例章节位置...")
        validation_results = validate_chapter_position(self.big_cases)
        invalid_count = sum(1 for r in validation_results if not r['valid'])
        if invalid_count > 0:
            print(f"⚠️ 发现 {invalid_count} 个大用例位置不匹配")
        else:
            print("✓ 所有大用例位置正确")

        for big_idx, big_case in enumerate(self.big_cases, 1):
            # 清理大用例名字中的数字前缀
            clean_name = clean_case_number(big_case["name"])
            # 大用例标题 (3.1, 3.2, ...)
            add_heading_with_number(doc, f'3.{big_idx} {clean_name}', level=2)

            for small_idx, small_case in enumerate(big_case['small_cases'], 1):
                # 清理小用例名字中的数字前缀
                clean_small_name = clean_case_number(small_case["name"])
                # 小用例标题 (3.1.1, 3.1.2, ...)
                add_heading_with_number(doc, f'3.{big_idx}.{small_idx} {clean_small_name}', level=3)

                # 小用例表格
                create_testcase_table(doc, small_case['data'])

        # 保存文档
        doc.save(self.word_path)
        print(f"Word报告已生成: {self.word_path}")
        return self.word_path


def detect_excel_format(file_path):
    """
    检测Excel文件的实际格式（通过文件头）
    
    返回:
        'xlsx' - Excel 2007+ (Office Open XML)
        'xls' - Excel 97-2003 (BIFF)
    """
    with open(file_path, 'rb') as f:
        header = f.read(8)
    
    # xlsx 实际是 ZIP 压缩包，以 PK 开头
    if header[:2] == b'PK':
        return 'xlsx'
    
    # xls 是 OLE 复合文档，以 D0 CF 11 E0 开头
    if header[:4] == b'\xD0\xCF\x11\xE0':
        return 'xls'
    
    # 默认返回xlsx
    return 'xlsx'


def list_sheets(excel_path):
    """
    列出Excel中所有sheet名称

    参数:
        excel_path: Excel文件路径

    返回:
        sheet名称列表
    """
    # 先检测实际格式
    actual_format = detect_excel_format(excel_path)
    file_ext = os.path.splitext(excel_path)[1].lower()
    
    print(f"文件扩展名: {file_ext}, 实际格式: {actual_format}")
    
    if actual_format == 'xls':
        # xls格式用xlrd
        try:
            import xlrd
            wb = xlrd.open_workbook(excel_path)
            return wb.sheet_names()
        except Exception as e:
            print(f"xlrd读取失败: {e}")
            # 尝试用pandas
            import pandas as pd
            xl = pd.ExcelFile(excel_path, engine='xlrd')
            return xl.sheet_names
    else:
        # xlsx格式用openpyxl
        try:
            from openpyxl import load_workbook
            wb = load_workbook(excel_path, read_only=True)
            sheets = wb.sheetnames
            wb.close()
            return sheets
        except Exception as e:
            print(f"openpyxl读取失败: {e}")
            # 可能实际是xls格式，尝试xlrd
            try:
                import xlrd
                wb = xlrd.open_workbook(excel_path)
                return wb.sheet_names()
            except:
                pass
            # 最后尝试pandas
            import pandas as pd
            xl = pd.ExcelFile(excel_path)
            return xl.sheet_names


def process_sheets(excel_path, sheets=None, output_dir=None, merge=False):
    """
    处理指定的sheet，生成Word报告

    参数:
        excel_path: Excel文件路径
        sheets: 要处理的sheet列表，可以是：
                - None: 处理所有sheet
                - int: 单个sheet索引
                - str: 单个sheet名称
                - list: [0, 1, 2] 或 ["Sheet1", "Sheet2"]
        output_dir: 输出目录，None表示与Excel同目录
        merge: 是否合并多个sheet到一个Word文件

    返回:
        生成的Word文件路径列表
    """
    # 获取所有sheet名称
    all_sheets = list_sheets(excel_path)
    print(f"Excel包含 {len(all_sheets)} 个sheet: {all_sheets}")

    # 确定要处理的sheet
    sheets_to_process = _resolve_sheets(all_sheets, sheets)
    
    if not sheets_to_process:
        return []

    print(f"将处理 {len(sheets_to_process)} 个sheet: {sheets_to_process}")

    # 确定输出目录
    if output_dir is None:
        output_dir = os.path.dirname(excel_path)

    # 合并模式
    if merge and len(sheets_to_process) > 1:
        return _merge_sheets_to_word(excel_path, sheets_to_process, output_dir)
    
    # 单独生成模式
    return _generate_separate_reports(excel_path, sheets_to_process, output_dir)


def _resolve_sheets(all_sheets, sheets):
    """解析要处理的sheet列表"""
    sheets_to_process = []

    if sheets is None:
        sheets_to_process = all_sheets
    elif isinstance(sheets, int):
        if 0 <= sheets < len(all_sheets):
            sheets_to_process = [all_sheets[sheets]]
        else:
            print(f"错误: sheet索引 {sheets} 超出范围 (0-{len(all_sheets)-1})")
    elif isinstance(sheets, str):
        if sheets in all_sheets:
            sheets_to_process = [sheets]
        else:
            print(f"错误: 未找到名为 '{sheets}' 的sheet")
    elif isinstance(sheets, list):
        for s in sheets:
            if isinstance(s, int):
                if 0 <= s < len(all_sheets):
                    sheets_to_process.append(all_sheets[s])
                else:
                    print(f"警告: sheet索引 {s} 超出范围，跳过")
            elif isinstance(s, str):
                if s in all_sheets:
                    sheets_to_process.append(s)
                else:
                    print(f"警告: 未找到名为 '{s}' 的sheet，跳过")
    else:
        print(f"错误: 不支持的sheets参数类型: {type(sheets)}")

    return list(dict.fromkeys(sheets_to_process))


def _generate_separate_reports(excel_path, sheets_to_process, output_dir):
    """为每个sheet生成单独的Word报告"""
    output_files = []
    base_name = os.path.splitext(os.path.basename(excel_path))[0]

    for sheet_name in sheets_to_process:
        print(f"\n{'='*50}")
        print(f"正在处理sheet: {sheet_name}")
        print('='*50)

        # 生成输出文件名
        if len(sheets_to_process) == 1:
            word_path = os.path.join(output_dir, f"{base_name}_报告.docx")
        else:
            word_path = os.path.join(output_dir, f"{base_name}_{sheet_name}_报告.docx")

        try:
            converter = ExcelToWordReport(excel_path, word_path)
            converter.load_excel(sheet_name)
            converter.parse_test_cases()
            output_path = converter.generate_word_report()
            output_files.append(output_path)
        except Exception as e:
            print(f"处理sheet '{sheet_name}' 时出错: {e}")
            import traceback
            traceback.print_exc()

    return output_files


def _merge_sheets_to_word(excel_path, sheets_to_process, output_dir):
    """将多个sheet合并到一个Word文件"""
    from docx import Document
    from docx.oxml.ns import qn
    
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    word_path = os.path.join(output_dir, f"{base_name}_合并报告.docx")
    
    # 创建合并文档
    doc = Document()
    doc.styles['Normal'].font.name = '宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    
    # 添加总标题
    title = doc.add_heading(f'{base_name} 测试报告', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    
    print(f"\n{'='*50}")
    print(f"合并模式: 将 {len(sheets_to_process)} 个sheet合并到一个Word")
    print('='*50)
    
    for idx, sheet_name in enumerate(sheets_to_process):
        print(f"\n正在处理sheet [{idx+1}/{len(sheets_to_process)}]: {sheet_name}")
        
        try:
            # 创建临时转换器解析数据
            converter = ExcelToWordReport(excel_path)
            converter.load_excel(sheet_name)
            converter.parse_test_cases()
            
            # 添加分页符（第一个sheet不分页）
            if idx > 0:
                doc.add_page_break()
            
            # 添加sheet标题
            doc.add_heading(f'{idx+1} {sheet_name}', level=1)
            
            # 添加概述
            doc.add_heading(f'{idx+1}.1 概述', level=2)
            for field in ['产品信息', '试验信息', '工作模式', '测试仪器设备']:
                doc.add_paragraph(f"{field}: {converter.overview_data.get(field, '（待填写）')}")
            
            # 添加试验结果汇总
            if converter.summary_data:
                doc.add_heading(f'{idx+1}.2 试验结果汇总', level=2)
                summary_table = doc.add_table(rows=len(converter.summary_data) + 1, cols=4)
                summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                set_table_border(summary_table)
                
                headers = ['序号', '试验分类', '试验项目', '测试结论']
                for i, header in enumerate(headers):
                    cell = summary_table.cell(0, i)
                    cell.text = header
                    set_cell_font(cell, bold=True)
                
                for row_idx, item in enumerate(converter.summary_data):
                    for col_idx, header in enumerate(headers):
                        cell = summary_table.cell(row_idx + 1, col_idx)
                        cell.text = str(item.get(header, ''))
                        set_cell_font(cell)
                
                # 合并相同试验分类的单元格
                if len(converter.summary_data) > 1:
                    current_category = None
                    merge_start = 1
                    
                    for row_idx, item in enumerate(converter.summary_data, 1):
                        category = item.get('试验分类', '')
                        
                        if category != current_category:
                            if current_category is not None and row_idx > merge_start + 1:
                                summary_table.cell(merge_start, 1).merge(summary_table.cell(row_idx - 1, 1))
                            current_category = category
                            merge_start = row_idx
                    
                    if merge_start < len(converter.summary_data):
                        summary_table.cell(merge_start, 1).merge(summary_table.cell(len(converter.summary_data), 1))
            
            # 添加测试数据
            if converter.big_cases:
                doc.add_heading(f'{idx+1}.3 测试数据', level=2)
                for big_idx, big_case in enumerate(converter.big_cases):
                    clean_big_name = clean_case_number(big_case["name"])
                    doc.add_heading(f'{idx+1}.3.{big_idx+1} {clean_big_name}', level=3)
                    for small_idx, small_case in enumerate(big_case['small_cases']):
                        clean_small_name = clean_case_number(small_case["name"])
                        doc.add_heading(f'{idx+1}.3.{big_idx+1}.{small_idx+1} {clean_small_name}', level=4)
                        create_testcase_table(doc, small_case['data'])
            
            print(f"  ✓ {sheet_name} 处理完成")
            
        except Exception as e:
            print(f"  ✗ 处理sheet '{sheet_name}' 时出错: {e}")
            import traceback
            traceback.print_exc()
    
    # 保存文档
    doc.save(word_path)
    print(f"\n{'='*50}")
    print(f"合并报告已生成: {word_path}")
    
    return [word_path]


def main():
    """主函数"""
    # ===== 配置区域 =====
    excel_file = r"D:\AI\test_data.xlsx"  # Excel输入文件路径
    output_dir = None  # 输出目录，None表示与Excel同目录

    # 要处理的sheet配置:
    # 方式1: 处理所有sheet
    # sheets = None

    # 方式2: 处理单个sheet（索引或名称）
    # sheets = 0  # 第一个sheet
    # sheets = "Sheet1"  # 按名称

    # 方式3: 处理多个sheet（指定索引或名称）
    # sheets = [0, 1, 2]  # 按索引
    # sheets = ["Sheet1", "Sheet2"]  # 按名称
    # sheets = [0, "Sheet2", 2]  # 混合

    sheets = None  # 默认处理所有sheet
    
    # 合并模式配置:
    # True: 多个sheet合并到一个Word文件
    # False: 每个sheet生成单独的Word文件
    merge = False
    # ====================

    # 检查文件是否存在
    if not os.path.exists(excel_file):
        print(f"错误: Excel文件不存在 - {excel_file}")
        print("请修改 excel_file 配置为实际文件路径")
        return

    # 先列出所有sheet
    all_sheets = list_sheets(excel_file)
    print(f"\n可用sheet列表:")
    for i, name in enumerate(all_sheets):
        print(f"  [{i}] {name}")

    # 处理sheet
    output_files = process_sheets(excel_file, sheets, output_dir, merge)

    # 输出结果
    print(f"\n{'='*50}")
    print(f"转换完成！共生成 {len(output_files)} 个报告:")
    for f in output_files:
        print(f"  - {f}")


if __name__ == '__main__':
    main()
