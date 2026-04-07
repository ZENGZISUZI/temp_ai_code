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
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import re
from datetime import datetime


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


def set_cell_font(cell, font_name='微软雅黑', font_size=10.5, bold=False, align_center=True):
    """
    设置单元格字体和对齐方式
    
    参数:
        cell: 单元格对象
        font_name: 字体名称（默认微软雅黑）
        font_size: 字体大小（默认五号10.5pt）
        bold: 是否加粗
        align_center: 是否居中对齐
    """
    # 设置字体
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        
        # 设置段落水平居中
        if align_center:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 设置单元格垂直居中
    if align_center:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


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


def add_header_footer(doc, report_name, report_number, logo_path=None, company_name="公司"):
    """
    添加页眉和页脚（带横线分隔）
    
    参数:
        doc: Word文档对象
        report_name: 报告名称（文件名）
        report_number: 报告编号
        logo_path: Logo图片路径（可选）
        company_name: 公司名称（用于保密信息）
    """
    from docx.enum.section import WD_ORIENT
    
    # 获取文档的第一个节
    section = doc.sections[0]
    
    # ===== 设置页眉 =====
    header = section.header
    header.is_linked_to_previous = False
    
    # 删除默认段落（避免页眉上方出现回车符号）
    for para in header.paragraphs:
        p = para._element
        p.getparent().remove(p)
    
    # 创建页眉表格（2列：Logo+报告名 | 报告编号）
    header_table = header.add_table(rows=1, cols=2, width=Inches(7.5))
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 设置列宽（各占一半）
    header_table.columns[0].width = Inches(3.75)  # Logo + 报告名
    header_table.columns[1].width = Inches(3.75)  # 报告编号
    
    # 第1列：Logo + 报告名称（左对齐）
    cell_left = header_table.cell(0, 0)
    para_left = cell_left.paragraphs[0]
    para_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para_left.paragraph_format.space_before = Pt(0)
    para_left.paragraph_format.space_after = Pt(0)
    para_left.paragraph_format.line_spacing = 1.0
    
    # 添加Logo图片
    if logo_path and os.path.exists(logo_path):
        try:
            run_logo = para_left.add_run()
            run_logo.add_picture(logo_path, width=Cm(0.75), height=Cm(0.49))
            para_left.add_run("  ")
        except Exception as e:
            print(f"警告: 无法添加Logo图片 - {e}")
    
    # 添加报告名称
    run_name = para_left.add_run(report_name)
    run_name.font.name = '微软雅黑'
    run_name.font.size = Pt(9)
    run_name._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    
    # 第2列：报告编号（右对齐）
    cell_right = header_table.cell(0, 1)
    para_right = cell_right.paragraphs[0]
    para_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    para_right.paragraph_format.space_before = Pt(0)
    para_right.paragraph_format.space_after = Pt(0)
    para_right.paragraph_format.line_spacing = 1.0
    run_number = para_right.add_run(f"报告编号:{report_number}")
    run_number.font.name = '微软雅黑'
    run_number.font.size = Pt(9)
    run_number._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    
    # 设置表格只有底部边框（横线紧贴文字）
    set_table_border_with_bottom_line(header_table)
    
    # ===== 设置页脚 =====
    footer = section.footer
    footer.is_linked_to_previous = False
    
    # 删除默认段落（避免页脚上方出现回车符号）
    for para in footer.paragraphs:
        p = para._element
        p.getparent().remove(p)
    
    # 创建页脚表格（2列：保密信息 | 页码）
    footer_table = footer.add_table(rows=1, cols=2, width=Inches(7.5))
    footer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 设置列宽（保密信息列宽，页码列窄）
    footer_table.columns[0].width = Inches(5.0)  # 保密信息
    footer_table.columns[1].width = Inches(2.5)  # 页码
    
    # 第1列：保密信息（居中）
    cell_secret = footer_table.cell(0, 0)
    para_secret = cell_secret.paragraphs[0]
    para_secret.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_secret.paragraph_format.space_before = Pt(0)
    para_secret.paragraph_format.space_after = Pt(0)
    para_secret.paragraph_format.line_spacing = 1.0
    run_secret = para_secret.add_run(f"{company_name}保密信息，未经授权禁止扩散！")
    run_secret.font.name = '微软雅黑'
    run_secret.font.size = Pt(9)
    run_secret.font.bold = True
    run_secret._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    
    # 第2列：页码（右对齐）
    cell_page = footer_table.cell(0, 1)
    para_page = cell_page.paragraphs[0]
    para_page.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    para_page.paragraph_format.space_before = Pt(0)
    para_page.paragraph_format.space_after = Pt(0)
    para_page.paragraph_format.line_spacing = 1.0
    
    # 添加页码字段
    run_page = para_page.add_run("第 ")
    run_page.font.name = '微软雅黑'
    run_page.font.size = Pt(9)
    run_page._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    
    add_page_number_field(para_page, '微软雅黑', 9)
    
    run_page2 = para_page.add_run(" 页，共 ")
    run_page2.font.name = '微软雅黑'
    run_page2.font.size = Pt(9)
    run_page2._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    
    add_total_pages_field(para_page, '微软雅黑', 9)
    
    run_page3 = para_page.add_run(" 页")
    run_page3.font.name = '微软雅黑'
    run_page3.font.size = Pt(9)
    run_page3._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    
    # 设置表格只有顶部边框（横线紧贴文字）
    set_table_border_with_top_line(footer_table)
    
    # 删除表格后面可能自动添加的空段落
    for para in footer.paragraphs:
        if not para.text.strip():
            p = para._element
            p.getparent().remove(p)


def set_table_border_with_bottom_line(table):
    """
    设置表格只有底部边框线（用于页眉）
    
    参数:
        table: 表格对象
    """
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')

    # 只设置底部边框
    for border_name in ['top', 'left', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'nil')
        tblBorders.append(border)
    
    # 底部边框显示
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), '000000')
    tblBorders.append(bottom)

    tblPr.append(tblBorders)
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def set_table_border_with_top_line(table):
    """
    设置表格只有顶部边框线（用于页脚）
    
    参数:
        table: 表格对象
    """
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')

    # 只设置顶部边框
    for border_name in ['bottom', 'left', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'nil')
        tblBorders.append(border)
    
    # 顶部边框显示
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '6')
    top.set(qn('w:color'), '000000')
    tblBorders.append(top)

    tblPr.append(tblBorders)
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def set_table_border(table, show_border=True):
    """
    设置表格边框
    
    参数:
        table: 表格对象
        show_border: 是否显示边框，False则隐藏边框
    """
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')

    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        if show_border:
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:color'), '000000')
        else:
            border.set(qn('w:val'), 'nil')
        tblBorders.append(border)

    tblPr.append(tblBorders)
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def add_page_number_field(paragraph, font_name='微软雅黑', font_size=9):
    """添加当前页码字段"""
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.text = "PAGE"
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    
    # 设置字体
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def add_total_pages_field(paragraph, font_name='微软雅黑', font_size=9):
    """添加总页数字段"""
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.text = "NUMPAGES"
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    
    # 设置字体
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def add_watermark_to_docx(doc, watermark_text):
    """
    为文档添加斜向水印（从左下角到右上角）
    水印在底层，正文内容浮在上面
    
    参数:
        doc: Word文档对象
        watermark_text: 水印文字
    """
    from lxml import etree
    
    # VML和Office命名空间
    VML_NS = 'urn:schemas-microsoft-com:vml'
    OFFICE_NS = 'urn:schemas-microsoft-com:office:office'
    
    # 为每个节添加水印
    for section in doc.sections:
        header = section.header
        
        # 获取或创建header的XML元素
        header_elem = header._element
        
        # 创建水印段落
        watermark_para = OxmlElement('w:p')
        
        # 创建段落属性
        pPr = OxmlElement('w:pPr')
        watermark_para.append(pPr)
        
        # 创建VML形状（斜向水印）
        # 使用带命名空间的XML字符串
        shape_xml = f'''
        <v:shape xmlns:v="urn:schemas-microsoft-com:vml" 
                 xmlns:o="urn:schemas-microsoft-com:office:office"
                 id="Watermark" 
                 style="position:absolute;margin-left:0;margin-top:0;width:400pt;height:80pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-vertical:center;mso-position-horizontal-relative:page;mso-position-vertical-relative:page"
                 coordsize="21600,21600" 
                 allowincell="f" 
                 filled="t" 
                 stroked="f">
            <v:fill opacity="0.3" on="t"/>
            <v:textpath style="font-family:&quot;Arial&quot;;font-size:36pt" on="t" string="{watermark_text}"/>
        </v:shape>
        '''
        
        # 解析XML并添加到段落
        shape_elem = etree.fromstring(shape_xml)
        watermark_para.append(shape_elem)
        
        # 将水印段落添加到header开头
        header_elem.insert(0, watermark_para)


def add_heading_with_number(doc, text, level=1, font_config=None):
    """
    添加带编号的标题（使用Word标题样式，支持目录和导航窗格）
    
    参数:
        doc: Word文档对象
        text: 标题文字
        level: 标题级别（1=大标题，2=次标题，3=小标题）
        font_config: 字体配置字典
    """
    # 默认配置
    default_config = {
        'font_name': '微软雅黑',
        'title1_size': 16,
        'title1_bold': True,
        'title2_size': 12,
        'title2_bold': True,
        'title3_size': 12,
        'title3_bold': False,
    }
    config = default_config.copy()
    if font_config:
        config.update(font_config)
    
    # 使用Word标题样式（支持目录和导航窗格）
    heading = doc.add_heading(text, level=level)
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # 设置字体格式
    for run in heading.runs:
        run.font.name = config['font_name']
        run._element.rPr.rFonts.set(qn('w:eastAsia'), config['font_name'])
        
        if level == 1:
            run.font.size = Pt(config['title1_size'])
            run.font.bold = config['title1_bold']
        elif level == 2:
            run.font.size = Pt(config['title2_size'])
            run.font.bold = config['title2_bold']
        else:
            run.font.size = Pt(config['title3_size'])
            run.font.bold = config['title3_bold']
    
    return heading


def close_word_document(file_path):
    """
    关闭占用指定文件的Word文档
    
    参数:
        file_path: 文件路径
        
    返回:
        True: 成功关闭或文件未被占用
        False: 关闭失败
    """
    try:
        import win32com.client
        
        # 获取已运行的Word实例
        word = win32com.client.Dispatch("Word.Application")
        
        abs_path = os.path.abspath(file_path)
        
        # 遍历所有打开的文档
        for doc in word.Documents:
            try:
                # 比较文件路径
                doc_path = os.path.abspath(doc.FullName)
                if doc_path.lower() == abs_path.lower():
                    print(f"  检测到文件被占用，正在关闭: {os.path.basename(file_path)}")
                    doc.Close(SaveChanges=False)  # 不保存更改
                    print(f"  ✓ 已关闭占用文件")
                    break
            except:
                continue
                
        return True
    except Exception as e:
        # Word未运行或其他错误，文件未被占用
        return True


def update_toc_in_word(word_path):
    """
    使用win32com打开Word文档并更新目录
    
    参数:
        word_path: Word文档路径
    """
    try:
        import win32com.client
        
        # 使用DispatchEx创建独立的Word实例，避免影响用户已打开的Word
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False  # 后台运行
        
        try:
            # 打开文档
            doc = word.Documents.Open(os.path.abspath(word_path))
            
            # 更新所有域（包括目录）
            doc.Fields.Update()
            
            # 保存并关闭文档
            doc.Save()
            doc.Close()
            
            print("✓ 目录已自动更新")
            return True
        finally:
            # 退出独立的Word实例
            word.Quit()
            
    except ImportError:
        print("⚠️ 未安装win32com，无法自动更新目录")
        print("  提示: pip install pywin32")
        return False
    except Exception as e:
        print(f"⚠️ 更新目录失败: {e}")
        return False


def add_bookmark(paragraph, bookmark_name):
    """
    为段落添加书签
    
    参数:
        paragraph: 段落对象
        bookmark_name: 书签名称
    """
    tag = paragraph._p
    bookmark_start = OxmlElement('w:bookmarkStart')
    bookmark_start.set(qn('w:id'), str(hash(bookmark_name) % 10000))
    bookmark_start.set(qn('w:name'), bookmark_name)
    
    bookmark_end = OxmlElement('w:bookmarkEnd')
    bookmark_end.set(qn('w:id'), str(hash(bookmark_name) % 10000))
    
    tag.insert(0, bookmark_start)
    tag.append(bookmark_end)


def add_hyperlink(paragraph, text, bookmark_name, font_name='微软雅黑', font_size=10.5):
    """
    添加超链接到段落（指向书签）
    
    参数:
        paragraph: 段落对象
        text: 显示文字
        bookmark_name: 目标书签名称
        font_name: 字体名称
        font_size: 字体大小
    """
    # 创建超链接元素
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), bookmark_name)
    
    # 创建run元素
    new_run = OxmlElement('w:r')
    
    # 设置字体
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rPr.append(rFonts)
    
    # 设置字号
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(font_size * 2)))  # Word字号单位是半磅
    rPr.append(sz)
    
    new_run.append(rPr)
    
    # 设置文字
    text_elem = OxmlElement('w:t')
    text_elem.text = text
    new_run.append(text_elem)
    
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def setup_heading_styles(doc, font_config=None):
    """
    设置Word标题样式（确保导航窗格能识别）
    
    参数:
        doc: Word文档对象
        font_config: 字体配置字典
    """
    from docx.shared import RGBColor
    
    default_config = {
        'font_name': '微软雅黑',
        'title1_size': 16,
        'title1_bold': True,
        'title2_size': 12,
        'title2_bold': True,
        'title3_size': 12,
        'title3_bold': False,
    }
    config = default_config.copy()
    if font_config:
        config.update(font_config)
    
    # 设置 Heading 1 样式
    style1 = doc.styles['Heading 1']
    style1.font.name = config['font_name']
    style1.font.size = Pt(config['title1_size'])
    style1.font.bold = config['title1_bold']
    style1.font.color.rgb = RGBColor(0, 0, 0)  # 黑色
    style1._element.rPr.rFonts.set(qn('w:eastAsia'), config['font_name'])
    
    # 设置 Heading 2 样式
    style2 = doc.styles['Heading 2']
    style2.font.name = config['font_name']
    style2.font.size = Pt(config['title2_size'])
    style2.font.bold = config['title2_bold']
    style2.font.color.rgb = RGBColor(0, 0, 0)  # 黑色
    style2._element.rPr.rFonts.set(qn('w:eastAsia'), config['font_name'])
    
    # 设置 Heading 3 样式
    style3 = doc.styles['Heading 3']
    style3.font.name = config['font_name']
    style3.font.size = Pt(config['title3_size'])
    style3.font.bold = config['title3_bold']
    style3.font.color.rgb = RGBColor(0, 0, 0)  # 黑色
    style3._element.rPr.rFonts.set(qn('w:eastAsia'), config['font_name'])


def add_body_paragraph(doc, text, font_config=None):
    """
    添加正文段落
    
    参数:
        doc: Word文档对象
        text: 正文内容
        font_config: 字体配置字典
    """
    default_config = {
        'font_name': '微软雅黑',
        'body_size': 10.5,
        'body_bold': False,
    }
    config = default_config.copy()
    if font_config:
        config.update(font_config)
    
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    run = para.add_run(text)
    run.font.name = config['font_name']
    run._element.rPr.rFonts.set(qn('w:eastAsia'), config['font_name'])
    run.font.size = Pt(config['body_size'])
    run.font.bold = config['body_bold']
    
    return para


def format_date_only(value):
    """
    格式化日期，只保留日期部分（去掉时分秒）
    
    参数:
        value: 日期值（可能是字符串、datetime对象等）
        
    返回:
        只包含日期的字符串
    """
    if not value:
        return ''
    
    value_str = str(value).strip()
    
    # 尝试解析常见日期格式
    import re
    
    # 匹配 YYYY-MM-DD 或 YYYY/MM/DD 格式
    date_match = re.match(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', value_str)
    if date_match:
        return date_match.group(1).replace('/', '-')
    
    # 匹配 YYYYMMDD 格式
    date_match = re.match(r'(\d{4})(\d{2})(\d{2})', value_str)
    if date_match:
        return f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"
    
    # 如果是 datetime 对象
    try:
        from datetime import datetime
        if isinstance(value, datetime):
            return value.strftime('%Y-%m-%d')
    except:
        pass
    
    # 返回原始值
    return value_str


def create_testcase_table(doc, data_dict, font_config=None):
    """
    创建测试用例表格
    
    参数:
        doc: Word文档对象
        data_dict: 数据字典 {字段名: 值}
        font_config: 字体配置字典
    """
    # 默认配置
    default_config = {
        'font_name': '微软雅黑',
        'body_size': 10.5,
    }
    config = default_config.copy()
    if font_config:
        config.update(font_config)
    
    font_name = config['font_name']
    font_size = config['body_size']
    
    # 后续行字段（2列）
    remaining_fields = ['试验机构', '试验环境', '试验标准', '试验条件', '规格要求', '试验数据', '试验结论']
    
    # 计算总行数
    total_rows = 2 + len(remaining_fields)
    
    # 创建表格（最多4列）
    table = doc.add_table(rows=total_rows, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_border(table)
    
    # 第一行：开始日期 | 值 | 结束日期 | 值
    table.cell(0, 0).text = '开始日期'
    set_cell_font(table.cell(0, 0), font_name=font_name, font_size=font_size, bold=True)
    table.cell(0, 1).text = format_date_only(data_dict.get('开始日期', ''))
    set_cell_font(table.cell(0, 1), font_name=font_name, font_size=font_size)
    table.cell(0, 2).text = '结束日期'
    set_cell_font(table.cell(0, 2), font_name=font_name, font_size=font_size, bold=True)
    table.cell(0, 3).text = format_date_only(data_dict.get('结束日期', ''))
    set_cell_font(table.cell(0, 3), font_name=font_name, font_size=font_size)
    
    # 第二行：样机数量 | 值 | 样机编号 | 值
    table.cell(1, 0).text = '样机数量'
    set_cell_font(table.cell(1, 0), font_name=font_name, font_size=font_size, bold=True)
    table.cell(1, 1).text = str(data_dict.get('样机数量', ''))
    set_cell_font(table.cell(1, 1), font_name=font_name, font_size=font_size)
    table.cell(1, 2).text = '样机编号'
    set_cell_font(table.cell(1, 2), font_name=font_name, font_size=font_size, bold=True)
    table.cell(1, 3).text = str(data_dict.get('样机编号', ''))
    set_cell_font(table.cell(1, 3), font_name=font_name, font_size=font_size)
    
    # 第三行起：字段名占1列，值合并3列
    for i, field in enumerate(remaining_fields):
        row_idx = i + 2
        
        # 合并第2-4列（值占3列）
        table.cell(row_idx, 1).merge(table.cell(row_idx, 2)).merge(table.cell(row_idx, 3))
        
        # 填充内容
        table.cell(row_idx, 0).text = field
        set_cell_font(table.cell(row_idx, 0), font_name=font_name, font_size=font_size, bold=True)
        value = data_dict.get(field, '')
        table.cell(row_idx, 1).text = str(value) if value else ''
        set_cell_font(table.cell(row_idx, 1), font_name=font_name, font_size=font_size)

    doc.add_paragraph()
    return table


class ExcelToWordReport:
    """Excel转Word报告主类"""

    # 默认字体配置
    DEFAULT_FONT_CONFIG = {
        'font_name': '微软雅黑',
        'title1_size': 16,      # 大标题：三号
        'title1_bold': True,
        'title2_size': 12,      # 次标题：小四
        'title2_bold': True,
        'title3_size': 12,      # 小标题：小四
        'title3_bold': False,
        'body_size': 10.5,      # 正文：五号
        'body_bold': False,
    }

    def __init__(self, excel_path, word_path=None, logo_path=None, report_number=None, 
                 company_name="公司", watermark_text=None, report_name=None, font_config=None):
        """
        初始化

        参数:
            excel_path: Excel文件路径
            word_path: Word输出路径，默认同名
            logo_path: Logo图片路径（页眉用）
            report_number: 报告编号（页眉用），默认自动生成
            company_name: 公司名称（页脚保密信息用）
            watermark_text: 水印文字，如 "xxxx to xxxx"
            report_name: 报告名称（页眉用），默认使用文件名
            font_config: 字体配置字典，可覆盖默认配置
        """
        self.excel_path = excel_path
        self.word_path = word_path or os.path.splitext(excel_path)[0] + '_报告.docx'
        self.logo_path = logo_path
        self.report_number = report_number or self._generate_report_number()
        self.company_name = company_name
        self.watermark_text = watermark_text
        self.report_name = report_name
        
        # 合并字体配置
        self.font_config = self.DEFAULT_FONT_CONFIG.copy()
        if font_config:
            self.font_config.update(font_config)

        # 读取Excel
        self.df = None
        self.excel_columns = []

        # 解析后的数据
        self.overview_data = {}
        self.big_cases = []
        self.summary_data = []
        self.col_name_to_idx = {}
    
    def _generate_report_number(self):
        """生成默认报告编号（日期+时间）"""
        from datetime import datetime
        return datetime.now().strftime("RPT%Y%m%d%H%M%S")

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
        font_name = self.font_config.get('font_name', '微软雅黑')
        body_size = self.font_config.get('body_size', 10.5)
        
        doc.styles['Normal'].font.name = font_name
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        doc.styles['Normal'].font.size = Pt(body_size)
        
        # 设置标题样式（确保导航窗格能识别）
        setup_heading_styles(doc, self.font_config)

        # ===== 添加页眉页脚 =====
        report_name = self.report_name or os.path.splitext(os.path.basename(self.word_path))[0]
        add_header_footer(doc, report_name, self.report_number, self.logo_path, self.company_name)

        # ===== 添加水印 =====
        if self.watermark_text:
            add_watermark_to_docx(doc, self.watermark_text)

        # ===== 目录 =====
        # 目录标题（居中，不作为标题样式）
        toc_title = doc.add_paragraph()
        toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        toc_run = toc_title.add_run('目录')
        toc_run.font.name = font_name
        toc_run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        toc_run.font.size = Pt(self.font_config.get('title1_size', 16))
        toc_run.font.bold = self.font_config.get('title1_bold', True)
        toc_run.font.color.rgb = RGBColor(0, 0, 0)
        
        doc.add_paragraph()  # 空行
        
        # 生成目录项（带超链接）
        toc_items = [
            ('1 概述', 'toc_1', 1),
            ('1.1 产品信息', 'toc_1_1', 2),
            ('1.2 试验信息', 'toc_1_2', 2),
            ('1.3 工作模式', 'toc_1_3', 2),
            ('1.4 测试仪器设备', 'toc_1_4', 2),
            ('2 试验结果汇总', 'toc_2', 1),
            ('3 测试数据', 'toc_3', 1),
        ]
        
        # 添加大用例和小用例到目录
        for big_idx, big_case in enumerate(self.big_cases, 1):
            clean_name = clean_case_number(big_case["name"])
            toc_items.append((f'3.{big_idx} {clean_name}', f'toc_3_{big_idx}', 2))
            for small_idx, small_case in enumerate(big_case['small_cases'], 1):
                clean_small_name = clean_case_number(small_case["name"])
                toc_items.append((f'3.{big_idx}.{small_idx} {clean_small_name}', f'toc_3_{big_idx}_{small_idx}', 3))
        
        # 输出目录（带超链接和页码格式）
        for item_text, bookmark_name, level in toc_items:
            toc_para = doc.add_paragraph()
            
            # 根据级别添加缩进
            indent = '    ' * (level - 1)
            
            # 添加超链接标题
            add_hyperlink(toc_para, indent + item_text, bookmark_name, font_name, body_size)
            
            # 添加点号
            dots_run = toc_para.add_run(' ' + '.' * 40 + ' ')
            dots_run.font.name = font_name
            dots_run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            dots_run.font.size = Pt(body_size)
            
            # 添加页码占位符
            page_run = toc_para.add_run('1')
            page_run.font.name = font_name
            page_run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            page_run.font.size = Pt(body_size)
        
        doc.add_paragraph()  # 空行

        # ===== 1. 概述（大标题）=====
        h1 = add_heading_with_number(doc, '1 概述', level=1, font_config=self.font_config)
        add_bookmark(h1, 'toc_1')

        # 1.1 产品信息（次标题）
        h1_1 = add_heading_with_number(doc, '1.1 产品信息', level=2, font_config=self.font_config)
        add_bookmark(h1_1, 'toc_1_1')
        add_body_paragraph(doc, self.overview_data.get('产品信息', '（待填写）'), font_config=self.font_config)

        # 1.2 试验信息（次标题）
        h1_2 = add_heading_with_number(doc, '1.2 试验信息', level=2, font_config=self.font_config)
        add_bookmark(h1_2, 'toc_1_2')
        add_body_paragraph(doc, self.overview_data.get('试验信息', '（待填写）'), font_config=self.font_config)

        # 1.3 工作模式（次标题）
        h1_3 = add_heading_with_number(doc, '1.3 工作模式', level=2, font_config=self.font_config)
        add_bookmark(h1_3, 'toc_1_3')
        add_body_paragraph(doc, self.overview_data.get('工作模式', '（待填写）'), font_config=self.font_config)

        # 1.4 测试仪器设备（次标题）
        h1_4 = add_heading_with_number(doc, '1.4 测试仪器设备', level=2, font_config=self.font_config)
        add_bookmark(h1_4, 'toc_1_4')
        add_body_paragraph(doc, self.overview_data.get('测试仪器设备', '（待填写）'), font_config=self.font_config)

        # ===== 2. 试验结果汇总（大标题）=====
        h2 = add_heading_with_number(doc, '2 试验结果汇总', level=1, font_config=self.font_config)
        add_bookmark(h2, 'toc_2')

        # 创建汇总表格
        summary_table = doc.add_table(rows=len(self.summary_data) + 1, cols=4)
        summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_border(summary_table)
        
        # 设置列宽（2.67cm、6.4cm、5.75cm、2.67cm）
        from docx.shared import Cm
        summary_table.columns[0].width = Cm(2.67)   # 序号
        summary_table.columns[1].width = Cm(6.4)    # 试验分类
        summary_table.columns[2].width = Cm(5.75)   # 试验项目
        summary_table.columns[3].width = Cm(2.67)   # 测试结论

        # 表头
        headers = ['序号', '试验分类', '试验项目', '测试结论']
        for i, header in enumerate(headers):
            cell = summary_table.cell(0, i)
            cell.text = header
            set_cell_font(cell, font_name=font_name, font_size=body_size, bold=True)

        # 数据行
        for row_idx, item in enumerate(self.summary_data):
            for col_idx, header in enumerate(headers):
                cell = summary_table.cell(row_idx + 1, col_idx)
                cell.text = str(item.get(header, ''))
                set_cell_font(cell, font_name=font_name, font_size=body_size)
        
        # 调试输出
        print("\n试验结果汇总数据:")
        for idx, item in enumerate(self.summary_data):
            print(f"  行{idx+1}: 试验分类={item.get('试验分类', '')}, 试验项目={item.get('试验项目', '')}")
        
        # 合并相同试验分类的单元格并设置居中
        if len(self.summary_data) > 0:
            # 找出所有需要合并的范围
            merge_ranges = []  # [(start_row, end_row), ...]
            current_category = None
            merge_start = 1
            
            for row_idx, item in enumerate(self.summary_data, 1):
                category = item.get('试验分类', '')
                
                if category != current_category:
                    if current_category is not None:
                        merge_ranges.append((merge_start, row_idx - 1))
                    current_category = category
                    merge_start = row_idx
            
            # 添加最后一个分类
            if current_category is not None:
                merge_ranges.append((merge_start, len(self.summary_data)))
            
            # 执行合并
            print(f"\n合并范围: {merge_ranges}")
            for start, end in merge_ranges:
                if end > start:  # 至少2行才合并
                    print(f"  合并第{start}行到第{end}行")
                    
                    # 获取第一个单元格的值
                    first_cell = summary_table.cell(start, 1)
                    value = first_cell.text
                    
                    # 清空其他单元格的内容
                    for row_idx in range(start + 1, end + 1):
                        cell = summary_table.cell(row_idx, 1)
                        for para in cell.paragraphs:
                            para.clear()
                    
                    # 执行合并
                    merged_cell = first_cell.merge(summary_table.cell(end, 1))
                    
                    # 重新设置值（合并后可能丢失）
                    # 清空合并后的段落，重新添加一个居中的段落
                    for para in merged_cell.paragraphs:
                        para.clear()
                    
                    # 添加新的段落并设置居中
                    new_para = merged_cell.paragraphs[0]
                    new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = new_para.add_run(value)
                    run.font.name = font_name
                    run.font.size = Pt(body_size)
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                    
                    # 设置垂直居中
                    merged_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        doc.add_paragraph()

        # ===== 3. 测试数据 =====
        h3 = add_heading_with_number(doc, '3 测试数据', level=1, font_config=self.font_config)
        add_bookmark(h3, 'toc_3')
        
        # 验证大用例章节位置
        print("\n验证大用例章节位置...")
        validation_results = validate_chapter_position(self.big_cases)
        invalid_count = sum(1 for r in validation_results if not r['valid'])
        if invalid_count > 0:
            print(f"⚠️ 发现 {invalid_count} 个大用例位置不匹配")
        else:
            print("✓ 所有大用例位置正确")

        for big_idx, big_case in enumerate(self.big_cases, 1):
            clean_name = clean_case_number(big_case["name"])
            h3_big = add_heading_with_number(doc, f'3.{big_idx} {clean_name}', level=2, font_config=self.font_config)
            add_bookmark(h3_big, f'toc_3_{big_idx}')

            for small_idx, small_case in enumerate(big_case['small_cases'], 1):
                clean_small_name = clean_case_number(small_case["name"])
                h3_small = add_heading_with_number(doc, f'3.{big_idx}.{small_idx} {clean_small_name}', level=3, font_config=self.font_config)
                add_bookmark(h3_small, f'toc_3_{big_idx}_{small_idx}')
                create_testcase_table(doc, small_case['data'], font_config=self.font_config)

        # 保存文档
        # 先检查并关闭占用该文件的Word文档
        close_word_document(self.word_path)
        
        # 如果文件已存在，删除旧文件
        if os.path.exists(self.word_path):
            try:
                os.remove(self.word_path)
                print(f"  已删除旧文件: {os.path.basename(self.word_path)}")
            except Exception as e:
                print(f"  警告: 无法删除旧文件 - {e}")
        
        try:
            doc.save(self.word_path)
            print(f"Word报告已生成: {self.word_path}")
        except PermissionError:
            print(f"\n❌ 错误: 文件仍被占用，无法保存!")
            print(f"   请手动关闭 Word 中打开的 '{os.path.basename(self.word_path)}' 后重试")
            return None
        except Exception as e:
            print(f"\n❌ 保存失败: {e}")
            return None
        
        # 自动更新目录页码
        update_toc_in_word(self.word_path)
        
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


def process_sheets(excel_path, sheets=None, output_dir=None, merge=False, 
                   logo_path=None, report_number=None, company_name="公司", 
                   watermark_text=None, report_name=None, font_config=None):
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
        logo_path: Logo图片路径
        report_number: 报告编号
        company_name: 公司名称
        watermark_text: 水印文字
        report_name: 报告名称（页眉用）
        font_config: 字体配置字典

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
        return _merge_sheets_to_word(excel_path, sheets_to_process, output_dir, 
                                     logo_path, report_number, company_name, 
                                     watermark_text, report_name, font_config)
    
    # 单独生成模式
    return _generate_separate_reports(excel_path, sheets_to_process, output_dir,
                                      logo_path, report_number, company_name, 
                                      watermark_text, report_name, font_config)


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


def _generate_separate_reports(excel_path, sheets_to_process, output_dir,
                                logo_path=None, report_number=None, company_name="公司", 
                                watermark_text=None, report_name=None, font_config=None):
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
            converter = ExcelToWordReport(excel_path, word_path, logo_path, report_number, 
                                          company_name, watermark_text, report_name, font_config)
            converter.load_excel(sheet_name)
            converter.parse_test_cases()
            output_path = converter.generate_word_report()
            output_files.append(output_path)
        except Exception as e:
            print(f"处理sheet '{sheet_name}' 时出错: {e}")
            import traceback
            traceback.print_exc()

    return output_files


def _merge_sheets_to_word(excel_path, sheets_to_process, output_dir,
                          logo_path=None, report_number=None, company_name="公司", 
                          watermark_text=None, report_name=None, font_config=None):
    """将多个sheet合并到一个Word文件"""
    from docx import Document
    from docx.oxml.ns import qn
    
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    word_path = os.path.join(output_dir, f"{base_name}_合并报告.docx")
    
    # 获取字体配置
    fc = font_config or {}
    font_name = fc.get('font_name', '微软雅黑')
    body_size = fc.get('body_size', 10.5)
    
    # 创建合并文档
    doc = Document()
    doc.styles['Normal'].font.name = font_name
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    doc.styles['Normal'].font.size = Pt(body_size)
    
    # 添加页眉页脚
    actual_report_name = report_name or f"{base_name}_合并报告"
    actual_report_number = report_number or f"RPT{datetime.now().strftime('%Y%m%d%H%M%S')}"
    add_header_footer(doc, actual_report_name, actual_report_number, logo_path, company_name)
    
    # 添加水印
    if watermark_text:
        add_watermark_to_docx(doc, watermark_text)
    
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
                
                # 设置列宽（2.67cm、6.4cm、5.75cm、2.67cm）
                from docx.shared import Cm
                summary_table.columns[0].width = Cm(2.67)   # 序号
                summary_table.columns[1].width = Cm(6.4)    # 试验分类
                summary_table.columns[2].width = Cm(5.75)   # 试验项目
                summary_table.columns[3].width = Cm(2.67)   # 测试结论
                
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
                
                # 合并相同试验分类的单元格并居中
                if len(converter.summary_data) > 1:
                    merge_ranges = []
                    current_category = None
                    merge_start = 1
                    
                    for row_idx, item in enumerate(converter.summary_data, 1):
                        category = item.get('试验分类', '')
                        
                        if category != current_category:
                            if current_category is not None:
                                merge_ranges.append((merge_start, row_idx - 1))
                            current_category = category
                            merge_start = row_idx
                    
                    if current_category is not None:
                        merge_ranges.append((merge_start, len(converter.summary_data)))
                    
                    for start, end in merge_ranges:
                        if end > start:
                            first_cell = summary_table.cell(start, 1)
                            value = first_cell.text
                            
                            for row_idx in range(start + 1, end + 1):
                                cell = summary_table.cell(row_idx, 1)
                                for para in cell.paragraphs:
                                    para.clear()
                            
                            merged_cell = first_cell.merge(summary_table.cell(end, 1))
                            
                            for para in merged_cell.paragraphs:
                                para.clear()
                            
                            new_para = merged_cell.paragraphs[0]
                            new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run = new_para.add_run(value)
                            run.font.name = font_name
                            run.font.size = Pt(body_size)
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                            
                            merged_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            
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
    # 先检查并关闭占用该文件的Word文档
    close_word_document(word_path)
    
    # 如果文件已存在，删除旧文件
    if os.path.exists(word_path):
        try:
            os.remove(word_path)
            print(f"  已删除旧文件: {os.path.basename(word_path)}")
        except Exception as e:
            print(f"  警告: 无法删除旧文件 - {e}")
    
    try:
        doc.save(word_path)
        print(f"\n{'='*50}")
        print(f"合并报告已生成: {word_path}")
    except PermissionError:
        print(f"\n❌ 错误: 文件仍被占用，无法保存!")
        print(f"   请手动关闭 Word 中打开的 '{os.path.basename(word_path)}' 后重试")
        return []
    except Exception as e:
        print(f"\n❌ 保存失败: {e}")
        return []
    
    return [word_path]


def load_config(config_path):
    """
    从配置文件读取参数
    
    配置文件格式（每行一个配置，#开头为注释）：
    excel_file=D:\AI\test_data.xlsx
    logo_path=D:\AI\logo.png
    font_name=微软雅黑
    title1_size=16
    
    参数:
        config_path: 配置文件路径
        
    返回:
        配置字典
    """
    config = {
        'excel_file': r"D:\AI\test_data.xlsx",
        'output_dir': None,
        'logo_path': None,
        'report_number': None,
        'company_name': "公司",
        'report_name': None,
        'watermark_text': None,
        'sheets': None,
        'merge': False,
        # 字体配置
        'font_name': '微软雅黑',
        'title1_size': 16,
        'title1_bold': True,
        'title2_size': 12,
        'title2_bold': True,
        'title3_size': 12,
        'title3_bold': False,
        'body_size': 10.5,
        'body_bold': False,
    }
    
    if not os.path.exists(config_path):
        print(f"配置文件不存在: {config_path}")
        print("将使用默认配置")
        return config
    
    with open(config_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            
            # 跳过空行和注释
            if not line or line.startswith('#'):
                continue
            
            # 解析 key=value
            if '=' in line:
                key, value = line.split('=', 1)
                key = key.strip()
                value = value.strip()
                
                # 处理特殊值
                if value.lower() == 'none' or value == '':
                    config[key] = None
                elif value.lower() == 'true':
                    config[key] = True
                elif value.lower() == 'false':
                    config[key] = False
                elif key in ['title1_size', 'title2_size', 'title3_size', 'body_size']:
                    # 字体大小配置（支持小数）
                    try:
                        config[key] = float(value)
                    except:
                        config[key] = float(config.get(key, 10.5))
                elif key == 'sheets':
                    if ',' in value:
                        config[key] = [int(s.strip()) if s.strip().isdigit() else s.strip() for s in value.split(',')]
                    elif value.isdigit():
                        config[key] = int(value)
                    else:
                        config[key] = value
                else:
                    config[key] = value
    
    print(f"已从配置文件加载: {config_path}")
    return config


def create_default_config(config_path):
    """
    创建默认配置文件
    
    参数:
        config_path: 配置文件路径
    """
    default_config = '''# Excel转Word报告配置文件
# 每行一个配置，格式：key=value
# #开头的行为注释，会被忽略

# ===== 必填项 =====
# Excel输入文件路径
excel_file=D:\\AI\\test_data.xlsx

# ===== 页眉页脚配置 =====
# Logo图片路径（可选，None则不显示）
logo_path=None

# 报告编号（可选，None则自动生成）
report_number=None

# 公司名称（用于页脚保密信息）
company_name=公司

# 报告名称（页眉显示，None则使用文件名）
report_name=None

# ===== 水印配置 =====
# 水印文字（设置后会在页面中央显示灰色斜向水印）
# 示例：watermark_text=张三 to 李四
# 不需要水印则设为None
watermark_text=None

# ===== Sheet配置 =====
# 要处理的sheet（None=全部，数字=索引，字符串=名称，逗号分隔=多个）
sheets=None

# 合并模式（True=合并到一个Word，False=每个sheet单独生成）
merge=False

# ===== 字体配置 =====
# 字体名称（默认微软雅黑）
font_name=微软雅黑

# 大标题字号（默认16=三号）
title1_size=16
# 大标题是否加粗（true/false）
title1_bold=true

# 次标题字号（默认12=小四）
title2_size=12
# 次标题是否加粗（true/false）
title2_bold=true

# 小标题字号（默认12=小四）
title3_size=12
# 小标题是否加粗（true/false）
title3_bold=false

# 正文字号（默认10.5=五号）
body_size=10.5
# 正文是否加粗（true/false）
body_bold=false

# ===== 输出配置 =====
# 输出目录（None表示与Excel同目录）
output_dir=None
'''
    
    with open(config_path, 'w', encoding='utf-8') as f:
        f.write(default_config)
    
    print(f"已创建默认配置文件: {config_path}")


def main():
    """主函数"""
    # ===== 配置文件路径 =====
    config_path = r"D:\AI\report_config.txt"
    
    # 如果配置文件不存在，创建默认配置
    if not os.path.exists(config_path):
        create_default_config(config_path)
        print("\n请修改配置文件后重新运行：")
        print(f"  {config_path}")
        return
    
    # 从配置文件加载参数
    config = load_config(config_path)
    
    # 提取字体配置
    font_config = {
        'font_name': config.get('font_name', '微软雅黑'),
        'title1_size': config.get('title1_size', 16),
        'title1_bold': config.get('title1_bold', True),
        'title2_size': config.get('title2_size', 12),
        'title2_bold': config.get('title2_bold', True),
        'title3_size': config.get('title3_size', 12),
        'title3_bold': config.get('title3_bold', False),
        'body_size': config.get('body_size', 10.5),
        'body_bold': config.get('body_bold', False),
    }
    
    # 检查文件是否存在
    if not os.path.exists(config['excel_file']):
        print(f"错误: Excel文件不存在 - {config['excel_file']}")
        print("请修改配置文件中的 excel_file")
        return

    # 先列出所有sheet
    all_sheets = list_sheets(config['excel_file'])
    print(f"\n可用sheet列表:")
    for i, name in enumerate(all_sheets):
        print(f"  [{i}] {name}")

    # 处理sheet
    output_files = process_sheets(
        config['excel_file'], 
        config['sheets'], 
        config['output_dir'], 
        config['merge'], 
        logo_path=config['logo_path'], 
        report_number=config['report_number'], 
        company_name=config['company_name'],
        watermark_text=config['watermark_text'],
        report_name=config['report_name'],
        font_config=font_config
    )

    # 输出结果
    print(f"\n{'='*50}")
    print(f"转换完成！共生成 {len(output_files)} 个报告:")
    for f in output_files:
        print(f"  - {f}")


if __name__ == '__main__':
    main()
