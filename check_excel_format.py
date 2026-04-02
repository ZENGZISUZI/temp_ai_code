# -*- coding: utf-8 -*-
"""
检测Excel文件实际格式
"""

import os
import struct


def detect_excel_format(file_path):
    """
    检测Excel文件的实际格式
    
    返回:
        'xlsx' - Excel 2007+ (Office Open XML)
        'xls' - Excel 97-2003 (BIFF)
        'unknown' - 未知格式
    """
    if not os.path.exists(file_path):
        return f"文件不存在: {file_path}"
    
    with open(file_path, 'rb') as f:
        header = f.read(8)
    
    # xlsx 实际是 ZIP 压缩包，以 PK 开头
    if header[:2] == b'PK':
        return 'xlsx (实际是ZIP/Office Open XML格式)'
    
    # xls 是 OLE 复合文档，以 D0 CF 11 E0 开头
    if header[:4] == b'\xD0\xCF\x11\xE0':
        return 'xls (实际是OLE/BIFF格式)'
    
    return f'unknown (文件头: {header.hex()})'


def check_with_pandas(file_path):
    """用pandas尝试读取"""
    import pandas as pd
    
    print(f"\n尝试用pandas读取...")
    
    # 尝试不同引擎
    engines = ['openpyxl', 'xlrd']
    
    for engine in engines:
        try:
            df = pd.read_excel(file_path, engine=engine, nrows=1)
            print(f"  ✓ 引擎 {engine} 成功，读取到 {len(df)} 行")
            return engine
        except Exception as e:
            print(f"  ✗ 引擎 {engine} 失败: {e}")
    
    return None


if __name__ == '__main__':
    # 配置文件路径
    file_path = r"D:\AI\test_data.xlsx"  # 改成你的文件路径
    
    print(f"检测文件: {file_path}")
    print(f"文件扩展名: {os.path.splitext(file_path)[1]}")
    print(f"文件大小: {os.path.getsize(file_path) / 1024:.2f} KB")
    
    # 检测实际格式
    actual_format = detect_excel_format(file_path)
    print(f"实际格式: {actual_format}")
    
    # 用pandas测试
    working_engine = check_with_pandas(file_path)
    
    if working_engine:
        print(f"\n建议使用引擎: {working_engine}")
