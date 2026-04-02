# -*- coding: utf-8 -*-
"""
xlsx 转 csv 转换工具
依赖: pip install pandas openpyxl
"""

import pandas as pd
import os


def xlsx_to_csv(xlsx_path, csv_path=None, sheet_name=0, encoding='utf-8-sig'):
    """
    将 xlsx 文件转换为 csv 格式
    
    参数:
        xlsx_path: xlsx 文件路径
        csv_path: csv 输出路径，默认与xlsx同目录同名
        sheet_name: 要读取的工作表，默认第一个(0)，也可以传工作表名称
        encoding: csv编码，默认utf-8-sig(兼容Excel打开)
    
    返回:
        csv文件路径
    """
    # 读取xlsx
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    
    # 如果是字典(多个sheet)，取第一个
    if isinstance(df, dict):
        df = list(df.values())[0]
    
    # 生成csv路径
    if csv_path is None:
        csv_path = os.path.splitext(xlsx_path)[0] + '.csv'
    
    # 保存为csv
    df.to_csv(csv_path, index=False, encoding=encoding)
    
    return csv_path


def batch_convert(folder_path, encoding='utf-8-sig'):
    """
    批量转换文件夹下所有xlsx文件
    
    参数:
        folder_path: 文件夹路径
        encoding: csv编码
    """
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') and not filename.startswith('~$'):
            xlsx_path = os.path.join(folder_path, filename)
            try:
                csv_path = xlsx_to_csv(xlsx_path, encoding=encoding)
                print(f"✓ 转换成功: {filename} -> {os.path.basename(csv_path)}")
            except Exception as e:
                print(f"✗ 转换失败: {filename}, 错误: {e}")


if __name__ == '__main__':
    # ===== 配置区域 =====
    # 单文件转换
    xlsx_file = r"D:\AI\test.xlsx"  # 输入xlsx文件路径
    csv_file = None  # 输出csv路径，None表示自动生成
    
    # 批量转换（注释掉单文件转换后使用）
    # folder = r"D:\AI"  # 要转换的文件夹
    
    # 编码设置
    file_encoding = 'utf-8-sig'  # utf-8-sig 兼容Excel，也可用 'gbk' 中文
    # ====================
    
    # 单文件转换
    if os.path.exists(xlsx_file):
        result = xlsx_to_csv(xlsx_file, csv_file, encoding=file_encoding)
        print(f"转换完成: {result}")
    else:
        print(f"文件不存在: {xlsx_file}")
    
    # 批量转换（取消注释使用）
    # if os.path.isdir(folder):
    #     batch_convert(folder, encoding=file_encoding)
    # else:
    #     print(f"文件夹不存在: {folder}")
