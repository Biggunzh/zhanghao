#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

def inspect_excel(file_path, sheet_name=None):
    """查看Excel文件结构"""
    print(f"\n{'='*60}")
    print(f"📊 文件: {os.path.basename(file_path)}")
    print(f"{'='*60}")
    
    try:
        if file_path.endswith('.xls'):
            # 使用xlrd读取旧格式
            import xlrd
            book = xlrd.open_workbook(file_path)
            print(f"工作表: {book.sheet_names()}")
            for sheet_name in book.sheet_names()[:1]:  # 只看第一个表
                sheet = book.sheet_by_name(sheet_name)
                print(f"\n--- 工作表: {sheet_name} ---")
                print(f"行数: {sheet.nrows}, 列数: {sheet.ncols}")
                # 打印前10行
                for i in range(min(10, sheet.nrows)):
                    row = sheet.row_values(i)
                    print(f"行{i}: {row[:10]}")  # 只显示前10列
        else:
            # 使用pandas读取
            df = pd.read_excel(file_path, sheet_name=0)
            print(f"\n形状: {df.shape}")
            print(f"\n列名: {list(df.columns)}")
            print(f"\n前10行数据:")
            print(df.head(10).to_string())
    except Exception as e:
        print(f"错误: {e}")

if __name__ == '__main__':
    base_dir = r'D:\月报自动化\月报原始数据'
    
    # 检查各个文件
    files = [
        '2026-03月报资源使用率详情列表.xls',
        '2026-03工单总量.xlsx',
        '2026-03-堡垒机.xlsx',
        '2026-03vpn审计.xlsx',
    ]
    
    for f in files:
        file_path = os.path.join(base_dir, f)
        if os.path.exists(file_path):
            inspect_excel(file_path)
        else:
            print(f"❌ 文件不存在: {file_path}")
