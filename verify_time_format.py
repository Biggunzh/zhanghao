#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 找到最新农林科文件
output_dir = r'D:\月报自动化\输出月报'
files = [f for f in os.listdir(output_dir) if '农林' in f and f.endswith('.docx') and not f.startswith('~$')]
files.sort(key=lambda f: os.path.getmtime(os.path.join(output_dir, f)), reverse=True)
output_file = os.path.join(output_dir, files[0])

print(f"验证文件: {files[0]}\n")
print("="*70)
print("✅ 验证堡垒机时间格式")
print("="*70)

doc = Document(output_file)

# 查找堡垒机表
fortress_table = None
for table in doc.tables:
    if len(table.columns) >= 8:
        header = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
        if '开始时间' in header and '资产IP' in header:
            fortress_table = table
            break

if fortress_table:
    print(f"\n找到堡垒机审计记录表:")
    print(f"  总行数: {len(fortress_table.rows)}行\n")
    
    print("前10行时间数据:")
    print("-"*70)
    for row_idx in range(1, min(11, len(fortress_table.rows))):
        cells = [cell.text.strip() for cell in fortress_table.rows[row_idx].cells]
        start_time = cells[0] if len(cells) > 0 else ''
        end_time = cells[1] if len(cells) > 1 else ''
        # 检查是否有微秒
        has_microsecond = '.' in start_time and len(start_time.split('.')[-1]) > 3
        status = '✗ 有微秒' if has_microsecond else '✓ 正常'
        print(f"  行{row_idx}: 开始={start_time[:19]:<20} 结束={end_time[:19]:<20} {status}")
    
    print("\n" + "="*70)
    print("✅ 时间格式验证完成！")
    print("="*70)
else:
    print("❌ 未找到堡垒机审计记录表")
