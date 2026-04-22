#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 找到最新文件
output_dir = r'D:\月报自动化\输出月报'
files = [f for f in os.listdir(output_dir) if '农林' in f and f.endswith('.docx') and not f.startswith('~$')]
files.sort(key=lambda f: os.path.getmtime(os.path.join(output_dir, f)), reverse=True)
output_file = os.path.join(output_dir, files[0])

print(f"验证文件: {files[0]}")
print(f"文件大小: {os.path.getsize(output_file):,} bytes\n")

doc = Document(output_file)

# 查找快照备份表
snapshot_table = None
for i, table in enumerate(doc.tables):
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    header_str = ' '.join(headers)
    if ('云主机名称' in header_str or '主机名称' in header_str) and \
       ('备份时间' in header_str or '备份类型' in header_str):
        snapshot_table = table
        print(f"✅ 找到快照备份表（表格{i}）")
        print(f"  表头: {headers}")
        break

if snapshot_table:
    print(f"\n  总行数: {len(snapshot_table.rows)}行")
    print(f"  记录数: {len(snapshot_table.rows) - 1}条\n")
    
    # 统计备份时间
    backup_dates = {}
    for row_idx in range(1, len(snapshot_table.rows)):
        backup_time = snapshot_table.rows[row_idx].cells[3].text.strip()
        date = backup_time[:10] if backup_time else ''
        backup_dates[date] = backup_dates.get(date, 0) + 1
    
    print("备份日期分布:")
    for date, count in sorted(backup_dates.items()):
        print(f"  {date}: {count} 条")
    
    print(f"\n前3行数据:")
    for row_idx in range(1, min(4, len(snapshot_table.rows))):
        cells = [cell.text.strip() for cell in snapshot_table.rows[row_idx].cells]
        print(f"  {cells}")
    
    # 检查是否是3月份的日期
    is_march = all('2026-03' in date for date in backup_dates.keys() if date)
    print(f"\n✅ 备份日期都是2026年3月: {is_march}")
else:
    print("❌ 未找到快照备份表")
