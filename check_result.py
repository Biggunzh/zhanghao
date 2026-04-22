#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document
import os

output_dir = r'D:\月报自动化\输出月报'
files = [f for f in os.listdir(output_dir) if '农林' in f and f.endswith('.docx')]
files.sort(key=lambda f: os.path.getmtime(os.path.join(output_dir, f)), reverse=True)

if not files:
    print("未找到文件")
    exit()

fp = os.path.join(output_dir, files[0])
doc = Document(fp)

print(f"文件: {files[0]}")
print(f"大小: {os.path.getsize(fp):,} bytes")
print()

for i, table in enumerate(doc.tables):
    headers = [c.text.strip() for c in table.rows[0].cells]
    h = ' '.join(headers)
    
    if '备份类型' in h:
        bt = table.rows[1].cells[headers.index('备份类型')].text
        print(f'表格{i} 快照备份类型: {bt}')
    elif '日期' in h and '防篡改' in h:
        d = table.rows[1].cells[0].text
        print(f'表格{i} 防篡改日期格式: {d}')
        d2 = table.rows[2].cells[0].text
        print(f'         第二行: {d2}')
        d31 = table.rows[31].cells[0].text
        print(f'         最后一行: {d31}')
    elif '用户名' in h:
        print(f'表格{i} VPN审计: {len(table.rows)-1} 条')

print()
print('== 农林科2026年3月月报生成完成 ==')
