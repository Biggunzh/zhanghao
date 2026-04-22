#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

output_dir = r'D:\月报自动化\输出月报'
files = [f for f in os.listdir(output_dir) if '农林' in f and f.endswith('.docx') and not f.startswith('~$')]
files.sort(key=lambda f: os.path.getmtime(os.path.join(output_dir, f)), reverse=True)
doc = Document(os.path.join(output_dir, files[0]))

print(f"文件: {files[0]}\n")

# 查找VPN表
for i, table in enumerate(doc.tables):
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    if '用户名' in headers and '用户组' in headers:
        print(f"✅ VPN审计表（表格{i}）:")
        print(f"  总行数: {len(table.rows)}行")
        print(f"  记录数: {len(table.rows)-1}条")
        
        # 统计用户名
        users = {}
        for row_idx in range(1, len(table.rows)):
            user = table.rows[row_idx].cells[0].text.strip()
            users[user] = users.get(user, 0) + 1
        
        print(f"\n  唯一用户数: {len(users)}")
        print(f"\n  用户名分布 (前10):")
        for u, c in sorted(users.items(), key=lambda x: x[1], reverse=True)[:10]:
            print(f"    {u}: {c}条")
        
        print(f"\n  前3行数据:")
        for row_idx in range(1, min(4, len(table.rows))):
            cells = [cell.text.strip() for cell in table.rows[row_idx].cells]
            print(f"    {cells}")
        break
