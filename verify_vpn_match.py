#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 找到最新司法局文件
output_dir = r'D:\月报自动化\输出月报'
files = [f for f in os.listdir(output_dir) if '司法局' in f and f.endswith('.docx') and not f.startswith('~$')]
files.sort(key=lambda f: os.path.getmtime(os.path.join(output_dir, f)), reverse=True)
output_file = os.path.join(output_dir, files[0])

print(f"验证文件: {files[0]}\n")
print("="*70)
print("✅ 验证 VPN 按组织名匹配")
print("="*70)

doc = Document(output_file)

# 查找VPN审计记录表
vpn_table = None
for i, table in enumerate(doc.tables):
    if len(table.columns) >= 6:
        header = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
        if '用户名' in header and '用户组' in header:
            vpn_table = table
            print(f"\n找到 VPN 审计记录表 (表格{i}):")
            print(f"  列数: {len(table.columns)}, 行数: {len(table.rows)}")
            break

if vpn_table:
    headers = [cell.text.strip() for cell in vpn_table.rows[0].cells]
    print(f"\n表头: {headers}")
    
    # 统计用户组
    user_groups = {}
    unique_users = set()
    
    for row_idx in range(1, len(vpn_table.rows)):
        cells = [cell.text.strip() for cell in vpn_table.rows[row_idx].cells]
        if len(cells) >= 2:
            user_group = cells[1]  # 用户组列
            username = cells[0]    # 用户名列
            unique_users.add(username)
            user_groups[user_group] = user_groups.get(user_group, 0) + 1
    
    print(f"\n✅ 统计结果:")
    print(f"  总记录数: {len(vpn_table.rows) - 1} 条")
    print(f"  唯一用户数: {len(unique_users)} 个")
    print(f"\n  用户组分布:")
    for group, count in sorted(user_groups.items(), key=lambda x: x[1], reverse=True)[:10]:
        print(f"    {group}: {count} 条")
    
    print(f"\n  前5行数据:")
    for row_idx in range(1, min(6, len(vpn_table.rows))):
        cells = [cell.text.strip() for cell in vpn_table.rows[row_idx].cells]
        print(f"    {cells}")
    
    # 验证匹配
    has_sifa = any('司法局' in group for group in user_groups.keys())
    status = '是' if has_sifa else '否'
    print(f"\n  包含司法局记录: {status}")
    
else:
    print("❌ 未找到 VPN 审计记录表")

print("\n" + "="*70)
