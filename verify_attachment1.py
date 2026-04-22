#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证附件1更新结果"""
from docx import Document
import sys
sys.stdout.reconfigure(encoding='utf-8')

output_file = r'D:\月报自动化\输出月报\政务云服务运维月报-2026年03月_v2_171643-北京市农林科学院.docx'

doc = Document(output_file)

print("="*70)
print("✅ 验证附件1 - 政务云基础资源台账")
print("="*70)

# 查找附件1表格
for i, table in enumerate(doc.tables[4:8]):
    if len(table.columns) >= 6:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        if '序号' in headers and '业务系统' in str(headers):
            print(f"\n附件1表格 (表格{i+4}):")
            print(f"  表头: {headers}")
            print(f"  总行数: {len(table.rows)} (含表头)")
            print(f"  数据行数: {len(table.rows) - 1}")
            
            print("\n  前10行数据:")
            for row_idx in range(1, min(11, len(table.rows))):
                cells = [cell.text.strip() for cell in table.rows[row_idx].cells]
                print(f"    行{row_idx}: {cells[:7]}")
            
            print("\n  最后5行数据:")
            for row_idx in range(max(1, len(table.rows)-5), len(table.rows)):
                cells = [cell.text.strip() for cell in table.rows[row_idx].cells]
                print(f"    行{row_idx}: {cells[:7]}")
            
            # 统计各业务系统主机数
            print("\n  业务系统统计:")
            system_count = {}
            for row_idx in range(1, len(table.rows)):
                cells = [cell.text.strip() for cell in table.rows[row_idx].cells]
                if len(cells) > 1:
                    sys_name = cells[1]
                    system_count[sys_name] = system_count.get(sys_name, 0) + 1
            
            for sys_name, count in system_count.items():
                print(f"    - {sys_name}: {count}台")
            
            break

print("\n" + "="*70)
print("✅ 验证完成！")
print("="*70)
