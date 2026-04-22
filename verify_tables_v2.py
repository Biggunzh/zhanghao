#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document
import sys
sys.stdout.reconfigure(encoding='utf-8')

output_file = r'D:\月报自动化\输出月报\政务云服务运维月报-2026年03月_v2_165738-北京市农林科学院.docx'

try:
    doc = Document(output_file)
    print("✅ 文档验证 - 业务系统资源使用情况统计表格\n")
    
    # 查找并显示CPU、内存、存储三个表格
    for i, table in enumerate(doc.tables[1:5]):  # 查看表格1-4
        if len(table.rows) < 2:
            continue
        
        # 获取表头
        header = ' '.join([cell.text.strip() for cell in table.rows[0].cells])
        
        if 'CPU' in header or '内存' in header or '存储' in header or '磁盘' in header:
            print(f"{'='*60}")
            print(f"表格{i+1}: {header[:30]}")
            print(f"{'='*60}")
            
            # 显示所有行
            for row_idx, row in enumerate(table.rows):
                cells = [cell.text.strip() for cell in row.cells]
                print(f"  行{row_idx}: {cells}")
            print()
    
except Exception as e:
    print(f"❌ 错误: {e}")
    import traceback
    traceback.print_exc()
