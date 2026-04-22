#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document

doc = Document(r'D:\月报自动化\输出月报\政务云服务运维月报-2026年03_v2_192203-北京市公安局勤务指挥部.docx')

print("检查快照备份表格:")
for i, table in enumerate(doc.tables):
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    header_str = ' '.join(headers)
    
    if '备份时间' in header_str:
        print(f"\n表格{i}:")
        print(f"  表头: {headers}")
        print(f"  总行数: {len(table.rows)} (包含表头)")
        print(f"  数据行数: {len(table.rows) - 1}")
        
        # 显示全部备份时间
        backup_times = []
        for row_idx in range(1, min(25, len(table.rows))):
            cells = table.rows[row_idx].cells
            if len(cells) > 3:
                time_text = cells[3].text.strip()
                backup_times.append(time_text)
        
        print(f"  前20行备份时间:")
        for idx, t in enumerate(backup_times[:20]):
            print(f"    行{idx+1}: {t}")
