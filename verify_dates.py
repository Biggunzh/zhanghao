#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document

doc = Document(r'D:\月报自动化\输出月报\政务云服务运维月报-2026年03_v2_194056-北京市公安局勤务指挥部.docx')

print("检查快照备份和防篡改表的日期格式:")
print("="*70)

for i, table in enumerate(doc.tables):
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    header_str = ' '.join(headers)
    
    # 快照备份表格
    if '备份时间' in header_str and len(table.columns) >= 6:
        print(f"\n表格{i} - 快照备份 (共{len(table.rows)-1}行)")
        print(f"表头: {headers}")
        # 查看第1行和第20行
        for row_idx in [1, 20]:
            if row_idx < len(table.rows):
                cells = table.rows[row_idx].cells
                seq = cells[0].text.strip() if len(cells) > 0 else ''
                host = cells[1].text.strip()[:20] if len(cells) > 1 else ''
                backup_time = cells[3].text.strip() if len(cells) > 3 else ''
                print(f"  行{row_idx} (序号{seq}): {host}... | 备份时间: {backup_time}")
    
    # 网页防篡改表格
    if '日期' in header_str and len(table.columns) == 3:
        print(f"\n表格{i} - 网页防篡改 (共{len(table.rows)-1}行)")
        print(f"表头: {headers}")
        # 查看第1行和最后一行
        for row_idx in [1, len(table.rows)-1]:
            if row_idx < len(table.rows):
                cells = table.rows[row_idx].cells
                tamp_date = cells[0].text.strip()
                status = cells[1].text.strip()
                print(f"  行{row_idx}: 日期={tamp_date}, 状态={status}")
