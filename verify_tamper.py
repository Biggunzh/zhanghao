#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document

doc = Document(r'D:\月报自动化\输出月报\政务云服务运维月报-2026年03_v2_194415-北京市公安局勤务指挥部.docx')

print("验证网页防篡改日期格式:")
print("="*70)

for i, table in enumerate(doc.tables):
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    header_str = ' '.join(headers)
    
    if '日期' in header_str and len(table.columns) == 3:
        print(f"\n表格{i} - 网页防篡改")
        print(f"表头: {headers}")
        # 查看前3行和最后1行
        for row_idx in [1, 2, 3, len(table.rows)-1]:
            if row_idx < len(table.rows):
                cells = table.rows[row_idx].cells
                tamp_date = cells[0].text.strip()
                status = cells[1].text.strip()
                monitor = cells[2].text.strip()
                print(f"  行{row_idx}: {tamp_date} | {status} | {monitor}")
