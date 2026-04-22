#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document

doc = Document(r'D:\月报自动化\输出月报\政务云服务运维月报-2026年03_v2_192203-北京市公安局勤务指挥部.docx')

output = []
output.append('检查生成的Word文件中的日期格式:')
output.append('='*70)

for i, table in enumerate(doc.tables):
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    header_str = ' '.join(headers)
    
    # 快照备份表格
    if '备份时间' in header_str and len(table.columns) >= 6:
        output.append(f'\n表格{i} - 快照备份 (共{len(table.rows)}行)')
        # 查看后10行（新的数据在行尾）
        rows_to_check = min(10, len(table.rows) - 1)
        for row_idx in range(len(table.rows) - rows_to_check, len(table.rows)):
            if row_idx > 0:
                cells = table.rows[row_idx].cells
                backup_time = cells[3].text.strip() if len(cells) > 3 else ''
                output.append(f'  行{row_idx}: {backup_time}')
    
    # 网页防篡改表格
    if '日期' in header_str and '防篡改' in header_str:
        output.append(f'\n表格{i} - 网页防篡改 (共{len(table.rows)}行)')
        # 查看后5行
        rows_to_check = min(5, len(table.rows) - 1)
        for row_idx in range(len(table.rows) - rows_to_check, len(table.rows)):
            if row_idx > 0:
                cells = table.rows[row_idx].cells
                tamp_date = cells[0].text.strip()
                output.append(f'  行{row_idx}: {tamp_date}')

print('\n'.join(output))
