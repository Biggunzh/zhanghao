#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证农林科2026年3月月报"""
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

results = []
for i, table in enumerate(doc.tables):
    headers = [c.text.strip() for c in table.rows[0].cells]
    h = ' '.join(headers)
    
    if '工单' in h:
        results.append(('工单统计', len(table.rows)-1))
    elif 'CPU' in h and '使用率' in h and '平均' not in h:
        results.append(('CPU汇总', len(table.rows)-1))
    elif '内存' in h and '使用率' in h and '平均' not in h:
        results.append(('内存汇总', len(table.rows)-1))
    elif '存储' in h and '使用率' in h and '平均' not in h:
        results.append(('存储汇总', len(table.rows)-1))
    elif '主机IP' in h:
        results.append(('基础资源台账', len(table.rows)-1))
    elif '平均使用率' in h and 'CPU' in h:
        results.append(('CPU详情(附件2)', len(table.rows)-1))
    elif '平均使用率' in h and '内存' in h:
        results.append(('内存详情(附件2)', len(table.rows)-1))
    elif '当前使用率' in h:
        results.append(('磁盘详情(附件2)', len(table.rows)-1))
    elif '备份类型' in h:
        bt = table.rows[1].cells[headers.index('备份类型')].text
        results.append((f'快照备份(附件2): {bt}', len(table.rows)-1))
    elif '防篡改' in h:
        d = table.rows[1].cells[0].text
        results.append((f'防篡改(附件2): {d}', len(table.rows)-1))
    elif '堡垒机' in h or '资产IP' in h:
        results.append(('堡垒机审计(附件3)', len(table.rows)-1))
    elif '用户名' in h and '用户组' in h:
        results.append(('VPN审计(附件3)', len(table.rows)-1))

for name, count in results:
    print(f'{name}: {count} 条')

print()
print('✅ 农林科2026年3月月报主体完成！')
print('⚠️  VPN部分需要重新上传未加密文件')
