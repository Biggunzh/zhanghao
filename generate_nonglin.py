#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import os
import re
from datetime import datetime

sys.path.insert(0, r'C:\Users\Admin\.openclaw\workspace\skills\monthly-report-automation\scripts')

from monthly_report import generate_report, setup_data_paths, TEMPLATE_DIR, OUTPUT_DIR

# 配置
month = '2026-03'
template_name = '政务云服务运维月报-2025年11月-北京市农林科学院.docx'

# 设置路径
setup_data_paths(None, month)

# 生成输出文件名
timestamp = datetime.now().strftime('%H%M%S')
output_name = re.sub(r'\d{4}年\d{1,2}月', f'{month[:4]}年{int(month[5:7])}月_标准_{timestamp}', template_name)

# 完整路径
template_path = os.path.join(TEMPLATE_DIR, template_name)
output_path = os.path.join(OUTPUT_DIR, output_name)

print("="*60)
print("🚀 生成北京市农林科学院标准月报")
print("="*60)
print(f"\n模板: {template_path}")
print(f"输出: {output_path}")

# 执行生成
success = generate_report(template_path, output_path)

if success and os.path.exists(output_path):
    size = os.path.getsize(output_path)
    print(f"\n✅ 生成成功!")
    print(f"文件: {output_name}")
    print(f"大小: {size:,} bytes")
else:
    print(f"\n❌ 生成失败")
