#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r'C:\Users\Admin\.openclaw\workspace\skills\monthly-report-automation\scripts')

from monthly_report import main, setup_data_paths, generate_report, TEMPLATE_DIR, OUTPUT_DIR
import os
from datetime import datetime
import re

# 设置参数
template_name = '政务云服务运维月报-2025年11月-北京市农林科学院.docx'
month = '2026-03'

# 设置数据路径
setup_data_paths(None, month)

# 确定路径
template_path = os.path.join(TEMPLATE_DIR, template_name)
timestamp = datetime.now().strftime('%H%M%S')
output_name = re.sub(r'\d{4}年\d{1,2}月', f'{month[:4]}年{int(month[5:7])}月_v2_{timestamp}', template_name)
output_path = os.path.join(OUTPUT_DIR, output_name)

print(f"模板: {template_path}")
print(f"输出: {output_path}")
print(f"存在: {os.path.exists(template_path)}")

if os.path.exists(template_path):
    success = generate_report(template_path, output_path)
    print(f"\n结果: {'成功' if success else '失败'}")
    if success and os.path.exists(output_path):
        print(f"文件大小: {os.path.getsize(output_path)} bytes")
else:
    print("模板不存在!")
