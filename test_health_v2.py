#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试处理卫生健康人力资源发展中心模板 - v2"""

import os
import sys
import importlib.util
sys.stdout.reconfigure(encoding='utf-8')

# 导入主脚本中的函数
spec = importlib.util.spec_from_file_location("report_module", r'D:\月报自动化\月报自动化_v2.py')
report_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(report_module)
generate_report = report_module.generate_report

TEMPLATE_DIR = r'D:\月报自动化\月报模板'
OUTPUT_DIR = r'D:\月报自动化\输出月报'

# 卫生健康人力资源发展中心模板
test_template = '政务云服务运维月报-2025年11月-北京市卫生健康人力资源发展中心.docx'
template_path = os.path.join(TEMPLATE_DIR, test_template)

# 输出文件名
from datetime import datetime
timestamp = datetime.now().strftime('%H%M%S')
output_name = test_template.replace('2025年11月', f'2026年03月_final_{timestamp}')
output_path = os.path.join(OUTPUT_DIR, output_name)

print(f"="*70)
print(f"处理模板: {test_template}")
print(f"输出路径: {output_path}")
print(f"="*70)
print()

if os.path.exists(template_path):
    success = generate_report(template_path, output_path)
    if success and os.path.exists(output_path):
        size = os.path.getsize(output_path)
        print(f"\n{'='*70}")
        print(f"✅ 完成！")
        print(f"输出文件: {output_name}")
        print(f"文件大小: {size} bytes")
        print(f"{'='*70}")
    else:
        print(f"\n❌ 生成失败或文件未保存")
else:
    print(f"❌ 模板不存在: {template_path}")
