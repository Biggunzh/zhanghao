#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
sys.path.insert(0, r'D:\月报自动化')
from 月报自动化_v2 import generate_report, TEMPLATE_DIR, OUTPUT_DIR
from datetime import datetime

test_template = '政务云服务运维月报-2025年11月-司法局.docx'
template_path = os.path.join(TEMPLATE_DIR, test_template)
timestamp = datetime.now().strftime('%H%M%S')
output_name = test_template.replace('2025年11月', f'2026年03月_fix_{timestamp}')
output_path = os.path.join(OUTPUT_DIR, output_name)

print(f"处理: {test_template}")
print(f"输出: {output_name}")

if os.path.exists(template_path):
    success = generate_report(template_path, output_path)
    if success and os.path.exists(output_path):
        size = os.path.getsize(output_path)
        print(f"\n✅ 成功: {output_path}")
        print(f"大小: {size} bytes")
    else:
        print(f"\n❌ 失败")
else:
    print(f"❌ 模板不存在: {template_path}")
