#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试处理司法局模板"""

import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 导入主脚本中的函数
sys.path.insert(0, r'D:\月报自动化')
from 月报自动化_v2 import generate_report

TEMPLATE_DIR = r'D:\月报自动化\月报模板'
OUTPUT_DIR = r'D:\月报自动化\输出月报'

# 司法局模板
test_template = '政务云服务运维月报-2025年11月-司法局.docx'
template_path = os.path.join(TEMPLATE_DIR, test_template)

# 输出文件名
from datetime import datetime
timestamp = datetime.now().strftime('%H%M%S')
output_name = test_template.replace('2025年11月', f'2026年03月_v2_{timestamp}')
output_path = os.path.join(OUTPUT_DIR, output_name)

print(f"处理模板: {test_template}")
print(f"输出文件: {output_name}")
print()

if os.path.exists(template_path):
    generate_report(template_path, output_path)
    print(f"\n✅ 完成！文件已保存到: {output_path}")
else:
    print(f"❌ 模板不存在: {template_path}")
