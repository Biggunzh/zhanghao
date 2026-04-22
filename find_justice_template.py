#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE_DIR = r'D:\月报自动化\月报模板'

# 查找包含"司法局"的模板
files = [f for f in os.listdir(TEMPLATE_DIR) if '司法' in f and f.endswith('.docx') and not f.startswith('~$')]

print(f"找到 {len(files)} 个司法局相关模板:")
for f in files:
    print(f"  - {f}")
