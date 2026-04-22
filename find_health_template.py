#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE_DIR = r'D:\月报自动化\月报模板'

# 查找包含"健康"或"卫生"的模板
keywords = ['健康', '卫生']
files = []

for f in os.listdir(TEMPLATE_DIR):
    if f.endswith('.docx') and not f.startswith('~$'):
        for kw in keywords:
            if kw in f:
                files.append(f)
                break

print(f"找到 {len(files)} 个卫生健康相关模板:")
for f in files:
    print(f"  - {f}")
