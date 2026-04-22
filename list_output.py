#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

OUTPUT_DIR = r'D:\月报自动化\输出月报'

print("输出目录文件列表:")
print("="*70)

files = os.listdir(OUTPUT_DIR)
files.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_DIR, x)), reverse=True)

for f in files[:10]:
    if f.endswith('.docx'):
        size = os.path.getsize(os.path.join(OUTPUT_DIR, f))
        print(f"  - {f} ({size} bytes)")
