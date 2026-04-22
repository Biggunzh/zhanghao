#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 检查月报自动化_v2.py中所有包含output_path的行

with open(r'D:\月报自动化\月报自动化_v2.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

print("包含'output_path'的行:")
for i, line in enumerate(lines, 1):
    if 'output_path' in line:
        print(f"{i}: {line.rstrip()}")
