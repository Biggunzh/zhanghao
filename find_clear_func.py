#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 查找clear_and_set_cell函数定义和调用的行号

with open(r'D:\月报自动化\月报自动化_v2.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

print("查找 clear_and_set_cell 函数定义和调用:")
for i, line in enumerate(lines, 1):
    if 'clear_and_set_cell' in line:
        print(f"{i}: {line.rstrip()}")
