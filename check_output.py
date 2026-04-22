#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os

output_dir = r'D:\月报自动化\输出月报'

if os.path.exists(output_dir):
    files = os.listdir(output_dir)
    print(f"输出目录存在: {output_dir}")
    print(f"文件数: {len(files)}")
    for f in files:
        print(f"  - {f}")
else:
    print(f"输出目录不存在: {output_dir}")
    print("脚本可能运行失败了...")
