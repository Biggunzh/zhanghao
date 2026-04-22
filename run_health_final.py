#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import subprocess
import sys
import os

# 模板名称
template = '政务云服务运维月报-2025年11月-北京市卫生健康人力资源发展中心.docx'

# 运行主脚本
result = subprocess.run(
    [sys.executable, r'D:\月报自动化\月报自动化_v2.py'],
    cwd=r'D:\月报自动化',
    capture_output=True,
    text=True,
    encoding='utf-8'
)

print("STDOUT:")
print(result.stdout)
print("\nSTDERR:")
print(result.stderr)
print(f"\nReturn code: {result.returncode}")
