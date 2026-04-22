#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import subprocess
import sys

result = subprocess.run(
    [sys.executable, r'D:\月报自动化\月报自动化_v2.py'],
    cwd=r'D:\月报自动化',
    capture_output=True,
    text=True
)

print("STDOUT:")
print(result.stdout[-3000:] if len(result.stdout) > 3000 else result.stdout)  # 只打印最后3000字符
print("\nReturn code:", result.returncode)
