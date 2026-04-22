#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
filename = '政务云服务运维月报-2026年03月-北京市农林科学院.docx'
target = os.path.join(desktop, filename)

if os.path.exists(target):
    size = os.path.getsize(target)
    print(f"文件已在桌面: {filename}")
    print(f"文件大小: {size} bytes")
else:
    print("文件尚未复制到桌面")
    # 列出桌面文件
    if os.path.exists(desktop):
        files = [f for f in os.listdir(desktop) if '月报' in f]
        if files:
            print("桌面相关文件:")
            for f in files:
                print(f"  - {f}")
