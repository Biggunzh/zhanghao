#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import shutil
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 源文件（输出结果）
source = r'D:\月报自动化\输出月报\政务云服务运维月报-2026年03月-北京市农林科学院.docx'

# 目标路径（桌面）
desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
target = os.path.join(desktop, '政务云服务运维月报-2026年03月-北京市农林科学院.docx')

if os.path.exists(source):
    try:
        shutil.copy2(source, target)
        print(f"文件已复制到桌面: {target}")
        print("注意：这是一个生成的副本，原始模板文件未被修改")
    except Exception as e:
        print(f"复制失败: {e}")
else:
    print(f"源文件不存在: {source}")
