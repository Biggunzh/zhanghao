#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from docx import Document
import sys
sys.stdout.reconfigure(encoding='utf-8')

output_file = r'D:\月报自动化\输出月报\政务云服务运维月报-2026年03月-北京市农林科学院.docx'

try:
    doc = Document(output_file)
    print("✅ 文档打开成功！")
    
    # 查找基础资源台账概况段落
    print("\n查找基础资源台账概况:")
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if '基础资源' in text or '北京市农林科学院' in text and '运行' in text:
            print(f"\n段落 {i}:")
            print(text)
            # 检查是否重复
            if text.count('北京市农林科学院') > 1:
                print("\n⚠️ 警告: 文本有重复!")
            break
    
except Exception as e:
    print(f"❌ 错误: {e}")
