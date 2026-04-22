#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""读取Word文档并提取结构信息"""

import zipfile
from xml.etree import ElementTree as ET
import os
import sys

# 设置输出编码
sys.stdout.reconfigure(encoding='utf-8')

def get_text_from_docx(docx_path):
    """从docx文件中提取文本"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        # Read document.xml
        xml_content = z.read('word/document.xml')
    
    # Parse XML
    tree = ET.fromstring(xml_content)
    
    # 命名空间
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # 提取所有文本，保留段落结构
    paragraphs = []
    for para in tree.findall('.//w:p', ns):
        texts = []
        for text_elem in para.findall('.//w:t', ns):
            if text_elem.text:
                texts.append(text_elem.text)
        if texts:
            paragraphs.append(''.join(texts))
    
    return '\n'.join(paragraphs)

if __name__ == '__main__':
    docx_path = r'D:\月报自动化\月报模板\政务云服务运维月报-2025年11月-北京市农林科学院.docx'
    if not os.path.exists(docx_path):
        # 尝试列出目录
        template_dir = r'D:\月报自动化\月报模板'
        if os.path.exists(template_dir):
            print(f"目录内容: {template_dir}")
            for f in os.listdir(template_dir):
                print(f"  - {f}")
        else:
            print(f"目录不存在: {template_dir}")
        sys.exit(1)
    
    print(f"Reading: {docx_path}")
    print("=" * 50)
    text = get_text_from_docx(docx_path)
    print(text[:5000])  # Print first 5000 chars
