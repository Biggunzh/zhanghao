#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""读取Word文档并提取结构信息"""

import zipfile
from xml.etree import ElementTree as ET
import os

def get_text_from_docx(docx_path):
    """从docx文件中提取文本"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        xml_content = z.read('word/document.xml')
    
    tree = ET.fromstring(xml_content)
    
    # 命名空间
    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }
    
    # 提取文本，保留段落结构
    paragraphs = []
    for para in tree.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
        texts = []
        for elem in para.iter():
            if elem.tag.endswith('}t') and elem.text:
                texts.append(elem.text)
        if texts:
            paragraphs.append(''.join(texts))
    
    return '\n'.join(paragraphs)

if __name__ == '__main__':
    docx_path = r'D:\月报自动化\月报模板\政务云服务运维月报-2025年11月-北京市农林科学院.docx'
    print(f"Reading: {docx_path}")
    print("=" * 50)
    text = get_text_from_docx(docx_path)
    print(text[:5000])
