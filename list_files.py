#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

def list_files():
    base_dir = r'D:\月报自动化'
    
    print("月报自动化目录结构:")
    print("=" * 50)
    
    # 列出原始数据目录
    raw_data_dir = os.path.join(base_dir, '月报原始数据')
    if os.path.exists(raw_data_dir):
        print(f"\n📁 {raw_data_dir}")
        for f in os.listdir(raw_data_dir):
            print(f"  - {f}")
    else:
        print(f"\n❌ 目录不存在: {raw_data_dir}")
    
    # 列出模板目录
    template_dir = os.path.join(base_dir, '月报模板')
    if os.path.exists(template_dir):
        print(f"\n📁 {template_dir}")
        for f in os.listdir(template_dir):
            print(f"  - {f}")
    else:
        print(f"\n❌ 目录不存在: {template_dir}")

if __name__ == '__main__':
    list_files()
