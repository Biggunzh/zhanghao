#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import os

# 添加skill路径
sys.path.insert(0, r'C:\Users\Admin\.openclaw\workspace\skills\monthly-report-automation\scripts')

# 打印路径信息
print("Python路径:")
for p in sys.path:
    print(f"  {p}")

print(f"\n当前目录: {os.getcwd()}")

# 导入模块
try:
    import monthly_report
    print("\n✅ 模块导入成功")
    
    # 设置数据路径
    monthly_report.setup_data_paths(None, '2026-03')
    
    print(f"\n数据文件:")
    print(f"  资源: {monthly_report.RESOURCE_FILE}")
    print(f"  存在: {os.path.exists(monthly_report.RESOURCE_FILE)}")
    print(f"  工单: {monthly_report.WORKORDER_FILE}")
    print(f"  存在: {os.path.exists(monthly_report.WORKORDER_FILE)}")
    print(f"  堡垒机: {monthly_report.FORTRESS_FILE}")
    print(f"  存在: {os.path.exists(monthly_report.FORTRESS_FILE)}")
    print(f"  VPN: {monthly_report.VPN_FILE}")
    print(f"  存在: {os.path.exists(monthly_report.VPN_FILE)}")
    
    # 模板路径
    template_name = '政务云服务运维月报-2025年11月-北京市农林科学院.docx'
    template_path = os.path.join(monthly_report.TEMPLATE_DIR, template_name)
    print(f"\n模板:")
    print(f"  路径: {template_path}")
    print(f"  存在: {os.path.exists(template_path)}")
    
except Exception as e:
    print(f"\n❌ 错误: {e}")
    import traceback
    traceback.print_exc()
