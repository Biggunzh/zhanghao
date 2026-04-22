#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r'D:\月报自动化')

# 测试快照备份功能
from 月报自动化_v2 import (
    get_fridays_from_month, 
    get_last_month, 
    generate_backup_records,
    read_resource_data
)

# 测试周五计算
print("="*60)
print("测试：计算上个月的周五")
print("="*60)

last_year, last_month = get_last_month(2026, 3)
print(f"目标月份：2026年3月")
print(f"上个月：{last_year}年{last_month}月")

fridays = get_fridays_from_month(last_year, last_month)
print(f"\n上个月（{last_year}-{last_month}）的周五：")
for f in fridays:
    print(f"  {f.strftime('%Y-%m-%d')}")

print(f"\n共 {len(fridays)} 个周五")

# 测试读取资源数据
print("\n" + "="*60)
print("测试：读取资源数据")
print("="*60)

target_systems = ['微营销', '长城网']
resource_data, host_ips = read_resource_data(
    r'D:\月报自动化\月报原始数据\2026-03月报资源使用率详情列表.xls',
    target_systems
)

print(f"\n业务系统数量：{len(resource_data)}")
total_hosts = 0
for system, data in resource_data.items():
    print(f"  {system}: {len(data['hosts'])} 台主机")
    total_hosts += len(data['hosts'])
    # 取第一个主机作为示例
    if data['hosts']:
        host = data['hosts'][0]
        print(f"    示例: {host['host_name']} ({host['ip']})")

print(f"\n主机总数：{total_hosts} 台")

# 测试生成备份记录
print("\n" + "="*60)
print("测试：生成快照备份记录")
print("="*60)

# 收集所有主机
all_hosts = []
for system_name in sorted(resource_data.keys()):
    data = resource_data[system_name]
    for host in data['hosts']:
        all_hosts.append({
            'host_name': host['host_name'],
            'ip': host['ip']
        })

backup_records = generate_backup_records(all_hosts, last_year, last_month, backup_person="张昊")

print(f"\n备份记录总数：{len(backup_records)} 条")
print(f"计算方式：{len(all_hosts)} 台主机 × {len(fridays)} 个周五 = {len(all_hosts) * len(fridays)} 条")

print("\n前5条备份记录示例：")
for record in backup_records[:5]:
    print(f"  {record['seq']}. {record['host_name']}")
    print(f"     IP: {record['ip']}, 时间: {record['backup_time']}")
    print(f"     类型: {record['backup_type']}, 负责人: {record['person']}")

print("\n" + "="*60)
print("测试完成！")
print("="*60)
