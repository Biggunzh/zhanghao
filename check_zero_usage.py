#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import xlrd

book = xlrd.open_workbook(r'D:\月报自动化\月报原始数据\2026-03月报资源使用率详情列表.xls')
sheet = book.sheet_by_index(0)

# 找到表头
headers = sheet.row_values(0)
col_cpu_usage = 7  # CPU使用率/月/AVG
col_mem_usage = 9  # 内存使用率/月/AVG
col_disk_usage = 11  # 磁盘使用率/月/AVG
col_system = 1  # 业务系统名称

# 统计北京市行政执法信息服务平台的所有主机
target_system = '北京市行政执法信息服务平台'
hosts = []

for i in range(1, sheet.nrows):
    row = sheet.row_values(i)
    system_name = str(row[col_system]).strip()
    if system_name == target_system:
        cpu_usage = row[col_cpu_usage] if col_cpu_usage < len(row) else 0
        mem_usage = row[col_mem_usage] if col_mem_usage < len(row) else 0
        disk_usage = row[col_disk_usage] if col_disk_usage < len(row) else 0
        host_name = str(row[2]).strip()[:30]  # 云主机名称
        
        # 转换使用率
        try:
            cpu_usage = float(cpu_usage) if cpu_usage else 0
        except:
            cpu_usage = 0
        try:
            mem_usage = float(mem_usage) if mem_usage else 0
        except:
            mem_usage = 0
        try:
            disk_usage = float(disk_usage) if disk_usage else 0
        except:
            disk_usage = 0
        
        hosts.append({
            'name': host_name,
            'cpu_usage': cpu_usage,
            'mem_usage': mem_usage,
            'disk_usage': disk_usage
        })

print(f'业务系统: {target_system}')
print(f'主机总数: {len(hosts)}')
print()

# 统计有多少台主机的使用率为0
cpu_zero = sum(1 for h in hosts if h['cpu_usage'] == 0)
mem_zero = sum(1 for h in hosts if h['mem_usage'] == 0)
disk_zero = sum(1 for h in hosts if h['disk_usage'] == 0)

print(f'CPU使用率为0的主机: {cpu_zero} 台')
print(f'内存使用率为0的主机: {mem_zero} 台')
print(f'磁盘使用率为0的主机: {disk_zero} 台')
print()

# 列出使用率为0的主机
print('CPU使用率为0的主机:')
for h in hosts:
    if h['cpu_usage'] == 0:
        print(f'  - {h["name"]}')
print()

# 计算两种方式的使用率
cpu_values_all = [h['cpu_usage'] for h in hosts]
cpu_values_nonzero = [h['cpu_usage'] for h in hosts if h['cpu_usage'] > 0]

print('CPU使用率计算对比:')
print(f'  方式1 - 包含0值: sum({len(cpu_values_all)}个值) / {len(cpu_values_all)} = {sum(cpu_values_all)/len(cpu_values_all):.2f}%')
print(f'  方式2 - 排除0值: sum({len(cpu_values_nonzero)}个值) / {len(cpu_values_nonzero)} = {sum(cpu_values_nonzero)/len(cpu_values_nonzero):.2f}%')
