#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r'D:\月报自动化')

# 模拟 read_resource_data 函数
import xlrd
from collections import defaultdict

def read_resource_data(file_path, target_systems=None):
    """读取资源使用率数据，根据业务系统筛选"""
    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(0)
    
    print(f"读取资源数据: {file_path}")
    print(f"总行数: {sheet.nrows}")
    
    # 找到表头行
    header_row = None
    headers = []
    for i in range(min(15, sheet.nrows)):
        row_values = sheet.row_values(i)
        if any('业务系统' in str(cell) for cell in row_values):
            header_row = i
            headers = row_values
            break
    
    # 找到关键列索引
    col_idx = {}
    for i, h in enumerate(headers):
        h_str = str(h).strip()
        if '业务系统名称' in h_str:
            col_idx['system'] = i
        elif '云主机名称' in h_str:
            col_idx['host_name'] = i
        elif '浮动IP' in h_str:
            col_idx['ip'] = i
        elif h_str == 'CPU':
            col_idx['cpu'] = i
        elif '内存' in h_str and 'GB' in h_str:
            col_idx['memory'] = i
        elif '磁盘' in h_str and '总' in h_str:
            col_idx['storage'] = i
        elif 'CPU使用率' in h_str and 'AVG' in h_str:
            col_idx['cpu_usage'] = i
        elif '内存使用率' in h_str and 'AVG' in h_str:
            col_idx['mem_usage'] = i
        elif '磁盘使用率' in h_str and 'AVG' in h_str:
            col_idx['disk_usage'] = i
    
    print(f"列索引: {col_idx}\n")
    
    # 读取并筛选数据
    systems_data = defaultdict(lambda: {
        'hosts': [],
        'cpu_count': 0,
        'memory_gb': 0,
        'storage_gb': 0,
        'host_count': 0,
        'cpu_usage_values': [],
        'mem_usage_values': [],
        'disk_usage_values': []
    })
    
    for i in range(header_row + 1, sheet.nrows):
        row = sheet.row_values(i)
        if not row or len(row) < 3:
            continue
        
        system_name = str(row[col_idx.get('system', 1)]).strip()
        
        if not system_name:
            continue
        if target_systems and system_name not in target_systems:
            continue
        
        def get_val(idx, default=0):
            try:
                if idx and idx < len(row):
                    v = row[idx]
                    if isinstance(v, (int, float)):
                        return v
                    if str(v).strip():
                        return float(v)
            except:
                pass
            return default
        
        host_info = {
            'host_name': str(row[col_idx.get('host_name', 3)]).strip()[:30],
            'ip': str(row[col_idx.get('ip', 9)]).strip() if col_idx.get('ip', 9) < len(row) else '',
            'cpu': int(get_val(col_idx.get('cpu', 5))),
            'memory': int(get_val(col_idx.get('memory', 6))),
            'storage': int(get_val(col_idx.get('storage', 7))),
            'cpu_usage': get_val(col_idx.get('cpu_usage')),
            'mem_usage': get_val(col_idx.get('mem_usage')),
            'disk_usage': get_val(col_idx.get('disk_usage')),
        }
        
        systems_data[system_name]['hosts'].append(host_info)
        systems_data[system_name]['cpu_count'] += host_info['cpu']
        systems_data[system_name]['memory_gb'] += host_info['memory']
        systems_data[system_name]['storage_gb'] += host_info['storage']
        systems_data[system_name]['host_count'] += 1
        if host_info['cpu_usage'] > 0:
            systems_data[system_name]['cpu_usage_values'].append(host_info['cpu_usage'])
        if host_info['mem_usage'] > 0:
            systems_data[system_name]['mem_usage_values'].append(host_info['mem_usage'])
        if host_info['disk_usage'] > 0:
            systems_data[system_name]['disk_usage_values'].append(host_info['disk_usage'])
    
    # 计算平均使用率
    for sys_name, data in systems_data.items():
        data['cpu_usage_avg'] = sum(data['cpu_usage_values']) / len(data['cpu_usage_values']) if data['cpu_usage_values'] else 0
        data['mem_usage_avg'] = sum(data['mem_usage_values']) / len(data['mem_usage_values']) if data['mem_usage_values'] else 0
        data['disk_usage_avg'] = sum(data['disk_usage_values']) / len(data['disk_usage_values']) if data['disk_usage_values'] else 0
    
    return dict(systems_data)

# 测试
target_systems = ['北京市行政执法信息服务平台']
resource_data = read_resource_data(
    r'D:\月报自动化\月报原始数据\2026-03月报资源使用率详情列表.xls',
    target_systems
)

print('\n=== 详细计算验证 ===')
for system_name in sorted(resource_data.keys()):
    data = resource_data[system_name]
    print(f'\n业务系统: {system_name}')
    print(f'  主机数: {data["host_count"]}')
    print(f'  CPU总量: {data["cpu_count"]} 核')
    print(f'  内存总量: {data["memory_gb"]} GB')
    print(f'  存储总量: {data["storage_gb"]} GB')
    
    # 显示前5台主机的详细数据
    print(f'\n  前5台主机详情:')
    for i, host in enumerate(data['hosts'][:5]):
        print(f'    {i+1}. {host["host_name"]}')
        print(f'       CPU: {host["cpu"]}核, 使用率: {host["cpu_usage"]:.2f}%')
        print(f'       内存: {host["memory"]}GB, 使用率: {host["mem_usage"]:.2f}%')
        print(f'       存储: {host["storage"]}GB, 使用率: {host["disk_usage"]:.2f}%')
    
    print(f'\n  平均使用率计算:')
    if data['cpu_usage_values']:
        cpu_sum = sum(data['cpu_usage_values'])
        cpu_count = len(data['cpu_usage_values'])
        print(f'    CPU: {cpu_sum:.2f} / {cpu_count} = {data["cpu_usage_avg"]:.2f}%')
    if data['mem_usage_values']:
        mem_sum = sum(data['mem_usage_values'])
        mem_count = len(data['mem_usage_values'])
        print(f'    内存: {mem_sum:.2f} / {mem_count} = {data["mem_usage_avg"]:.2f}%')
    if data['disk_usage_values']:
        disk_sum = sum(data['disk_usage_values'])
        disk_count = len(data['disk_usage_values'])
        print(f'    磁盘: {disk_sum:.2f} / {disk_count} = {data["disk_usage_avg"]:.2f}%')
