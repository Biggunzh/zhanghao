#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证 VPN 用户名匹配逻辑"""
import sys
sys.path.insert(0, r'D:\月报自动化')

from 月报自动化_v2 import read_fortress_data, read_vpn_data_by_users, RAW_DATA_DIR, FORTRESS_FILE, VPN_FILE
import os

# 模拟主机IP
host_ips = {'192.168.178.226', '192.168.178.228', '192.169.230.6'}

print("="*70)
print("1. 读取堡垒机数据（按IP筛选）")
print("="*70)
fortress_records = read_fortress_data(FORTRESS_FILE, host_ips)
print(f"\n堡垒机记录数: {len(fortress_records)}")

# 提取用户名
vpn_users = set()
for record in fortress_records:
    user = record.get('user', '')
    if user and user.strip():
        vpn_users.add(user.strip())

print(f"\n从堡垒机提取的用户名:")
for u in sorted(vpn_users)[:10]:
    print(f"  - {u}")
print(f"  共 {len(vpn_users)} 个用户名")

print("\n" + "="*70)
print("2. 读取VPN审计数据（按用户名匹配）")
print("="*70)

if vpn_users:
    vpn_records = read_vpn_data_by_users(VPN_FILE, vpn_users)
    print(f"\nVPN匹配记录数: {len(vpn_records)}")
    
    print(f"\n前5条VPN记录:")
    for r in vpn_records[:5]:
        print(f"  用户: {r['username']}, 组: {r['user_group']}, 时间: {r['time']}")
else:
    print("\n未提取到用户名，跳过VPN匹配")
