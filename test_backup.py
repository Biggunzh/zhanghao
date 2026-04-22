#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import calendar
from datetime import date

target_year, target_month = 2026, 3

# 计算3月的周五
cal = calendar.monthcalendar(target_year, target_month)
fridays = []
for week in cal:
    friday_day = week[4]
    if friday_day != 0:
        fridays.append(date(target_year, target_month, friday_day))

friday_strs = [f.strftime("%Y/%m/%d") for f in fridays]
print("2026年3月的周五:", friday_strs)
print("周五数量:", len(fridays))

# 假设5台主机
num_hosts = 5
total_records = num_hosts * len(fridays)
print(f"5台主机 x {len(fridays)}个周五 = {total_records}条备份记录")

# 生成备份记录示例
print("\n前3条备份记录示例:")
count = 0
for friday in fridays[:2]:
    for host_idx in range(3):
        backup_time = f"{friday.year}/{friday.month}/{friday.day} 22:00"
        print(f"  序号{host_idx+1}: {backup_time}")
        count += 1
        if count >= 3:
            break
    if count >= 3:
        break
