#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub部署准备脚本
"""
import os
import shutil

def create_deploy_package():
    # 源文件路径
    source_script = r'D:\月报自动化\月报自动化_v2.py'
    source_skill = r'C:\Users\Admin\.openclaw\workspace\skills\monthly-report-automation\SKILL.md'
    source_changelog = r'C:\Users\Admin\.openclaw\workspace\skills\monthly-report-automation\CHANGELOG.md'
    
    # 部署目录
    deploy_dir = r'D:\github-deploy\monthly-report-automation'
    scripts_dir = os.path.join(deploy_dir, 'scripts')
    
    # 创建目录
    os.makedirs(deploy_dir, exist_ok=True)
    os.makedirs(scripts_dir, exist_ok=True)
    
    print("=" * 60)
    print("🚀 准备GitHub部署包")
    print("=" * 60)
    
    # 复制主程序
    dest_script = os.path.join(scripts_dir, 'monthly_report.py')
    shutil.copy2(source_script, dest_script)
    print(f"✅ 复制: 月报自动化_v2.py -> scripts/monthly_report.py")
    
    # 复制skill文档
    shutil.copy2(source_skill, deploy_dir)
    print(f"✅ 复制: SKILL.md")
    
    shutil.copy2(source_changelog, deploy_dir)
    print(f"✅ 复制: CHANGELOG.md")
    
    # 创建README.md
    readme_content = """# 政务云月报自动化工具

政务云服务运维月报自动生成工具，支持从Excel/Word模板自动生成完整的月报文档。

## ✨ 功能特性

- 📊 自动读取资源使用率Excel数据
- 📝 智能匹配业务系统
- 📈 自动计算CPU/内存/存储使用率
- 🔒 安全审计记录（堡垒机/VPN）
- 💾 快照备份报告自动生成
- 🛡️ 网页防篡改服务报告
- 📅 自动日期替换和审核

## 🗂️ 输入文件要求

### 必需文件
- 资源使用率详情列表 (.xls)
- 月报模板 (.docx)

### 可选文件
- 工单列表 (.xlsx)
- 堡垒机审计记录 (.xlsx)
- VPN审计记录 (.xlsx/.csv)

## 🔧 使用方法

```bash
python scripts/monthly_report.py "模板文件.docx"
```

程序会自动：
1. 提取模板中的业务系统
2. 读取对应月份的资源数据
3. 生成完整的月报文档

## 📋 支持的报告内容

1. **资源使用情况统计** - CPU/内存/存储使用率
2. **安全审计记录** - 堡垒机/VPN审计
3. **服务报告** - 快照备份、网页防篡改
4. **技术支撑统计** - 工单数量和分类

## 🔄 版本历史

详见 [CHANGELOG.md](CHANGELOG.md)

## 📄 许可证

MIT License
"""
    
    readme_path = os.path.join(deploy_dir, 'README.md')
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print(f"✅ 创建: README.md")
    
    # 创建requirements.txt
    requirements = """python-docx>=0.8.11
openpyxl>=3.0.10
xlrd>=2.0.1
xlwt>=1.3.0
"""
    
    req_path = os.path.join(deploy_dir, 'requirements.txt')
    with open(req_path, 'w', encoding='utf-8') as f:
        f.write(requirements)
    print(f"✅ 创建: requirements.txt")
    
    # 创建.gitignore
    gitignore = """# Python
__pycache__/
*.py[cod]
*$py.class
.Python
*.so

# IDE
.vscode/
.idea/
*.swp

# 数据文件（敏感）
*.xlsx
*.xls
*.csv
!示例数据/

# 输出
输出月报/
output/

# 日志
*.log
"""
    
    gitignore_path = os.path.join(deploy_dir, '.gitignore')
    with open(gitignore_path, 'w', encoding='utf-8') as f:
        f.write(gitignore)
    print(f"✅ 创建: .gitignore")
    
    # 创建LICENSE
    license_content = """MIT License

Copyright (c) 2024

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
"""
    
    license_path = os.path.join(deploy_dir, 'LICENSE')
    with open(license_path, 'w', encoding='utf-8') as f:
        f.write(license_content)
    print(f"✅ 创建: LICENSE")
    
    print("=" * 60)
    print("✅ 部署包创建成功！")
    print(f"📂 位置: {deploy_dir}")
    print("=" * 60)
    print()
    print("下一步操作：")
    print()
    print("1. 在GitHub创建新仓库:")
    print("   https://github.com/new")
    print("   仓库名: monthly-report-automation")
    print()
    print("2. 在命令行执行：")
    print()
    print(f"   cd {deploy_dir}")
    print("   git init")
    print("   git add .")
    print('   git commit -m "🚀 v3.0 月报自动化工具"')
    print("   git remote add origin https://github.com/YOUR_USERNAME/monthly-report-automation.git")
    print("   git branch -M main")
    print("   git push -u origin main")
    print()
    print("=" * 60)

if __name__ == "__main__":
    create_deploy_package()
