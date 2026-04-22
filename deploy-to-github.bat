@echo off
chcp 65001 >nul

REM 创建GitHub部署目录
mkdir D:\github-deploy\monthly-report-automation
cd D:\github-deploy\monthly-report-automation

REM 复制核心文件
xcopy /Y "D:\月报自动化\月报自动化_v2.py" "scripts\" 2>nul || (
    mkdir scripts
    copy /Y "D:\月报自动化\月报自动化_v2.py" "scripts\monthly_report.py"
)

REM 复制skill文件
copy /Y "C:\Users\Admin\.openclaw\workspace\skills\monthly-report-automation\SKILL.md" .
copy /Y "C:\Users\Admin\.openclaw\workspace\skills\monthly-report-automation\CHANGELOG.md" .

REM 创建必要的文件
echo # 政务云月报自动化工具 > README.md
echo. >> README.md
echo 政务云服务运维月报自动生成工具 >> README.md
echo. >> README.md
echo ## 使用方法 >> README.md
echo ```bash >> README.md
echo python 月报自动化_v2.py "模板文件.docx" >> README.md
echo ``` >> README.md

echo python-docx>=0.8.11 > requirements.txt
echo openpyxl>=3.0.10 >> requirements.txt
echo xlrd>=2.0.1 >> requirements.txt
echo xlwt>=1.3.0 >> requirements.txt

echo __pycache__/ > .gitignore
echo *.pyc >> .gitignore
echo *.pyd >> .gitignore
echo *.log >> .gitignore
echo output/ >> .gitignore
echo D:\月报自动化\输出月报/ >> .gitignore
echo *.xlsx >> .gitignore
echo *.xls >> .gitignore

echo MIT License > LICENSE

cd D:\github-deploy\monthly-report-automation

REM 初始化git
git init
git add .
git commit -m "🚀 v3.0 月报自动化工具 Initial commit"

echo.
echo ================================
echo ✅ 部署包已创建在:
echo D:\github-deploy\monthly-report-automation
echo.
echo 下一步:
echo 1. 在GitHub创建新仓库: monthly-report-automation
echo 2. 运行以下命令推送:
echo.
echo git remote add origin https://github.com/YOUR_USERNAME/monthly-report-automation.git
echo git branch -M main
echo git push -u origin main
echo.
echo ================================
pause
