@echo off
chcp 65001 >nul
echo ========================================
echo    每日论文追踪系统 - 一键部署脚本
echo ========================================
echo.

cd /d "%~dp0"

echo [1/5] 正在初始化 Git 仓库...
git init
if errorlevel 1 (
    echo ❌ Git 初始化失败！请确认 Git 已正确安装。
    pause
    exit /b 1
)
echo ✅ Git 初始化成功！
echo.

echo [2/5] 正在添加文件...
git add .
echo ✅ 文件添加成功！
echo.

echo [3/5] 正在提交...
git commit -m "Initial commit: Daily Paper Tracker"
if errorlevel 1 (
    echo ⚠️  提交失败或没有新文件，继续下一步...
)
echo ✅ 提交成功！
echo.

echo [4/5] 正在设置分支...
git branch -M main
echo ✅ 分支设置成功！
echo.

echo ========================================
echo.
echo ⚠️  接下来请手动执行以下步骤：
echo.
echo 1. 去 GitHub 创建仓库（如果还没创建）
echo 2. 然后在这个窗口继续输入：
echo    git remote add origin https://github.com/a-ka-k/你的仓库名.git
echo    git push -u origin main
echo.
echo ========================================
echo.
echo 按任意键退出...
pause >nul
