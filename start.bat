@echo off
chcp 65001 >nul
title 毅播快递对账系统

echo.
echo  ========================================
echo    毅播快递对账系统 启动中...
echo  ========================================
echo.

:: 切换到项目目录（自动获取bat文件所在目录）
cd /d "%~dp0"

:: 安装依赖（首次运行时自动安装flask）
pip install flask -q

echo  ✅ 服务启动成功！
echo  📌 请在浏览器打开以下地址：
echo.
echo     本机访问：http://127.0.0.1:5000
echo     局域网访问：查看下方IP地址
echo.

:: 显示本机局域网IP
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set IP=%%a
    echo     局域网地址：http:%%a:5000
)

echo.
echo  ⚠️  关闭此窗口将停止服务
echo  ========================================
echo.

:: 启动Flask服务
python app.py

pause