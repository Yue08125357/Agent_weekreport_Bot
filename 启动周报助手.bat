@echo off
chcp 65001 >nul
title 周报自动汇总助手

echo.
echo ========================================
echo   周报自动汇总助手 - 启动程序
echo ========================================
echo.

:: 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python，请先安装Python
    pause
    exit /b 1
)

:: 检查依赖
echo [1/2] 检查依赖...
pip show playwright >nul 2>&1 || pip install playwright -q
pip show python-docx >nul 2>&1 || pip install python-docx -q

:: 运行主程序（使用系统Chrome）
echo [2/2] 启动周报助手...
echo.
echo 注意: 将使用你已安装的 Chrome 浏览器
echo.
python "%~dp0weekreport_bot.py"

pause
