@echo off
chcp 65001 >nul
title 周报助手 - 网页版

echo.
echo ========================================
echo   周报助手 - 网页版
echo ========================================
echo.
echo 启动中...
echo.
echo 访问地址: http://localhost:8000
echo 按 Ctrl+C 停止服务
echo.

python "%~dp0web_app.py"

pause
