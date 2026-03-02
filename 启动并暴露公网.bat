@echo off
chcp 65001 >nul
title 周报助手 - 内网穿透版

echo.
echo ========================================
echo   周报助手 - 内网穿透版
echo ========================================
echo.

:: 检查 ngrok
where ngrok >nul 2>&1
if errorlevel 1 (
    echo [提示] 需要先安装 ngrok
    echo.
    echo 1. 访问 https://ngrok.com 注册账号
    echo 2. 下载 Windows 版本
    echo 3. 解压后把 ngrok.exe 放到系统 PATH 中
    echo.
    pause
    exit /b 1
)

:: 启动本地服务
echo [1/2] 启动本地服务...
start /b python "%~dp0web_app.py"

:: 等待服务启动
timeout /t 3 /nobreak >nul

:: 启动 ngrok
echo [2/2] 启动内网穿透...
echo.
echo ========================================
echo   以下地址可以分享给同事访问
echo ========================================
echo.

ngrok http 8000
