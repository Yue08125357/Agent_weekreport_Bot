@echo off
chcp 65001 >nul
title 周报助手 - 自动化服务

echo.
echo ========================================
echo   周报助手 - 自动化服务
echo ========================================
echo.

:: 检查 ngrok
where ngrok >nul 2>&1
if errorlevel 1 (
    echo [提示] 检测到未安装 ngrok
    echo.
    echo 请按以下步骤安装：
    echo 1. 访问 https://ngrok.com 注册账号（免费）
    echo 2. 下载 Windows 版本
    echo 3. 解压后运行 ngrok authtoken 你的token
    echo 4. 把 ngrok.exe 复制到 C:\Windows\System32\
    echo.
    echo 现在先启动本地服务...
    echo.
    goto :start_local
)

:: 启动本地服务（后台）
echo [1/2] 启动本地服务...
start /b python "%~dp0server.py"

:: 等待服务启动
timeout /t 5 /nobreak >nul

:: 启动 ngrok
echo [2/2] 启动内网穿透...
echo.
echo ========================================
echo   服务已启动！
echo ========================================
echo.
echo   本地地址: http://localhost:8000
echo.
echo   下方 Forwarding 后面的 https://xxx.ngrok.io
echo   就是你的公网地址，分享给同事即可
echo.
echo ========================================
echo.

ngrok http 8000
goto :end

:start_local
python "%~dp0server.py"

:end
pause
