@echo off
chcp 65001 >nul
title Excel 拆分/合并工具

:: Excel 拆分/合并工具启动脚本 (Windows)

cd /d %~dp0

echo ================================================
echo   Excel 拆分/合并工具
echo ================================================
echo.

:: 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未检测到 Python
    echo 请先安装 Python 3.8 或更高版本
    echo 下载地址: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo ✓ Python 已安装
python --version
echo.

:: 检查依赖是否安装
echo 正在检查依赖...
python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo ⚠ 首次运行，正在安装依赖包...
    echo 这可能需要几分钟时间，请耐心等待...
    echo.
    pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ❌ 依赖安装失败，尝试使用国内镜像源...
        pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
    )
    echo.
)

echo ✓ 依赖检查完成
echo.

:: 启动服务
echo ================================================
echo   正在启动服务...
echo ================================================
echo.

:: 启动 Python 服务
start /b python app.py

:: 等待服务启动
timeout /t 3 /nobreak >nul

:: 打开浏览器
echo 正在打开浏览器...
start http://127.0.0.1:5001

echo.
echo ================================================
echo   服务已启动！
echo ================================================
echo.
echo 拆分功能: http://127.0.0.1:5001
echo 合并功能: http://127.0.0.1:5001/merger
echo.
echo 按任意键停止服务并退出...
pause >nul

:: 停止 Python 进程
taskkill /f /im python.exe /fi "WINDOWTITLE eq Excel*" >nul 2>&1
