@echo off
title 竞争分析面板 - 关闭此窗口即停止服务
cd /d "%~dp0"
echo ================================================
echo   竞争分析面板启动中（智能启动器）...
echo   自动识别/替换旧版进程，无需手动清理
echo   浏览器将自动打开面板页面
echo   关闭此窗口即停止服务
echo ================================================
python scripts\launch_panel.py
if errorlevel 1 (
    echo.
    echo [启动失败] 请检查 Python 是否安装，或运行 python --version 验证
    pause
)
