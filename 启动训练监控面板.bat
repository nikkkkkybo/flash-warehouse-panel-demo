@echo off
title 小模型微调监控面板
cd /d "%~dp0"
echo ================================================
echo   小模型微调监控面板 http://127.0.0.1:5010
echo   关闭此窗口即停止监控服务
echo ================================================
start "" http://127.0.0.1:5010
python scripts\train_panel.py
if errorlevel 1 (
    echo.
    echo [启动失败] 请检查 Python 是否安装
    pause
)
