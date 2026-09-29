@echo off
rem ============================================================
rem 注册"每天 03:30 自动备份"计划任务（闪电仓选品工具）
rem 任务实际执行 scripts\backup_run.bat，输出写进 backup\备份任务.log。
rem 为什么要日志：备份曾静默停摆 17 天无人发现，因为没有地方留下"它跑过／它失败了"的痕迹。
rem 用法：双击运行；若报权限错误，右键"以管理员身份运行"再试。
rem ============================================================
setlocal
set TASKNAME=闪电仓选品-每日备份
set RUNBAT=%~dp0scripts\backup_run.bat
if not exist "%RUNBAT%" (
  echo [失败] 找不到 %RUNBAT%
  exit /b 1
)
echo 注册：%TASKNAME%  每天 03:30  执行 %RUNBAT%
schtasks /create /f /tn "%TASKNAME%" /tr "%RUNBAT%" /sc daily /st 03:30
if errorlevel 1 (
  echo [失败] 多为权限不足：右键本脚本"以管理员身份运行"重试
  echo        期间可先手动备份：python scripts\backup_daily.py
  exit /b 1
)
echo.
echo ---- 系统里的任务 ----
schtasks /query /tn "%TASKNAME%" /fo LIST
echo.
echo ---- 立即试跑（验证真能产出，不只是注册上）----
schtasks /run /tn "%TASKNAME%"
echo   十几秒后看：backup\最近备份.json 与 backup\备份任务.log
echo   日常：手动备份 python scripts\backup_daily.py ^| 只体检 python scripts\backup_daily.py --check
endlocal
