@echo off
rem ============================================================
rem  内网部署：注册为 Windows 服务（需 NSSM，方案A）
rem  步骤：
rem   1. 下载 nssm.exe: https://nssm.cc/download → 放到本目录或 PATH
rem   2. 首次使用：先在下面设置面板账号密码（PANEL_USER/PANEL_PASS）
rem      密码建议：12位以上字母数字混合
rem   3. 右键"以管理员身份运行"本脚本
rem   4. 同事浏览器访问 http://<本机内网IP>:5000 （输入账号密码）
rem  卸载服务：nssm remove 闪电仓选品面板 confirm
rem ============================================================
set SERVICE=闪电仓选品面板
set PY=%~dp0..\dist\闪电仓选品工具\python\python.exe
if not exist "%PY%" set PY=python
set APP=%~dp0web_panel.py

rem >>> 部署时修改下面两行（引号保留）<<<
set PANEL_USER=chaotu
set PANEL_PASS=请改成强密码
set PANEL_NO_BROWSER=1

nssm install "%SERVICE%" "%PY%" "%APP%"
nssm set "%SERVICE%" AppDirectory "%~dp0"
nssm set "%SERVICE%" AppEnvironmentExtra PANEL_USER=%PANEL_USER% PANEL_PASS=%PANEL_PASS% PANEL_NO_BROWSER=%PANEL_NO_BROWSER%
nssm set "%SERVICE%" Start SERVICE_AUTO_START
nssm set "%SERVICE%" AppStdout "%~dp0..\output\logs\service.log"
nssm set "%SERVICE%" AppStderr "%~dp0..\output\logs\service-err.log"
nssm set "%SERVICE%" AppRotateFiles 1
nssm start "%SERVICE%"

echo.
echo 服务已注册并启动。验证：http://127.0.0.1:5000/healthz
echo 内网访问：http://本机IP:5000 （账号 %PANEL_USER%）
echo 查看服务：nssm status "%SERVICE%"
pause
