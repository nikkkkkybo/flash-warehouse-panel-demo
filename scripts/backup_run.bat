@echo off
rem Wrapper invoked by scheduled task "ShanDianCang-XuanPin - Daily Backup".
rem Rules for this file: pure ASCII, no redirection, no Chinese.
rem Logging is done inside backup_daily.py (it writes backup\backup_task.log and the
rem JSON status file) - an earlier version redirected from here and the batch body
rem got flattened onto a single line, so the task ran but returned 1 and wrote a
rem malformed log name. Keep this wrapper dumb on purpose.
setlocal
cd /d "%~dp0.."
set PY=python
for /d %%D in ("dist\*") do if exist "%%D\python\python.exe" set PY=%%D\python\python.exe
"%PY%" "scripts\backup_daily.py"
exit /b %errorlevel%
