@echo off
setlocal
set PYTHONUNBUFFERED=1
set "TB_QUEUE_DIR=D:\projects\mass-email-sender-desktop\tb_queue"
"D:\projects\mass-email-sender-desktop\venv\Scripts\python.exe" -u "%~dp0native_host.py"
exit /b %ERRORLEVEL%
