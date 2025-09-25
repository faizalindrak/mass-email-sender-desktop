@echo off
setlocal
set PYTHONUNBUFFERED=1
set "TB_QUEUE_DIR=d:\projects\mass-email-sender-desktop\tb_queue"
"C:\Users\faizal.api00650\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\python.exe" -u "%~dp0native_host.py"
exit /b %ERRORLEVEL%
