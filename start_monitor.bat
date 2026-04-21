@echo off
cd /d "%~dp0"
REM 使用 pythonw 以无控制台窗口方式运行
start "" ".venv\Scripts\python.exe" main.py

