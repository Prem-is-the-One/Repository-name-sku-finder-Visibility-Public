@echo off
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python tools\watch_excel.py
pause
