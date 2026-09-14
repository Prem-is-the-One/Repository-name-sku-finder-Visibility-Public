@echo off
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python -m http.server 5500
pause
