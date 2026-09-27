@echo off
chcp 65001 >nul
title BankReport Mapping Hub
echo Khoi dong...
start http://127.0.0.1:5050
python "%~dp0app.py"
pause
