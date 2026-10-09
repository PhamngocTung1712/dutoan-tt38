@echo off
setlocal
title He Thong Tham Tra Du Toan - Public Online
chcp 65001 > nul

set "UV_PATH=C:\Users\DELL\.local\bin\uv.exe"
set "RUNNER_SCRIPT=%~dp0chay_online.py"

"%UV_PATH%" run --with flask --with python-calamine --with openpyxl --with python-docx python "%RUNNER_SCRIPT%"

pause
