@echo off
setlocal
chcp 65001 > nul

echo ===============================================================================
echo   MO HE THONG APP WEB THAM TRA DU TOAN THONG TU 38/2026/TT-BXD
echo   Ung dung web truc quan, xem truc tiep PDF va tai bao cao Word, Excel
echo ===============================================================================
echo.

set "UV_PATH=C:\Users\DELL\.local\bin\uv.exe"

if not exist "%UV_PATH%" (
    echo [LOI] Khong tim thay uv.exe tai: %UV_PATH%
    pause
    exit /b 1
)

set "WEB_SCRIPT=%~dp0app_web.py"

echo [INFO] Dang khoi dong Web Server tai http://localhost:5000 ...
echo [INFO] Trinh duyet Chrome/Edge se tu dong mo trong vai giay...
echo.

"%UV_PATH%" run --with flask --with python-calamine --with openpyxl --with python-docx python "%WEB_SCRIPT%"

pause
