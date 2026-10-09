@echo off
chcp 65001 > nul
echo ======================================================================
echo    HỆ THỐNG GẮN LINK TRA CỨU ĐỊNH MỨC THÔNG TƯ 38/2026/TT-BXD
echo ======================================================================
echo.

set UV_PATH=C:\Users\DELL\.local\bin\uv.exe

if not exist "%UV_PATH%" (
    echo [LOI] Khong tim thay uv.exe tai %UV_PATH%
    pause
    exit /b 1
)

if "%~1"=="" (
    echo [INFO] Dang chay tu dong tren cac tep trong thu muc 05.4.2_L1...
    "%UV_PATH%" run --with python-calamine --with openpyxl python "%~dp0Tao_Link_TraCuu_TT38.py"
) else (
    echo [INFO] Dang xu ly tep duoc keo tha vao: %~1
    "%UV_PATH%" run --with python-calamine --with openpyxl python "%~dp0Tao_Link_TraCuu_TT38.py" "%~1"
)

echo.
echo ======================================================================
echo [HOAN TAT] File ket qua da duoc tao voi duoi ten _TT38_Linked.xlsx!
echo ======================================================================
pause
