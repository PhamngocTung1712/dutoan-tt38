@echo off
setlocal
chcp 65001 > nul

echo ===============================================================================
echo   PHAN MEM THAM TRA DU TOAN TU DONG - DOI CHIEU THONG TU 38/2026/TT-BXD VA BAO GIA
echo   Tu dong so sanh dinh muc, doi chieu bao gia, tinh giam tru va xuat bao cao
echo ===============================================================================
echo.

set "UV_PATH=C:\Users\DELL\.local\bin\uv.exe"

if not exist "%UV_PATH%" (
    echo [LOI] Khong tim thay uv.exe tai: %UV_PATH%
    echo Vui long kiem tra lai duong dan Python hoac uv.
    pause
    exit /b 1
)

set "SCRIPT_PATH=%~dp0main_tham_tra.py"
set "OUT_DIR=%~dp0Output_Tham_Tra_Du_Toan"

if "%~1"=="" (
    echo [INFO] Dang quet va tham tra toan bo ho so trong thu muc 05.4.2_L1...
    "%UV_PATH%" run --with python-calamine --with openpyxl --with python-docx python "%SCRIPT_PATH%"
) else (
    echo [INFO] Dang tham tra tep duoc keo tha vao: %~1
    "%UV_PATH%" run --with python-calamine --with openpyxl --with python-docx python "%SCRIPT_PATH%" --file "%~1"
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [LOI] Co loi xay ra trong qua trinh chay tham tra! Ma loi: %ERRORLEVEL%
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [INFO] Dang mo thu muc ket qua tham tra: Output_Tham_Tra_Du_Toan...
if exist "%OUT_DIR%" (
    start "" explorer "%OUT_DIR%"
)

echo.
echo ===============================================================================
echo   [HOAN TAT] Toan bo file Tham tra va Bao cao da san sang trong Output folder!
echo ===============================================================================
echo.
pause
