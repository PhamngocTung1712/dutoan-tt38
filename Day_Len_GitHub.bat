@echo off
setlocal
chcp 65001 > nul
title Day Code Len GitHub - Du Toan Chuyen Nghiep

echo ===============================================================================
echo        DAY MA NGUON LEN GITHUB DE DUA LEN DAM MAY RENDER (24/7)
echo ===============================================================================
echo.
echo Tai khoan GitHub da duoc ghi nhan: PhamngocTung1712
echo.
echo [BUOC 1] Neu anh da tao san Repository tren GitHub (vi du: dutoan-tt38),
echo          hay dan link HTTPS cua Repository do vao day:
echo          (Vi du: https://github.com/PhamngocTung1712/dutoan-tt38.git)
echo.
set /p REPO_URL="Nhap link GitHub Repo: "

if "%REPO_URL%"=="" (
    echo [LOI] Chua nhap link GitHub Repository.
    pause
    exit /b 1
)

echo.
echo Dang cau hinh va day toan bo du lieu len GitHub...
git branch -M main
git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%
git push -u origin main --force

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ===============================================================================
    echo   [THANH CONG] Toan bo ma nguon da duoc day len GitHub!
    echo   Bay gio anh co the vao Render.com de lien ket va chay 24/7.
    echo ===============================================================================
) else (
    echo.
    echo [CHUY Y] Neu GitHub yeu cau dang nhap hoac nhap Token, anh vui long dang nhap tren trinh duyet.
)

echo.
pause
