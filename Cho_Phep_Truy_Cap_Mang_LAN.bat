@echo off
setlocal
chcp 65001 > nul

echo ===============================================================================
echo   MO CONG 5000 TREN WINDOWS FIREWALL CHO PHEP MAY KHAC TRUY CAP MANG LAN
echo ===============================================================================
echo.
echo Dang them quy tac vao Windows Defender Firewall...

netsh advfirewall firewall add rule name="Du Toan Chuyen Nghiep App (Port 5000)" dir=in action=allow protocol=TCP localport=5000

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [THANH CONG] Da mo cong 5000! Gio cac may khac trong mang Wi-Fi co the vao duoc.
) else (
    echo.
    echo [CHU Y] Vui long nhap chuot phai vao file nay va chon "Run as administrator" (Chay voi quyen Admin).
)

echo.
pause
