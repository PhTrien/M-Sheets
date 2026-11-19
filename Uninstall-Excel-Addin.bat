@echo off
setlocal
title Go cai dat M-Sheets cho Excel
set "DEST=%APPDATA%\Microsoft\Excel\XLSTART"
set "TARGET=%DEST%\M-Sheets.xlam"
set "LOCKFILE=%DEST%\~$M-Sheets.xlam"
set "OLD_TARGET=%DEST%\TienIchTQV.xlam"
set "OLD_LOCK=%DEST%\~$TienIchTQV.xlam"

echo ============================================================
echo   GO CAI DAT M-SHEETS CHO MICROSOFT EXCEL
echo ============================================================
echo.

tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    echo Excel dang mo. Dang dong Excel de go bo Add-in...
    taskkill /F /IM EXCEL.EXE >nul 2>nul
    ping -n 2 127.0.0.1 >nul
)

if exist "%LOCKFILE%" del /F /Q /A "%LOCKFILE%" 2>nul
if exist "%OLD_LOCK%" del /F /Q /A "%OLD_LOCK%" 2>nul
if exist "%OLD_TARGET%" del /F /Q /A "%OLD_TARGET%" 2>nul

if exist "%TARGET%" (
    del /F /Q /A "%TARGET%" 2>nul
    echo Da xoa file Add-in M-Sheets tai thu muc XLSTART.
) else (
    echo Khong tim thay Add-in M-Sheets trong XLSTART.
)

echo.
echo ============================================================
echo   GO CAI DAT THANH CONG!
echo ============================================================
echo.
pause
