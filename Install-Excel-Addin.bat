@echo off
setlocal
title Cai dat M-Sheets cho Excel
set "SRC=%~dp0dist\M-Sheets.xlam"
if not exist "%SRC%" set "SRC=%~dp0M-Sheets.xlam"
set "DEST=%APPDATA%\Microsoft\Excel\XLSTART"
set "TARGET=%DEST%\M-Sheets.xlam"
set "LOCKFILE=%DEST%\~$M-Sheets.xlam"
set "OLD_TARGET=%DEST%\TienIchTQV.xlam"
set "OLD_LOCK=%DEST%\~$TienIchTQV.xlam"

echo ============================================================
echo   CAI DAT M-SHEETS CHO MICROSOFT EXCEL
echo ============================================================
echo.

if not exist "%SRC%" (
    echo [LOI] Khong tim thay file M-Sheets.xlam
    echo       Vui long chay Build-Excel-Addin.ps1 truoc hoac de file cung thu muc.
    echo.
    pause
    exit /b 1
)

if not exist "%DEST%" mkdir "%DEST%"

rem Kiem tra neu Excel dang chay
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    echo Excel dang mo. Dang dong Excel de cap nhat Add-in...
    taskkill /F /IM EXCEL.EXE >nul 2>nul
    ping -n 2 127.0.0.1 >nul
)

if exist "%LOCKFILE%" del /F /Q /A "%LOCKFILE%" 2>nul
if exist "%OLD_LOCK%" del /F /Q /A "%OLD_LOCK%" 2>nul
if exist "%OLD_TARGET%" del /F /Q /A "%OLD_TARGET%" 2>nul

copy /Y "%SRC%" "%TARGET%" >nul
if errorlevel 1 (
    echo [LOI] Khong the sao chep vao thu muc XLSTART.
    echo       Vui long tat han Excel roi chay lai file nay.
    echo.
    pause
    exit /b 1
)

rem Bo chan Macro (Mark of the Web)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -Path '%TARGET%' -ErrorAction SilentlyContinue"

echo.
echo ============================================================
echo   CAI DAT THANH CONG!
echo ============================================================
echo.
echo File add-in da duoc dat tai:
echo   %TARGET%
echo.
echo Tu bay gio moi khi mo Excel, the "M-Sheets" se tu dong
echo xuat hien tren thanh Ribbon voi day du 7 nhom chuc nang!
echo.
pause
