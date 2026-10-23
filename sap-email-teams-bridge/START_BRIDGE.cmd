@echo off
setlocal
title SAP - Outlook - Teams Bridge Worker
cd /d "%~dp0"
echo ==================================================================
echo  KHOI DONG SAP - OUTLOOK - TEAMS BRIDGE WORKER
echo  (0 OPEN PORT / 0 POWERSHELL / AUTO KEEP-ALIVE ENABLED)
echo ==================================================================

set "PYTHON_EXE="

:: 1. Uu tien py.exe (Python Launcher chuan Windows)
py -3 -c "import sys" >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_EXE=py -3"
    goto :RUN_APP
)

:: 2. Uu tien duong dan cai dat trong AppData cua User
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python312\python.exe""
    goto :RUN_APP
)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python311\python.exe""
    goto :RUN_APP
)
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" (
    set "PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python310\python.exe""
    goto :RUN_APP
)

:: 3. Uu tien duong dan Program Files
if exist "C:\Program Files\Python312\python.exe" (
    set "PYTHON_EXE="C:\Program Files\Python312\python.exe""
    goto :RUN_APP
)
if exist "C:\Program Files\Python311\python.exe" (
    set "PYTHON_EXE="C:\Program Files\Python311\python.exe""
    goto :RUN_APP
)

:: 4. Kiem tra python trong PATH
python -c "import sys" >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_EXE=python"
    goto :RUN_APP
)

echo [ERROR] Khong tim thay Python hop le tren he thong!
echo.
echo Cach xu ly nhanh:
echo 1. Vao Windows Settings ^> Apps ^> Advanced app settings ^> App execution aliases
echo    ^> Tat (OFF) ca 2 muc 'App Installer (python.exe)' va 'App Installer (python3.exe)'.
echo 2. Hoac cai dat Python tu https://www.python.org (nho tich 'Add Python to PATH').
echo.
pause
exit /b 1

:RUN_APP
echo [OK] Phat hien Python: %PYTHON_EXE%
echo [INFO] Dang khoi chay tien trinh main.py...
echo.

%PYTHON_EXE% main.py %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Tien trinh bi ngat hoac xay ra loi. Nhan phim bat ky de thoat...
    pause >nul
)
