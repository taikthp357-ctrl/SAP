@echo off
setlocal
title SAP - Outlook - Teams Bridge Worker
cd /d "%~dp0"
echo ==================================================================
echo  KHOI DONG SAP - OUTLOOK - TEAMS BRIDGE WORKER
echo  (0 OPEN PORT / 0 POWERSHELL / AUTO KEEP-ALIVE ENABLED)
echo ==================================================================

set "PYTHON_EXE="

:: 0. Uu tien 1: Python Portable Runtime tich hop san trong thu muc tool
if exist "%~dp0runtime\python.exe" (
    set "PYTHON_EXE="%~dp0runtime\python.exe""
    goto :RUN_APP
)

:: 1. Uu tien 2: py.exe (Python Launcher chuan Windows)
py -3 -c "import sys" >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_EXE=py -3"
    goto :RUN_APP
)

:: 2. Uu tien 3: Kiem tra duong dan cai dat trong AppData cua User
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

:: 3. Uu tien 4: Duong dan Program Files
if exist "C:\Program Files\Python312\python.exe" (
    set "PYTHON_EXE="C:\Program Files\Python312\python.exe""
    goto :RUN_APP
)
if exist "C:\Program Files\Python311\python.exe" (
    set "PYTHON_EXE="C:\Program Files\Python311\python.exe""
    goto :RUN_APP
)

:: 4. Uu tien 5: Kiem tra python trong PATH
python -c "import sys" >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_EXE=python"
    goto :RUN_APP
)

echo [ERROR] May cua ban hien chua co Python hoac Runtime tich hop san!
echo.
echo Giai phap khuyen nghi:
echo  - Ban nen su dung ban 'SAP_EMAIL_TEAMS_BRIDGE_PORTABLE_v1.0.0.zip'
echo    (Ban nay da tich hop san Python 3.12 Portable, chi can giai nen la chay ngay 100%%).
echo.
echo Hoac neu muon dung ban nhe:
echo  1. Vao Windows Settings ^> Apps ^> Advanced app settings ^> App execution aliases
echo     ^> Tat (OFF) ca 2 muc 'App Installer (python.exe)' va 'App Installer (python3.exe)'.
echo  2. Cai dat Python tu https://www.python.org (nho tich vao o 'Add Python to PATH').
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
