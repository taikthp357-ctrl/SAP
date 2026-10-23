@echo off
title SAP - Outlook - Teams Bridge Worker
cd /d "%~dp0"
echo ==================================================================
echo  KHOI DONG SAP - OUTLOOK - TEAMS BRIDGE WORKER
echo  (0 OPEN PORT / 0 POWERSHELL / AUTO KEEP-ALIVE ENABLED)
echo ==================================================================
python main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Tien trinh bi ngat hoac xay ra loi. Nhan phim bat ky de thoat...
    pause >nul
)
