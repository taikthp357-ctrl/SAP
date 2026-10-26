@echo off
setlocal enabledelayedexpansion
title Dung SAP Email Teams Bridge

echo [STOP] Dang gui tin hieu dung an toan (Cooperative Stop)...
if exist "python_runtime\python.exe" (
    "python_runtime\python.exe" -c "from app.runtime_control import WorkerLease; WorkerLease.request_stop('data')" >nul 2>&1
) else (
    python -c "from app.runtime_control import WorkerLease; WorkerLease.request_stop('data')" >nul 2>&1
)

echo [STOP] Cho tien trinh hoan tat chu trinh (toi da 5 giay)...
timeout /t 3 /nobreak >nul

set PID=
if exist "data\bridge.pid" (
    set /p PID=<"data\bridge.pid"
)
if "!PID!"=="" if exist "data\worker.json" (
    for /f "tokens=2 delims=:, " %%a in ('findstr "pid" data\worker.json') do set PID=%%a
)

if not "!PID!"=="" (
    tasklist /fi "PID eq !PID!" 2>nul | findstr /i "!PID!" >nul
    if not errorlevel 1 (
        echo [STOP] Tien trinh PID !PID! van con chay, thuc hien dung an toan...
        taskkill /PID !PID! /T /F >nul 2>&1
    )
)

echo [OK] Da gui tin hieu dung tien trinh an toan.
pause
