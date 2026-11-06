@echo off
chcp 65001 >nul
echo ==========================================================
echo ĐANG ĐÓNG GÓI WH_ENGINE.EXE (PORTABLE ZERO-INSTALL)...
echo ==========================================================

cd /d "%~dp0"
pyinstaller --noconfirm --onefile --noconsole --clean --name "wh_engine" wh_engine.py

if exist "dist\wh_engine.exe" (
    copy /Y "dist\wh_engine.exe" "wh_engine.exe" >nul
    rmdir /S /Q build dist
    del /Q wh_engine.spec
    echo [THÀNH CÔNG] Đã tạo file wh_engine.exe độc lập tại thư mục hiện tại!
) else (
    echo [LỖI] Quá trình đóng gói thất bại.
)
pause
