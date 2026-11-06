@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "WH_DanhSo_TuDong.py"
) else (
    python "WH_DanhSo_TuDong.py"
)
if errorlevel 1 (
  echo.
  echo TOOL KHONG CHAY DUOC: Kiem tra Python 3 va PyMuPDF.
  echo Neu duoc phep cai thu vien: python -m pip install --user pymupdf
  pause
)
