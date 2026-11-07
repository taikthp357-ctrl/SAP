@echo off
setlocal
cd /d "%~dp0"

set PYTHON_EXE=C:\Users\taikt\.gemini\antigravity\brain\09dccbf3-c1cd-456a-bbe5-1e523440b2c4\scratch\build_env\Scripts\python.exe
set PYINSTALLER_EXE=C:\Users\taikt\.gemini\antigravity\brain\09dccbf3-c1cd-456a-bbe5-1e523440b2c4\scratch\build_env\Scripts\pyinstaller.exe

echo ===============================================================
echo          DONG GOI SAP_ZONE.EXE STANDALONE (v1.6)
echo ===============================================================

"%PYINSTALLER_EXE%" --noconfirm --clean ^
    --onefile ^
    --windowed ^
    --name "SAP_ZONE" ^
    --add-data "zone_map.json;." ^
    main.py

if %ERRORLEVEL% EQU 0 (
    echo [OK] Dong goi thanh cong tai dist\SAP_ZONE.exe!
) else (
    echo [ERR] Dong goi that bai voi ma loi %ERRORLEVEL%
)

endlocal
