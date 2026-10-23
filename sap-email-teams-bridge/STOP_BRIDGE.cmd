@echo off
title Dung SAP Email Teams Bridge
echo Dang tim va dung tien trinh Python chay SAP Bridge...
wmic process where "commandline like '%%projects\\sap-email-teams-bridge\\main.py%%'" call terminate >nul 2>&1
echo [OK] Da dung tien trinh thanh cong.
pause
