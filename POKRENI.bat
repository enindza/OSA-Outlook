@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Pokrece se Outlook EML Saver...
echo.
python save_emails.py
pause