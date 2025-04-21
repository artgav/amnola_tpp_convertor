@echo off
cd /d %~dp0

echo Running PDF to Google Drive converter...
python batch_convert_upload.py

echo.
echo Script finished. Press any key to close this window.
pause
