@echo off
title YiBo Express Bill System

echo.
echo  ========================================
echo    YiBo Express Bill System Starting...
echo  ========================================
echo.

cd /d "%~dp0"

pip install flask python-dotenv -q

echo.
echo  Local:   http://127.0.0.1:5000
echo  Check LAN IP in the console below.
echo.
echo  Close this window to stop the service.
echo  ========================================
echo.

python app.py
pause