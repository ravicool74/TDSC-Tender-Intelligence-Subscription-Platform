@echo off
title Bihar e-Tender Intelligence Dashboard
echo ========================================================
echo   TDSC Bihar eProc2 ^& Central CPPP Intelligence Platform
echo ========================================================
echo.
echo Connecting to local server: http://127.0.0.1:8000 ...
cd /d "%~dp0"
python server.py
pause
