@echo off
title Bihar e-Tender Intelligence Dashboard
echo ========================================================
echo   Bihar eProc2 Tender Intelligence & Management System
echo ========================================================
echo.
echo Starting local dashboard server...
cd /d "%~dp0"
python server.py
pause
