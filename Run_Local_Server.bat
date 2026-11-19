@echo off
title MASTER EXCEL TOOLS - Ultra-Fast Local Engine
cd /d "%~dp0"
echo ========================================================
echo   MASTER EXCEL TOOLS (Localhost Mode)
echo   Zero Internet Upload - 100%% Local Processing
echo ========================================================
echo.
echo Starting server on http://localhost:8080 ...
start http://localhost:8080
echo.
python server.py
pause
