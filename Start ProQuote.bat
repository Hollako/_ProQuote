@echo off
setlocal
title ProQuote - SmartWay Systems (leave this window open)
cd /d "%~dp0"

REM ---- Find a working Python: py launcher, then python on PATH, then known install ----
set "PYCMD="
where py >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD (
  where python >nul 2>nul && set "PYCMD=python"
)
if not defined PYCMD (
  if exist "%LocalAppData%\Programs\Python\Python313\python.exe" set PYCMD="%LocalAppData%\Programs\Python\Python313\python.exe"
)
if not defined PYCMD (
  echo.
  echo   ERROR: Python was not found on this PC.
  echo   Install Python 3 from https://www.python.org/downloads/
  echo   and tick "Add python.exe to PATH" during setup, then run this again.
  echo.
  pause
  exit /b 1
)

echo ============================================================
echo   ProQuote is starting...
echo   Open in browser:  http://localhost:8501
echo   From another PC:  http://192.168.100.126:8501
echo   (keep this window open; close it to stop the app)
echo ============================================================

%PYCMD% -m streamlit run app.py --server.headless=true --server.address=0.0.0.0 --server.port=8501 --browser.gatherUsageStats=false
pause
