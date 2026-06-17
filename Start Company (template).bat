@echo off
setlocal
REM ============================================================
REM  Per-company launcher (Model A: one shared code, one data
REM  profile per company). Copy this file for each company and
REM  edit the two lines below.
REM ============================================================
REM  COMPANY_DIR : folder holding that company's proquote.db
REM                and assets\header_banner.png (created on first run).
REM  PORT        : give each company a different port.
set "COMPANY_DIR=%~dp0data\company-b"
set "PORT=8502"
REM ============================================================

title ProQuote - %COMPANY_DIR%  (leave this window open)
cd /d "%~dp0"
set "BOQ_DATA_DIR=%COMPANY_DIR%"

REM ---- Find a working Python: py launcher, then python, then known install ----
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
  echo   ERROR: Python was not found. Install Python 3 from https://www.python.org/downloads/
  echo   and tick "Add python.exe to PATH", then run this again.
  echo.
  pause
  exit /b 1
)

echo ============================================================
echo   ProQuote is starting...
echo   Company data:  %BOQ_DATA_DIR%
echo   Open in browser:  http://localhost:%PORT%
echo   (keep this window open; close it to stop the app)
echo ============================================================

%PYCMD% -m streamlit run app.py --server.headless=true --server.address=0.0.0.0 --server.port=%PORT% --browser.gatherUsageStats=false
pause
