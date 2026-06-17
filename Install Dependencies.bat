@echo off
setlocal
title ProQuote dependency installer
cd /d "%~dp0"

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
  echo   and tick "Add python.exe to PATH", then run this again.
  echo.
  pause
  exit /b 1
)

echo Installing ProQuote Python dependencies...
%PYCMD% -m pip install --upgrade pip
%PYCMD% -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Dependency installation failed. Check the internet connection and try again.
  pause
  exit /b 1
)

echo.
echo Done.
pause