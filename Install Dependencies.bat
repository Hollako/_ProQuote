@echo off
setlocal
title ProQuote dependency installer
cd /d "%~dp0"

set "SILENT="
if /I "%~1"=="/silent" set "SILENT=1"

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
  if not defined SILENT pause
  exit /b 1
)

echo Installing ProQuote Python dependencies...
%PYCMD% -m pip install --upgrade pip
if errorlevel 1 goto :failed
%PYCMD% -m pip install -r requirements.txt
if errorlevel 1 goto :failed

echo.
echo Done.
if not defined SILENT pause
exit /b 0

:failed
echo.
echo Dependency installation failed. Check the internet connection and try again.
if not defined SILENT pause
exit /b 1
