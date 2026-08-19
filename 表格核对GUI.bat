@echo off
title Table Match Tool - Auto Setup + GUI
echo ============================================
echo   Table Match Tool - Auto Setup + GUI
echo ============================================
echo.

REM ---- 1. Locate Python (tkinter included) ----
set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY (
    where py >nul 2>nul && set "PY=py -3"
)
REM Check tkinter exists
if defined PY (
    %PY% -c "import tkinter" >nul 2>nul
    if errorlevel 1 set "PY="
    %PY% -c "import openpyxl" >nul 2>nul || %PY% -m pip install openpyxl
)
if not defined PY (
    echo [Info] No Python with tkinter found. Installing full Python...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "$ErrorActionPreference='Stop';" ^
      "$u='https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe';" ^
      "$f=\"$env:TEMP\python-full.exe\";" ^
      "Write-Host 'Downloading...';" ^
      "Invoke-WebRequest -Uri $u -OutFile $f -UseBasicParsing;" ^
      "Write-Host 'Installing (includes tcl/tk)...';" ^
      "Start-Process -Wait -FilePath $f -ArgumentList '/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=0 Include_tcltk=1 Include_test=0';" ^
      "Write-Host 'Installed'"
    if errorlevel 1 (
        echo [Error] Auto-install failed. Please install Python from:
        echo   https://www.python.org/downloads/
        echo   IMPORTANT: keep default "tcl/tk and IDLE" checked.
        pause
        exit /b 1
    )
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
    set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    if not exist "%PY%" set "PY=python"
)

echo [OK] Python:
%PY% --version 2>&1
%PY% -c "import tkinter; print('   tkinter OK')" 2>&1
%PY% -c "import openpyxl; print('   openpyxl OK')" 2>&1

REM ---- 2. Run GUI ----
set "SCRIPT=%~dp0table_match_gui.py"
if not exist "%SCRIPT%" set "SCRIPT=%BASE%..\table_match_gui.py"
if not exist "%SCRIPT%" (
    echo [Error] table_match_gui.py not found.
    pause
    exit /b 1
)
echo [Start] Launching GUI...
"%PY%" "%SCRIPT%"
echo.
pause
