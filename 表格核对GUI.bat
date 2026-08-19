@echo off
title Table Match Tool - Auto Setup + GUI
echo ============================================
echo   Table Match Tool - Auto Setup + GUI
echo ============================================
echo.

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY (
    where py >nul 2>nul && set "PY=py -3"
)

REM --- validate tkinter ---
set "HAS_TK="
if defined PY (
    %PY% -c "import tkinter" >nul 2>nul && set "HAS_TK=1"
)

if not defined HAS_TK (
    echo [Info] No Python with tkinter found. Installing full Python...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "$ErrorActionPreference='Stop';" ^
      "$u='https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe';" ^
      "$f=\"$env:TEMP\python-full.exe\";" ^
      "Invoke-WebRequest -Uri $u -OutFile $f -UseBasicParsing;" ^
      "Start-Process -Wait -FilePath $f -ArgumentList '/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=0 Include_tcltk=1 Include_test=0';" ^
      "Write-Host 'Installed'"
    if errorlevel 1 (
        echo [Error] Auto-install failed. Install Python from:
        echo   https://www.python.org/downloads/
        pause
        exit /b 1
    )
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
    if not exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        set "PY=" & where python >nul 2>nul && set "PY=python"
    ) else (
        set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    )
    if not defined PY set "PY=python"
)

echo [OK] Python:
%PY% --version 2>&1
%PY% -c "import tkinter; print('   tkinter OK')" 2>&1

REM --- ensure openpyxl (always try install if missing) ---
%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo [Info] Installing openpyxl...
    %PY% -m pip install openpyxl
    IF ERRORLEVEL 1 (
        echo [Error] openpyxl install failed.
        pause
        exit /b 1
    )
)
%PY% -c "import openpyxl; print('   openpyxl OK')" 2>&1

REM --- run GUI ---
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
