@echo off
title Table Match Tool - Auto Setup + GUI
echo ============================================
echo   Table Match Tool - Auto Setup + GUI
echo ============================================
echo.

REM ---- 1. Locate/install Python ----
set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY (
    where py >nul 2>nul && set "PY=py -3"
)
if not defined PY (
    echo [Info] Python not found. Installing via winget...
    winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements >nul 2>nul
    if errorlevel 1 (
        echo [Info] winget failed. Trying direct download...
        powershell -NoProfile -ExecutionPolicy Bypass -Command ^
          "$u='https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe';" ^
          "$f='%TEMP%\python-installer.exe';" ^
          "Invoke-WebRequest -Uri $u -OutFile $f -UseBasicParsing;" ^
          "Start-Process -Wait -FilePath $f -ArgumentList '/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=0 Include_test=0';" ^
          "Write-Host 'Installed'"
        if errorlevel 1 (
            echo [Error] Auto-install failed. Install Python from:
            echo   https://www.python.org/downloads/
            pause
            exit /b 1
        )
    )
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
    where python >nul 2>nul && set "PY=python"
    if not defined PY set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    if not exist "%PY%" (
        echo [Error] Cannot locate python. Restart this script.
        pause
        exit /b 1
    )
)

echo [OK] Python: 
%PY% --version

REM ---- 2. Auto-install deps ----
echo [Info] Checking dependencies...
%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo [Info] Installing openpyxl...
    %PY% -m pip install openpyxl
)
%PY% -c "import tkinter" >nul 2>nul
if errorlevel 1 (
    echo [Error] tkinter missing in Python. Reinstall Python with "tcl/tk and IDLE" option.
    echo [Hint] Use the installer and add "tcl/tk and IDLE" feature.
    pause
    exit /b 1
)

REM ---- 3. Run GUI ----
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
