@echo off
cd /d "%~dp0"

echo Checking for updates...

where python >nul 2>nul
if %errorlevel% equ 0 goto RUN_SCRIPT

echo [ERROR] Python is not installed or not added to PATH.
echo Please install Python and check "Add Python to PATH" during installation.
pause
exit /b 1

:RUN_SCRIPT
python -c "import os, sys, subprocess; d=bytes([208,191,208,176,208,185,209,130,208,190,208,189]).decode('utf-8'); f=bytes([208,190,208,189,208,190,208,178,208,184,209,130,208,184,46,112,121]).decode('utf-8'); subprocess.run([sys.executable, os.path.join(d, f)])"
if %errorlevel% neq 0 goto ERROR_OCCURRED
exit /b 0

:ERROR_OCCURRED
echo.
echo [ERROR] Update failed.
pause
exit /b 1
