@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo Запуск створення табелів...

where python >nul 2>nul
if %errorlevel% equ 0 goto RUN_SCRIPT

echo [ERROR] Python is not installed or not added to PATH.
echo Please install Python and check "Add Python to PATH" during installation.
pause
exit /b 1

:RUN_SCRIPT
python пайтон\табель.py %*
if %errorlevel% neq 0 goto ERROR_OCCURRED
exit /b 0

:ERROR_OCCURRED
echo.
echo [ERROR] Script execution failed.
pause
exit /b 1
