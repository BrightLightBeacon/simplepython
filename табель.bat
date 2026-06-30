@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo Запуск створення табелів...

python пайтон\табель.py %*
if %errorlevel% equ 0 exit /b 0

echo.
echo [ERROR] Script execution failed.
pause
exit /b 1
