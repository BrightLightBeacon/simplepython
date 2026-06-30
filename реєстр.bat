@echo off
cd /d "%~dp0"

echo Running registry generator...

python -c "import os, sys, subprocess; d=bytes([208,191,208,176,208,185,209,130,208,190,208,189]).decode('utf-8'); f=bytes([209,128,208,181,209,148,209,129,209,130,209,128,46,112,121]).decode('utf-8'); subprocess.run([sys.executable, os.path.join(d, f)])"
if %errorlevel% equ 0 exit /b 0

echo.
echo [ERROR] Script execution failed.
pause
exit /b 1
