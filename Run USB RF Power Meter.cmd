@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHONW=D:\.pyvenv\Scripts\pythonw.exe"
set "PYTHON=D:\.pyvenv\Scripts\python.exe"

if exist "%PYTHONW%" (
    start "" /d "%ROOT%" "%PYTHONW%" "%ROOT%main.py"
    exit /b 0
)

if exist "%PYTHON%" (
    start "" /d "%ROOT%" "%PYTHON%" "%ROOT%main.py"
    exit /b 0
)

echo Python was not found at D:\.pyvenv\Scripts.
echo Update "Run USB RF Power Meter.cmd" if your virtual environment moved.
pause
exit /b 1
