@echo off
cd /d "%~dp0"

set PYTHON_CMD=python
if exist "%USERPROFILE%\Anaconda3\python.exe" set PYTHON_CMD="%USERPROFILE%\Anaconda3\python.exe"
if exist "%USERPROFILE%\anaconda3\python.exe" set PYTHON_CMD="%USERPROFILE%\anaconda3\python.exe"
if exist "C:\ProgramData\Anaconda3\python.exe" set PYTHON_CMD="C:\ProgramData\Anaconda3\python.exe"

echo Installing dependencies...
%PYTHON_CMD% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Failed. Try: py -m pip install -r requirements.txt
    pause
    exit /b 1
)
echo Done! Run run_app.bat to start.
pause
