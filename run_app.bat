@echo off
cd /d "%~dp0"

set PYTHON_CMD=python
if exist "%USERPROFILE%\Anaconda3\python.exe" set PYTHON_CMD="%USERPROFILE%\Anaconda3\python.exe"
if exist "%USERPROFILE%\anaconda3\python.exe" set PYTHON_CMD="%USERPROFILE%\anaconda3\python.exe"
if exist "C:\ProgramData\Anaconda3\python.exe" set PYTHON_CMD="C:\ProgramData\Anaconda3\python.exe"

%PYTHON_CMD% app.py
if errorlevel 1 pause
