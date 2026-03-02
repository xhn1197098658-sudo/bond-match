@echo off
cd /d "%~dp0"

echo ========================================
echo Bond Match - Build EXE
echo ========================================
echo.

REM Use Python (Anaconda if exists)
set PYTHON_CMD=python
if exist "%USERPROFILE%\Anaconda3\python.exe" set PYTHON_CMD="%USERPROFILE%\Anaconda3\python.exe"
if exist "%USERPROFILE%\anaconda3\python.exe" set PYTHON_CMD="%USERPROFILE%\anaconda3\python.exe"
if exist "C:\ProgramData\Anaconda3\python.exe" set PYTHON_CMD="C:\ProgramData\Anaconda3\python.exe"

echo Checking compatibility...
%PYTHON_CMD% -m pip uninstall enum34 -y 2>nul

echo Checking PyInstaller...
%PYTHON_CMD% -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    %PYTHON_CMD% -m pip install pyinstaller
)
if errorlevel 1 (
    echo ERROR: PyInstaller not found. Run install_requirements.bat first.
    pause
    exit /b 1
)

echo.
echo Building... (about 3-5 min, do not close)
echo.

if exist "dist" rmdir /s /q "dist" 2>nul
if exist "build" rmdir /s /q "build" 2>nul

%PYTHON_CMD% -m PyInstaller --clean bond_match.spec

if errorlevel 1 (
    echo.
    echo BUILD FAILED!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build OK!
echo Output: dist\BondBuyerMatch.exe (single file)
echo DB bond_buyer_match.db will be in that folder
echo ========================================
pause
