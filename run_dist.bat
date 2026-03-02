@echo off
cd /d "%~dp0"
if not exist "dist\BondBuyerMatch.exe" (
    echo dist\BondBuyerMatch.exe not found. Run build_exe.bat first.
    pause
    exit /b 1
)
echo Running BondBuyerMatch...
dist\BondBuyerMatch.exe
echo.
echo Exit code: %errorlevel%
pause
