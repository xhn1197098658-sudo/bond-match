@echo off
cd /d "%~dp0"

set OUT=BondBuyerMatch_Portable
echo Creating deployment folder...

if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%"
mkdir "%OUT%\database"
mkdir "%OUT%\data"
mkdir "%OUT%\hooks"

copy app.py "%OUT%\"
copy data_provider.py "%OUT%\"
copy ifind_api.py "%OUT%\"
copy requirements.txt "%OUT%\"
copy run_app.bat "%OUT%\"
copy install_requirements.bat "%OUT%\"

copy database\schema.sql "%OUT%\database\"
copy database\db_manager.py "%OUT%\database\"
if exist database\__init__.py copy database\__init__.py "%OUT%\database\"
if exist hooks\hook-matplotlib.py copy hooks\hook-matplotlib.py "%OUT%\hooks\"

if exist "database\bond_buyer_match.db" (
    copy "database\bond_buyer_match.db" "%OUT%\database\"
    echo [OK] Database copied
) else (
    echo [OK] Empty database on first run
)

if exist "dist\BondBuyerMatch.exe" (
    copy "dist\BondBuyerMatch.exe" "%OUT%\"
    echo [OK] EXE copied
)

(
echo ============================================
echo BondBuyerMatch - Deploy to another PC
echo ============================================
echo.
echo [Option 1] Use EXE - double-click BondBuyerMatch.exe
echo.
echo [Option 2] Use Python:
echo   STEP 1: Install Python 3.8+ or Anaconda
echo   STEP 2: Run install_requirements.bat (MUST do this first!)
echo   STEP 3: Run run_app.bat
echo.
echo If "No module named pandas" - run install_requirements.bat first!
echo.
echo Import: Help - Import Holdings / CanBuy / Contacts
echo iFinD: Help - iFinD Settings
echo ============================================
) > "%OUT%\README.txt"

echo.
echo Done: %OUT%\
explorer "%OUT%"
pause
