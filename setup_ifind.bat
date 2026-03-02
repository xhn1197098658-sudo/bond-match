@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 安装 iFinD 接口
echo ========================================
echo 正在配置 iFinD (同花顺) 数据接口...
echo ========================================
echo.

REM 尝试查找 Anaconda
set ANACONDA_PYTHON=
if exist "%USERPROFILE%\Anaconda3\python.exe" (
    set ANACONDA_PYTHON=%USERPROFILE%\Anaconda3\python.exe
    echo 检测到 Anaconda: %ANACONDA_PYTHON%
) else if exist "%USERPROFILE%\anaconda3\python.exe" (
    set ANACONDA_PYTHON=%USERPROFILE%\anaconda3\python.exe
    echo 检测到 Anaconda: %ANACONDA_PYTHON%
) else if exist "C:\ProgramData\Anaconda3\python.exe" (
    set ANACONDA_PYTHON=C:\ProgramData\Anaconda3\python.exe
    echo 检测到 Anaconda: %ANACONDA_PYTHON%
)

REM 优先使用 Anaconda，然后尝试 py，最后尝试 python/pip
if defined ANACONDA_PYTHON (
    echo.
    echo 使用 Anaconda Python 安装 iFinDAPI...
    "%ANACONDA_PYTHON%" -m pip install iFinDAPI
    if errorlevel 1 (
        echo.
        echo ========================================
        echo 安装失败！
        echo ========================================
        echo.
        echo 请检查网络连接，或尝试在 Cursor 终端运行:
        echo   conda activate base
        echo   pip install iFinDAPI
        echo.
        pause
        exit /b 1
    )
) else (
    echo 未检测到 Anaconda，尝试使用系统 Python...
    py -m pip install iFinDAPI
    if errorlevel 1 (
        python -m pip install iFinDAPI
    )
    if errorlevel 1 (
        pip install iFinDAPI
    )
    if errorlevel 1 (
        echo.
        echo ========================================
        echo 安装失败！
        echo ========================================
        echo.
        echo 请使用 Anaconda Prompt 或 Cursor 终端运行:
        echo   conda activate base
        echo   pip install iFinDAPI
        echo.
        pause
        exit /b 1
    )
)

echo.
echo ========================================
echo iFinD 接口包已安装！
echo ========================================
echo.
echo 请在程序中通过「帮助 - iFinD 设置」填写账号密码
echo 数据接口文档: https://quantapi.51ifind.com/
echo.
echo 按任意键关闭...
pause >nul
