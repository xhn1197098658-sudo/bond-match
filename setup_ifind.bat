@echo off
chcp 65001 >nul
echo 正在配置 iFinD (同花顺) 数据接口...
echo.
pip install iFinDAPI
if errorlevel 1 (
    echo 安装失败，请检查网络或尝试: pip install --upgrade iFinDAPI
    pause
    exit /b 1
)
echo.
echo iFinD 接口包已安装。请在程序中通过「帮助 - 数据源设置」选择 iFinD 并填写账号密码。
echo 数据接口文档: https://quantapi.51ifind.com/
pause
