@echo off
echo 正在啟動人資總務管理系統 - 測試版本...
echo.

REM 檢查Python是否安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤: 未找到Python，請先安裝Python 3.7+
    pause
    exit /b 1
)

REM 安裝依賴套件
echo 正在安裝依賴套件...
pip install -r requirements_test.txt

REM 啟動應用程式
echo.
echo 啟動應用程式...
python app_test.py

pause 