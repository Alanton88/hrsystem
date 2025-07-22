@echo off
echo 啟動生產環境伺服器...
echo.
echo 請確保已設定 .env 檔案
echo.
gunicorn --bind 0.0.0.0:8000 --workers 2 app:app
pause 