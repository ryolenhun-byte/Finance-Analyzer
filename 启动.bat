@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo.
echo  ========================================
echo   💰 消費分析工具 啟動中
echo  ========================================
echo.
echo  [1/2] 安裝/更新依賴套件...
pip install -r requirements.txt -q --disable-pip-version-check
if errorlevel 1 (
    echo  ⚠️  套件安裝失敗，嘗試繼續啟動...
)

echo  [2/2] 啟動伺服器...
echo  瀏覽器將自動開啟 http://localhost:8765
echo  按 Ctrl+C 停止服務
echo.
python main.py
pause
