@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
cd /d "%~dp0"

echo.
echo  ========================================
echo   消費分析工具 啟動中
echo  ========================================
echo.

:: 檢查 Python 是否存在
python --version >nul 2>&1
if errorlevel 1 (
    echo  [錯誤] 找不到 Python，請先安裝 Python 3.10+
    echo  下載: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 安裝/更新依賴套件
echo  [1/2] 確認依賴套件...
pip install fastapi "uvicorn[standard]" python-multipart pandas openpyxl chardet anthropic pdfplumber -q --disable-pip-version-check
if errorlevel 1 (
    echo  [警告] 部分套件安裝失敗，嘗試繼續...
)

:: 啟動伺服器
echo  [2/2] 啟動伺服器...
echo.
echo  瀏覽器將自動開啟 http://localhost:8765
echo  關閉此視窗即可停止服務
echo.
python main.py

pause
