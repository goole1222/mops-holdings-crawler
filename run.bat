@echo off
chcp 65001 > nul
echo ========================================
echo TSWE Holdings Crawler - Windows 執行器
echo ========================================

REM 檢查 Python 是否安裝
python --version > nul 2>&1
if errorlevel 1 (
    echo ❌ 錯誤：未找到 Python，請先安裝 Python 3.10 或更高版本
    echo.
    echo 請到 https://python.org 下載並安裝 Python
    pause
    exit /b 1
)

echo ✅ Python 已安裝
python --version

REM 檢查虛擬環境是否存在
if not exist "venv\" (
    echo.
    echo 📦 建立虛擬環境...
    python -m venv venv
    if errorlevel 1 (
        echo ❌ 建立虛擬環境失敗
        pause
        exit /b 1
    )
    echo ✅ 虛擬環境建立成功
)

REM 啟動虛擬環境
echo.
echo 🔄 啟動虛擬環境...
call venv\Scripts\activate.bat

REM 安裝/更新相依套件
echo.
echo 📚 安裝相依套件...
pip install -r requirements.txt --upgrade
if errorlevel 1 (
    echo ❌ 相依套件安裝失敗
    pause
    exit /b 1
)

echo.
echo ✅ 相依套件安裝完成

REM 檢查股票代號檔案
if not exist "股票代號.txt" (
    echo.
    echo ⚠️ 警告：找不到 "股票代號.txt" 檔案
    echo 請確保檔案存在並包含要爬取的股票代號
    pause
    exit /b 1
)

REM 執行爬蟲
echo.
echo 🚀 開始執行爬蟲...
echo ========================================
python src/fixed_input_crawler.py --codes-file 股票代號.txt

REM 檢查執行結果
if errorlevel 1 (
    echo.
    echo ❌ 執行過程中發生錯誤，請檢查日誌檔案
) else (
    echo.
    echo ✅ 執行完成！
    echo 📁 結果檔案位於 outputs\ 目錄
    echo 📋 詳細日誌位於 logs\ 目錄
)

echo.
echo ========================================
echo 按任意鍵結束...
pause > nul