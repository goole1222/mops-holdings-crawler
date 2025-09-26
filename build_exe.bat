@echo off
chcp 65001 > nul
echo ========================================
echo PyInstaller 打包工具 (Windows)
echo ========================================

REM 檢查 Python 是否安裝
python --version > nul 2>&1
if errorlevel 1 (
    echo ❌ 錯誤：未找到 Python，請先安裝 Python 3.10 或更高版本
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
)

REM 啟動虛擬環境
echo.
echo 🔄 啟動虛擬環境...
call venv\Scripts\activate.bat

REM 安裝相依套件
echo.
echo 📚 安裝相依套件（包含 PyInstaller）...
pip install -r requirements.txt --upgrade
if errorlevel 1 (
    echo ❌ 相依套件安裝失敗
    pause
    exit /b 1
)

REM 清理舊的打包結果
echo.
echo 🧹 清理舊的打包結果...
if exist "dist\" rmdir /s /q dist
if exist "build\" rmdir /s /q build
if exist "*.spec" del *.spec

REM 使用 PyInstaller 打包
echo.
echo 📦 開始使用 PyInstaller 打包...
echo ========================================

pyinstaller ^
    --onefile ^
    --windowed ^
    --name="TsweHoldingsCrawler" ^
    --add-data="config.yaml;." ^
    --add-data="股票代號.txt;." ^
    --hidden-import="selenium" ^
    --hidden-import="pandas" ^
    --hidden-import="openpyxl" ^
    --hidden-import="requests" ^
    --hidden-import="yaml" ^
    --clean ^
    src/fixed_input_crawler.py

if errorlevel 1 (
    echo ❌ PyInstaller 打包失敗
    pause
    exit /b 1
)

echo.
echo ✅ 打包完成！
echo.
echo 📁 執行檔位置：dist\TsweHoldingsCrawler.exe
echo.
echo ⚠️ 重要提醒：
echo   1. 執行檔仍需要本機安裝 Google Chrome 瀏覽器
echo   2. 首次執行時，Selenium Manager 會自動下載 ChromeDriver
echo   3. 請將 config.yaml 和 股票代號.txt 放在執行檔同目錄
echo   4. 建議在執行檔同目錄建立 downloads、outputs、logs 資料夾
echo.
echo ========================================
echo 按任意鍵結束...
pause > nul