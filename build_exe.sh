#!/usr/bin/env bash
# PyInstaller 打包工具 (macOS/Linux)

set -e

echo "========================================"
echo "PyInstaller 打包工具 (macOS/Linux)"
echo "========================================"

# 檢查 Python 是否安裝
if ! command -v python3 &> /dev/null; then
    echo "❌ 錯誤：未找到 Python3，請先安裝 Python 3.10 或更高版本"
    exit 1
fi

echo "✅ Python3 已安裝"
python3 --version

# 檢查虛擬環境是否存在
if [ ! -d "venv" ]; then
    echo ""
    echo "📦 建立虛擬環境..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "❌ 建立虛擬環境失敗"
        exit 1
    fi
fi

# 啟動虛擬環境
echo ""
echo "🔄 啟動虛擬環境..."
source venv/bin/activate

# 安裝相依套件
echo ""
echo "📚 安裝相依套件（包含 PyInstaller）..."
pip install -r requirements.txt --upgrade
if [ $? -ne 0 ]; then
    echo "❌ 相依套件安裝失敗"
    exit 1
fi

# 清理舊的打包結果
echo ""
echo "🧹 清理舊的打包結果..."
rm -rf dist/ build/ *.spec

# 使用 PyInstaller 打包
echo ""
echo "📦 開始使用 PyInstaller 打包..."
echo "========================================"

pyinstaller \
    --onefile \
    --windowed \
    --name="TsweHoldingsCrawler" \
    --add-data="config.yaml:." \
    --add-data="股票代號.txt:." \
    --hidden-import="selenium" \
    --hidden-import="pandas" \
    --hidden-import="openpyxl" \
    --hidden-import="requests" \
    --hidden-import="yaml" \
    --clean \
    src/fixed_input_crawler.py

if [ $? -ne 0 ]; then
    echo "❌ PyInstaller 打包失敗"
    exit 1
fi

echo ""
echo "✅ 打包完成！"
echo ""
echo "📁 執行檔位置：dist/TsweHoldingsCrawler"
echo ""
echo "⚠️ 重要提醒："
echo "  1. 執行檔仍需要本機安裝 Google Chrome 瀏覽器"
echo "  2. 首次執行時，Selenium Manager 會自動下載 ChromeDriver"
echo "  3. 請將 config.yaml 和 股票代號.txt 放在執行檔同目錄"
echo "  4. 建議在執行檔同目錄建立 downloads、outputs、logs 資料夾"
echo ""
echo "========================================"
echo "Done"