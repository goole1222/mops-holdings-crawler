#!/usr/bin/env bash
# TSWE Holdings Crawler - macOS/Linux 執行器

set -e  # 遇到錯誤就停止

echo "========================================"
echo "TSWE Holdings Crawler - macOS/Linux 執行器"
echo "========================================"

# 檢查 Python 是否安裝
if ! command -v python3 &> /dev/null; then
    echo "❌ 錯誤：未找到 Python3，請先安裝 Python 3.10 或更高版本"
    echo ""
    echo "macOS: brew install python"
    echo "Ubuntu/Debian: sudo apt install python3 python3-venv python3-pip"
    echo "CentOS/RHEL: sudo yum install python3 python3-venv python3-pip"
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
    echo "✅ 虛擬環境建立成功"
fi

# 啟動虛擬環境
echo ""
echo "🔄 啟動虛擬環境..."
source venv/bin/activate

# 安裝/更新相依套件
echo ""
echo "📚 安裝相依套件..."
pip install -r requirements.txt --upgrade
if [ $? -ne 0 ]; then
    echo "❌ 相依套件安裝失敗"
    exit 1
fi

echo ""
echo "✅ 相依套件安裝完成"

# 檢查股票代號檔案
if [ ! -f "股票代號.txt" ]; then
    echo ""
    echo "⚠️ 警告：找不到 '股票代號.txt' 檔案"
    echo "請確保檔案存在並包含要爬取的股票代號"
    exit 1
fi

# 執行爬蟲
echo ""
echo "🚀 開始執行爬蟲..."
echo "========================================"
python src/fixed_input_crawler.py --codes-file 股票代號.txt

# 檢查執行結果
if [ $? -eq 0 ]; then
    echo ""
    echo "✅ 執行完成！"
    echo "📁 結果檔案位於 outputs/ 目錄"
    echo "📋 詳細日誌位於 logs/ 目錄"
else
    echo ""
    echo "❌ 執行過程中發生錯誤，請檢查日誌檔案"
    exit 1
fi

echo ""
echo "========================================"
echo "Done"