# MOPS Holdings Crawler

這是一個 Python 爬蟲，專門抓取台灣證交所公開資訊觀測站 (MOPS) 的「董監事持股餘額」資料，並輸出成 Excel。

## 功能
- 自動打開 MOPS 網站，輸入公司代號
- 支援 CSV 下載或表格解析
- 輸出結果至 Excel（合併表 + 失敗記錄）
- 自動續跑功能，避免重複處理

## 使用方法
1. 安裝 Python 3.9+ 與套件
   ```bash
   pip install -r requirements.txt
建立 股票代號.txt，一行一個代號

執行：


python fixed_input_crawler.py --codes-file 股票代號.txt