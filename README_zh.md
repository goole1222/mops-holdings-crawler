MOPS Holdings Crawler / 董監事持股爬蟲

A Python script that fetches director/shareholder current holdings from TWSE MOPS and exports an Excel file.
本工具用 Python 從公開資訊觀測站（MOPS）擷取目前持股資料，並輸出 Excel。

1. Requirements / 系統需求

Python 3.9+（必須）

建議安裝 Google Chrome 與對應版本 ChromeDriver（若你的版本使用 Selenium）

作業系統：Windows / macOS / Linux 皆可

2. Install Python & packages / 安裝 Python 與套件
# 進到專案資料夾
cd mops-holdings-crawler

# 安裝相依套件（建議先確認已安裝 Python 3.9+）
pip install -r requirements.txt


進階（可選）：為避免污染系統環境，建議使用虛擬環境
Windows：

python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt


macOS/Linux：

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

3. Prepare input / 準備輸入檔

在專案根目錄建立 股票代號.txt（UTF-8 無 BOM），每行一個代號：

2330
1101
2603
# 允許空白行與註解（以 # 開頭的行會被忽略）


小提醒：若放入無效代號（例如 0000），程式不會當機，會把它記在輸出檔案的「失敗記錄」分頁。

4. Run / 執行
python fixed_input_crawler.py --codes-file 股票代號.txt


常用參數（若程式支援）：

--out <filename.xlsx> 自訂輸出檔名

--retry 1 失敗時重試次數

--throttle 1.5 每檔間隔秒數

5. Output / 輸出

產生 Excel，例如：董監事持股_合併_YYYYMMDD.xlsx

合併 工作表：彙整各公司「目前持股」

失敗記錄 工作表：查不到或錯誤的代號（如 0000）

系統會邊跑邊存入 Excel，同時用 processed_codes.txt 記錄已完成的代號。
👉 因此，如果全部跑完後，想要過一段時間再重跑一次，請記得先把 processed_codes.txt 清空，不然系統會直接跳過已完成的代號。

6. Troubleshooting / 常見問題

Q1. 執行後顯示亂碼或 .bat 無法執行？

請確保批次檔（若使用）以 UTF-8（無 BOM）、CRLF 換行儲存。

直接用命令列執行 python fixed_input_crawler.py --codes-file 股票代號.txt 最穩。

Q2. 程式讀不到 股票代號.txt？

確認檔案與程式在同一層資料夾，或用相對/絕對路徑指定。

檔案請用 UTF-8（無 BOM） 儲存；不要有奇怪字元。

第一列若是標題（例如「代號」），請移除或讓程式跳過。

Q3. Excel 裡抓到「選任時持股」而不是「目前持股」？

本版本已改為只認 「目前持股」系列欄位，若仍發生請附上日誌與範例檔以便排查。

Q4. Chrome/ChromeDriver 問題（若使用 Selenium）

請安裝與 Chrome 版本相同的 ChromeDriver，並放在專案根目錄或加入 PATH。

macOS/Linux 記得 chmod +x chromedriver。

7. Notes / 其他說明

請尊重目標網站的使用條款與爬取頻率；必要時調整 --throttle。

若 MOPS 網站改版，可能需要更新選取器或下載邏輯。

本專案預設只輸出目前持股欄位，避免誤抓「選任時持股」