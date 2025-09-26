# 董監事持股爬蟲

[English Version](README.md)

本工具用 Python 從公開資訊觀測站（MOPS）擷取 **目前持股**，並輸出為 Excel。

---

## 1. 系統需求
- Python 3.9+
- Google Chrome 與對應版本 ChromeDriver（若使用 Selenium）
- Windows / macOS / Linux

---

## 2. 安裝
```bash
cd mops-holdings-crawler
pip install -r requirements.txt
```

（可選）使用虛擬環境：
```bash
python -m venv .venv
.venv\Scripts\activate   # Windows
source .venv/bin/activate  # macOS/Linux
pip install -r requirements.txt
```

---

## 3. 建立輸入檔
在專案根目錄建立 **`股票代號.txt`**（UTF-8 無 BOM），每行一個股票代號：

```text
2330
1101
2603
```

若輸入無效代號（例如 `0000`），會被記錄在「失敗記錄」分頁。

---

## 4. 執行
```bash
python fixed_input_crawler.py --codes-file 股票代號.txt
```

若程式在 `src/` 資料夾，改成：
```bash
python src/fixed_input_crawler.py --codes-file 股票代號.txt
```

---

## 5. 輸出
- 產生 Excel，例如：`董監事持股_合併_YYYYMMDD.xlsx`
  - **合併**：各公司目前持股
  - **失敗記錄**：查不到或錯誤的代號
- 系統會**邊跑邊寫入 Excel**，同時把已完成的代號記錄在 **`processed_codes.txt`**  
  👉 若全部跑完後，過一段時間想要重新執行，請先**清空 `processed_codes.txt`**，否則會直接跳過已完成的代號。

---

## 6. 常見問題
- **亂碼或讀不到代號檔**：請將 `股票代號.txt` 儲存為 UTF-8（無 BOM）。  
- **Excel 抓到「選任時持股」**：本版本已修正，只會讀「目前持股」。  
- **ChromeDriver 問題**：請下載與 Chrome 相同版本的 ChromeDriver，並放到專案根目錄或 PATH。

---

## 授權
MIT License（或依需要調整）
