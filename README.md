# MOPS Holdings Crawler

[中文版說明](README_zh.md)

A Python script that fetches director/shareholder **current holdings** from TWSE MOPS and exports an Excel file.

---

## 1. Requirements
- Python 3.9+
- Google Chrome + matching ChromeDriver (if Selenium is used)
- OS: Windows / macOS / Linux

---

## 2. Installation
```bash
cd mops-holdings-crawler
pip install -r requirements.txt
(Optional) Virtual environment:

bash
複製程式碼
python -m venv .venv
.venv\Scripts\activate   # Windows
source .venv/bin/activate  # macOS/Linux
pip install -r requirements.txt
3. Prepare Input
Create a file 股票代號.txt in the project root, one stock code per line:

yaml
複製程式碼
2330
1101
2603
Invalid codes like 0000 will be logged in the "failure" sheet.

4. Run
bash
複製程式碼
python fixed_input_crawler.py --codes-file 股票代號.txt

5. Output
Excel file like: 董監事持股_合併_YYYYMMDD.xlsx

合併 (merged) sheet: current holdings

失敗記錄 (failures) sheet: invalid/unreachable codes

The system writes to Excel while running and records processed codes in processed_codes.txt.
👉 If you want to re-run later, clear processed_codes.txt first, otherwise the script will skip completed codes.