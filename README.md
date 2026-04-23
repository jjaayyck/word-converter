# Word Converter（舊版報告轉新版）

這是一個使用 **Python + python-docx** 的專案，用來把舊版 `.docx` Word 報告轉成新版格式。

## 實作計畫

1. 建立專案資料夾結構與必要檔案（`requirements.txt`、`README.md`、`.gitignore`）。
2. 實作 `.docx` 讀取與基本 CLI（指定輸入檔與輸出資料夾）。
3. 從內文/表格擷取「姓名」與「送檢編號」。
4. 依規則產生新檔名：`台-{送檢編號}_{姓名}-天賦30項.docx`。
5. 轉換表格欄位名稱（舊欄位 -> 新欄位）。
6. 依對照表轉換細胞解碼位點代碼。
7. 套用新版固定文案後輸出新 `.docx`。

## 專案結構

```text
word-converter/
├─ src/
│  └─ word_converter/
│     ├─ __init__.py
│     ├─ cli.py
│     ├─ config.py
│     ├─ converter.py
│     └─ mapping_loader.py
├─ requirements.txt
├─ .gitignore
└─ README.md
```

## 安裝

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 使用方式

```bash
PYTHONPATH=src python -m word_converter.cli <輸入檔案.docx> -o <輸出資料夾>
```

### 使用外部 JSON 對照表（建議正式上線時）

```bash
PYTHONPATH=src python -m word_converter.cli ./legacy-report.docx -o ./output -c ./mapping.json
```

`mapping.json` 範例：

```json
{
  "table_header_mapping": {
    "姓名": "受測者姓名"
  },
  "cell_code_mapping": {
    "A01": "TG-A01"
  },
  "fixed_text_mapping": {
    "本報告僅供參考": "本報告為天賦 30 項分析結果，僅供健康管理參考。"
  }
}
```

## 功能對應需求

- ✅ 讀取 `.docx`
- ✅ 擷取姓名與送檢編號（先掃描段落，再掃描表格）
- ✅ 產生新檔名（`台-{送檢編號}_{姓名}-天賦30項.docx`）
- ✅ 轉換表格欄位名稱
- ✅ 轉換細胞解碼位點代碼
- ✅ 套用新版固定文案
- ✅ 輸出新的 `.docx`

## MVP 判斷

目前屬於 **可用 MVP**（可處理標準化模板文件），但建議上線前補強：

1. 以真實樣本檔建立測試案例（含多種版型）。
2. 若需保留複雜 run-level 樣式，需改為更細粒度文字替換策略。
3. 補齊批次處理與錯誤報表（例如輸出轉換失敗清單）。

## 注意事項

- `python-docx` 採段落/儲存格文字重寫時，可能影響部分原始文字樣式。
- 若你提供實際「欄位對照表 / 代碼對照表 / 固定文案」，可直接用 `-c mapping.json` 覆蓋預設規則。
