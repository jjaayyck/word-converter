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
│     └─ converter.py
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

範例：

```bash
PYTHONPATH=src python -m word_converter.cli ./legacy-report.docx -o ./output
```

## 目前內建轉換規則

規則定義在 `src/word_converter/config.py`：

- `TABLE_HEADER_MAPPING`: 表格欄位名稱映射。
- `CELL_CODE_MAPPING`: 細胞解碼位點代碼映射。
- `FIXED_TEXT_MAPPING`: 固定文案替換。

你可以直接修改 `config.py` 來符合你的正式對照表。

## 功能對應需求

- ✅ 讀取 `.docx`
- ✅ 擷取姓名與送檢編號（先掃描段落，再掃描表格）
- ✅ 產生新檔名（`台-{送檢編號}_{姓名}-天賦30項.docx`）
- ✅ 轉換表格欄位名稱
- ✅ 轉換細胞解碼位點代碼
- ✅ 套用新版固定文案
- ✅ 輸出新的 `.docx`

## 注意事項

- 本工具以文字內容替換為主；若你的報告有複雜樣式（run-level 格式、特殊合併表格等），可能需要再加強保留格式的策略。
- 若你提供實際「欄位對照表 / 代碼對照表 / 固定文案」檔案，我可以幫你改成讀取外部 JSON/CSV 並完整對齊正式規格。
