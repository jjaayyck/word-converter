# Word Converter

用 Python 將舊版 `.docx` 報告讀入後，進行以下轉換並輸出新版檔案：

- 擷取「姓名」與「送檢編號」
- 欄位標題 mapping
- 主表格細胞解碼位點 mapping
- 固定聲明文案替換
- 高分/低分項目固定文案替換
- 依規則輸出新檔名

## 專案結構

```text
word-converter/
├─ src/
│  └─ word_converter/
│     ├─ __init__.py
│     ├─ main.py
│     ├─ cli.py
│     ├─ converter.py
│     ├─ config.py
│     └─ mapping_loader.py
├─ scripts/
│  └─ validate_samples.py
├─ samples/
│  └─ input/
│     └─ APT-01-009297.docx
├─ requirements.txt
├─ .gitignore
└─ README.md
```

## 本機執行方式

### 1) 建立與啟用虛擬環境

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2) 執行轉換

```bash
PYTHONPATH=src python -m word_converter.main <輸入檔案.docx> -o <輸出資料夾>
```

例如：

```bash
PYTHONPATH=src python -m word_converter.main samples/input/APT-01-009297.docx -o output
```

## 輸出檔名規則

```text
台-{送檢編號}_{姓名}-天賦30項.docx
```

## 測試方式

### 1) 單元測試

```bash
python -m pytest -q
```

### 2) 兩份真實樣本驗證

本專案提供兩份真實樣本路徑（同內容、不同目錄）可做文件結構與關鍵欄位驗證：

```bash
python scripts/validate_samples.py
```

此驗證會確認每份 `.docx` 皆可解析，且包含：

- 姓名
- 送檢編號
- 細胞解碼位點

## 固定文案替換說明

- 聲明文字：
  - `本報告僅供參考` → `本報告為天賦 30 項分析結果，僅供健康管理參考。`
  - `如有疑問請洽客服` → `如需進一步解讀，請聯繫專屬顧問或客服中心。`
- 高分/低分項目：
  - `高分項目代表先天優勢` → `高分項目代表相對優勢，建議持續強化並轉化為日常表現。`
  - `低分項目代表先天不足` → `低分項目代表目前較需補強，建議透過訓練與習慣養成逐步改善。`

以上替換會同時套用在「段落」與「表格儲存格」中。
