# Word Converter

用 Python 將舊版 `.docx` 報告讀入後，進行以下轉換並輸出新版檔案：

- 擷取「姓名」與「送檢編號」
- 欄位標題 mapping
- 主表格細胞解碼位點 mapping
- 檢測結果主表格（有基因代碼列）列高統一為 1.9 公分
- 固定聲明文案替換
- 高分/低分項目固定文案替換
- 頁面邊界調整（上 0.75 公分，左/右/下 1 公分）
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
PYTHONPATH=src python -m word_converter.main <輸入檔案.docx或資料夾> -o <輸出資料夾>
```

例如：

```bash
PYTHONPATH=src python -m word_converter.main samples/input/APT-01-009297.docx -o output
```

一次處理整個資料夾內所有 `.docx`：

```bash
PYTHONPATH=src python -m word_converter.main samples/input -o output
```


## 比對舊版/新版樣本格式

若你在 `src/word_converter/samples/input` 放了多份樣本，可先用下面指令快速判斷哪些是舊版格式、哪些已是新版格式：

```bash
python scripts/compare_sample_formats.py
```

若要輸出 JSON（方便後續程式化處理）：

```bash
python scripts/compare_sample_formats.py --json
```

主表格欄位判斷規則：

- **舊版**：`編號 / 功能 / 細胞解碼位點 / 解碼型 / 健康優勢評估 / 健康優勢評分`
- **新版**：`編號 / 心理天賦項目 / 細胞解碼位點 / 解碼型 / 心理潛能優勢評估 / 心理潛能優勢評分`

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
  - `本報告依細胞分子生物學分析及統計資料...請諮詢專業醫師。` →
    `本報告所提供之心理天賦優勢分析...請洽詢專業心理師或醫療人員。`
- 高分/低分項目：
  - `高分項目代表先天優勢` → `高分項目代表相對優勢，建議持續強化並轉化為日常表現。`
  - `低分項目代表先天不足` → `低分項目代表目前較需補強，建議透過訓練與習慣養成逐步改善。`

以上替換會同時套用在「段落」與「表格儲存格」中。

## 版面調整規則

- 檢測結果主表格中，所有有完成基因代碼 mapping 的資料列，列高固定為 **1.9 公分**。
- 文件頁面邊界統一調整為：
  - 上：**0.75 公分**
  - 左：**1 公分**
  - 右：**1 公分**
  - 下：**1 公分**
