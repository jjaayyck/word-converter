# Word Converter MVP v1

用 Python 將舊版 `.docx` 報告讀入，擷取「姓名」與「送檢編號」，依規則輸出新檔名並產生新的 `.docx`。

## 本版（MVP v1）完成項目

- Python 專案結構
- `requirements.txt`
- `README.md`
- 主程式入口
- 讀取 docx
- 擷取姓名與送檢編號
- 產生輸出檔名
- 輸出新的 docx

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
PYTHONPATH=src python -m word_converter.main <輸入檔案.docx> -o <輸出資料夾>
```

## 輸出檔名規則

```text
台-{送檢編號}_{姓名}-天賦30項.docx
```

## 備註

- 本版先專注 MVP 核心流程（讀取、擷取、命名、輸出）。
- 欄位轉換 / 細胞代碼對照 / 固定文案替換可在下一版接入。
