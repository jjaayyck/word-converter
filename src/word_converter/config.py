"""Default mappings for legacy/new Word report conversion."""

from __future__ import annotations

# 舊版欄位名稱 -> 新版欄位名稱
TABLE_HEADER_MAPPING: dict[str, str] = {
    "姓名": "受測者姓名",
    "送檢編號": "檢測編號",
    "檢測日期": "檢測日期(西元)",
    "結果": "分析結果",
    "建議": "個人化建議",
}

# 舊版細胞解碼位點代碼 -> 新版代碼
CELL_CODE_MAPPING: dict[str, str] = {
    "A01": "TG-A01",
    "A02": "TG-A02",
    "B01": "TG-B01",
    "B02": "TG-B02",
    "C10": "TG-C10",
}

# 舊版固定文案 -> 新版固定文案
FIXED_TEXT_MAPPING: dict[str, str] = {
    "本報告僅供參考": "本報告為天賦 30 項分析結果，僅供健康管理參考。",
    "如有疑問請洽客服": "如需進一步解讀，請聯繫專屬顧問或客服中心。",
}

# 可能出現的姓名與送檢編號標籤
NAME_LABELS = ["姓名", "受測者", "受測者姓名"]
SAMPLE_ID_LABELS = ["送檢編號", "檢測編號", "樣本編號"]
