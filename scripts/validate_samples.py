from __future__ import annotations

import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def extract_texts(docx_path: Path) -> list[str]:
    with zipfile.ZipFile(docx_path) as archive:
        xml_bytes = archive.read("word/document.xml")
    root = ElementTree.fromstring(xml_bytes)
    return [node.text.strip() for node in root.findall(".//w:t", WORD_NS) if node.text and node.text.strip()]


def validate_sample(docx_path: Path) -> None:
    if not docx_path.exists():
        raise FileNotFoundError(f"找不到樣本檔案：{docx_path}")

    texts = extract_texts(docx_path)
    blob = "\n".join(texts)
    normalized_blob = "".join(blob.split())

    required_keywords = ["姓名", "送檢編號", "細胞解碼位點"]
    missing = [keyword for keyword in required_keywords if keyword not in normalized_blob]
    if missing:
        raise ValueError(f"{docx_path} 缺少必要關鍵字：{', '.join(missing)}")

    print(f"[PASS] {docx_path}（共擷取 {len(texts)} 段文字）")


def main() -> int:
    sample_paths = [
        Path("samples/input/APT-01-009297.docx"),
        Path("src/word_converter/samples/input/APT-01-009297.docx"),
    ]
    for path in sample_paths:
        validate_sample(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
