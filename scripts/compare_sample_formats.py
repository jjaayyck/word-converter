from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

LEGACY_HEADERS = ["編號", "功能", "細胞解碼位點", "解碼型", "健康優勢評估", "健康優勢評分"]
NEW_HEADERS = ["編號", "心理天賦項目", "細胞解碼位點", "解碼型", "心理潛能優勢評估", "心理潛能優勢評分"]


def normalize(value: str) -> str:
    return "".join(value.split())


def extract_tables(docx_path: Path) -> list[list[list[str]]]:
    with zipfile.ZipFile(docx_path) as archive:
        xml_bytes = archive.read("word/document.xml")

    root = ElementTree.fromstring(xml_bytes)
    tables: list[list[list[str]]] = []

    for table in root.findall(".//w:tbl", WORD_NS):
        parsed_rows: list[list[str]] = []
        for row in table.findall("./w:tr", WORD_NS):
            parsed_cells: list[str] = []
            for cell in row.findall("./w:tc", WORD_NS):
                texts = [node.text for node in cell.findall(".//w:t", WORD_NS) if node.text]
                parsed_cells.append("".join(texts).strip())
            parsed_rows.append(parsed_cells)
        tables.append(parsed_rows)

    return tables


def find_main_table(tables: list[list[list[str]]]) -> tuple[int | None, list[str]]:
    for index, table in enumerate(tables):
        if not table:
            continue
        headers = [normalize(value) for value in table[0]]
        if headers[:6] == LEGACY_HEADERS or headers[:6] == NEW_HEADERS:
            return index, table[0]
    return None, []


def classify_main_table(header: list[str]) -> str:
    normalized = [normalize(value) for value in header]
    if normalized[:6] == NEW_HEADERS:
        return "new"
    if normalized[:6] == LEGACY_HEADERS:
        return "legacy"
    return "unknown"


def build_report(docx_path: Path) -> dict[str, object]:
    tables = extract_tables(docx_path)
    main_table_index, header = find_main_table(tables)

    return {
        "file": str(docx_path),
        "table_count": len(tables),
        "main_table_index": main_table_index,
        "main_table_header": header,
        "main_table_format": classify_main_table(header),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare Word sample formats (legacy/new).")
    parser.add_argument(
        "input_dir",
        nargs="?",
        default="src/word_converter/samples/input",
        help="Directory containing .docx samples.",
    )
    parser.add_argument("--json", action="store_true", help="Print JSON output.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    sample_dir = Path(args.input_dir)
    if not sample_dir.exists():
        raise FileNotFoundError(f"找不到目錄：{sample_dir}")

    files = sorted(sample_dir.glob("*.docx"))
    if not files:
        raise FileNotFoundError(f"{sample_dir} 內沒有 .docx 檔")

    reports = [build_report(path) for path in files]

    if args.json:
        print(json.dumps(reports, ensure_ascii=False, indent=2))
        return 0

    for report in reports:
        print(f"- {Path(str(report['file'])).name}")
        print(f"  主表格式: {report['main_table_format']}")
        print(f"  主表索引: {report['main_table_index']}")
        print(f"  主表標題: {report['main_table_header']}")

    legacy_files = [Path(str(r["file"])).name for r in reports if r["main_table_format"] == "legacy"]
    new_files = [Path(str(r["file"])).name for r in reports if r["main_table_format"] == "new"]
    print("\nSummary")
    print(f"  legacy: {legacy_files}")
    print(f"  new: {new_files}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
