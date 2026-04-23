"""Validate extraction rules against a real sample .docx file."""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path

from word_converter.config import NAME_LABELS, SAMPLE_ID_LABELS
from word_converter.converter import WordReportConverter

TARGET_FILENAME = "APT-01-009297.docx"


def extract_text_from_docx(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    return " ".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="驗證樣本檔案的姓名/送檢編號擷取結果")
    parser.add_argument("--input", type=Path, default=None, help="樣本 .docx 路徑（可省略，將自動搜尋）")
    return parser


def resolve_sample_path(input_path: Path | None) -> Path | None:
    if input_path is not None:
        return input_path if input_path.exists() else None

    matches = sorted(Path(".").rglob(TARGET_FILENAME))
    return matches[0] if matches else None


def main() -> None:
    args = build_parser().parse_args()
    sample_file = resolve_sample_path(args.input)

    if sample_file is None:
        print(f"Sample file not found: {TARGET_FILENAME}")
        print("Hint: place file under samples/input/ or pass --input <path>")
        return

    converter = WordReportConverter()
    text = extract_text_from_docx(sample_file)

    name = converter._extract_by_labels(text, NAME_LABELS) or "<NOT_FOUND>"
    sample_id = converter._extract_by_labels(text, SAMPLE_ID_LABELS) or "<NOT_FOUND>"

    if name != "<NOT_FOUND>" and sample_id != "<NOT_FOUND>":
        output_filename = converter._build_output_filename(sample_id=sample_id, name=name)
    else:
        output_filename = "<CANNOT_BUILD>"

    print(f"input_file: {sample_file.as_posix()}")
    print(f"name: {name}")
    print(f"sample_id: {sample_id}")
    print(f"output_filename: {output_filename}")


if __name__ == "__main__":
    main()
