"""CLI entrypoint for word report conversion."""

from __future__ import annotations

import argparse
from pathlib import Path

from .converter import WordReportConverter


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="舊版 Word 報告轉新版格式工具")
    parser.add_argument("input", type=Path, help="輸入 .docx 檔案路徑")
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="輸出資料夾（預設：output）",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    converter = WordReportConverter()
    result = converter.convert(args.input, args.output_dir)

    print("轉換完成")
    print(f"姓名: {result.name}")
    print(f"送檢編號: {result.sample_id}")
    print(f"輸出檔案: {result.output_path}")


if __name__ == "__main__":
    main()
