"""CLI entrypoint for word report conversion."""

from __future__ import annotations

import argparse
from pathlib import Path

from .config import CELL_CODE_MAPPING, FIXED_TEXT_MAPPING, TABLE_HEADER_MAPPING
from .converter import WordReportConverter
from .mapping_loader import load_mapping_overrides


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
    parser.add_argument(
        "-c",
        "--config",
        type=Path,
        help="可選：JSON 對照表設定檔（覆蓋預設 mapping）",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()

    table_mapping = dict(TABLE_HEADER_MAPPING)
    code_mapping = dict(CELL_CODE_MAPPING)
    text_mapping = dict(FIXED_TEXT_MAPPING)

    if args.config:
        overrides = load_mapping_overrides(args.config)
        table_mapping.update(overrides["table_header_mapping"])
        code_mapping.update(overrides["cell_code_mapping"])
        text_mapping.update(overrides["fixed_text_mapping"])

    converter = WordReportConverter(
        table_header_mapping=table_mapping,
        cell_code_mapping=code_mapping,
        fixed_text_mapping=text_mapping,
    )
    result = converter.convert(args.input, args.output_dir)

    print("轉換完成")
    print(f"姓名: {result.name}")
    print(f"送檢編號: {result.sample_id}")
    print(f"輸出檔案: {result.output_path}")


if __name__ == "__main__":
    main()
