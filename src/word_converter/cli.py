"""CLI entrypoint for word report conversion."""

from __future__ import annotations

import argparse
from pathlib import Path

from .config import CELL_CODE_MAPPING, FIXED_TEXT_MAPPING, TABLE_HEADER_MAPPING
from .converter import WordReportConverter
from .mapping_loader import load_mapping_overrides

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="舊版 Word 報告轉新版格式工具")
    parser.add_argument("input", type=Path, help="輸入 .docx 檔案路徑或資料夾路徑")
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


def _is_image_file(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTENSIONS


def _collect_input_files(input_path: Path) -> tuple[list[Path], list[Path]]:
    if input_path.is_dir():
        docx_files: list[Path] = []
        skipped_images: list[Path] = []
        for path in sorted(input_path.iterdir()):
            if not path.is_file():
                continue
            suffix = path.suffix.lower()
            if suffix == ".docx":
                docx_files.append(path)
            elif _is_image_file(path):
                skipped_images.append(path)
        return docx_files, skipped_images

    if _is_image_file(input_path):
        return [], [input_path]

    return [input_path], []


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
    input_path = args.input
    results = []
    input_files, skipped_images = _collect_input_files(input_path)

    for image_path in skipped_images:
        print(f"略過圖片檔: {image_path}")

    if not input_files:
        print("沒有可處理的 .docx 檔案，已結束。")
        return

    for input_file in input_files:
        results.append(converter.convert(input_file, args.output_dir))

    print("轉換完成")
    print(f"共處理 {len(results)} 份文件")
    for result in results:
        print(f"- 姓名: {result.name} | 送檢編號: {result.sample_id} | 輸出檔案: {result.output_path}")


if __name__ == "__main__":
    main()
