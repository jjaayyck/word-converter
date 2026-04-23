"""Core conversion logic for legacy Word reports."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument
    from docx.table import _Cell
from .config import (
    CELL_CODE_MAPPING,
    FIXED_TEXT_MAPPING,
    NAME_LABELS,
    SAMPLE_ID_LABELS,
    TABLE_HEADER_MAPPING,
)


@dataclass(slots=True)
class ConversionResult:
    input_path: Path
    output_path: Path
    name: str
    sample_id: str


class WordReportConverter:
    """Convert an old-format Word report to the new format."""

    def __init__(
        self,
        table_header_mapping: dict[str, str] | None = None,
        cell_code_mapping: dict[tuple[str, str], str] | None = None,
        fixed_text_mapping: dict[str, str] | None = None,
    ) -> None:
        self.table_header_mapping = table_header_mapping or TABLE_HEADER_MAPPING
        self.cell_code_mapping = cell_code_mapping or CELL_CODE_MAPPING
        self.fixed_text_mapping = fixed_text_mapping or FIXED_TEXT_MAPPING
        self.last_cell_code_report: dict[str, Any] = {}

    def convert(self, input_path: str | Path, output_dir: str | Path) -> ConversionResult:
        input_file = Path(input_path)
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        document = self._load_document(input_file)
        name, sample_id = self._extract_identity(document)
        self._convert_table_headers(document)
        self._convert_cell_codes(document)
        self._apply_fixed_text(document)

        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)
        output_path = output_dir_path / self._build_output_filename(sample_id=sample_id, name=name)
        document.save(str(output_path))

        return ConversionResult(
            input_path=input_file,
            output_path=output_path,
            name=name,
            sample_id=sample_id,
        )

    @staticmethod
    def _load_document(input_file: Path) -> Any:
        try:
            from docx import Document
        except ModuleNotFoundError as exc:
            raise ModuleNotFoundError(
                "python-docx is required. Install dependencies with: pip install -r requirements.txt"
            ) from exc

        return Document(str(input_file))

    def _extract_identity(self, document: "DocxDocument") -> tuple[str, str]:
        name: str | None = None
        sample_id: str | None = None

        for table in document.tables:
            t_name, t_id = self._extract_identity_from_table(table)
            name = name or t_name
            sample_id = sample_id or t_id
            if name and sample_id:
                return name, sample_id

        full_text = "\n".join(p.text for p in document.paragraphs if p.text)
        name = name or self._extract_by_labels(full_text, NAME_LABELS)
        sample_id = sample_id or self._extract_by_labels(full_text, SAMPLE_ID_LABELS)

        if name and sample_id:
            return name, sample_id

        missing = []
        if not name:
            missing.append("姓名")
        if not sample_id:
            missing.append("送檢編號")
        raise ValueError(f"無法從文件中擷取欄位：{', '.join(missing)}")

    @staticmethod
    def _normalize_label(value: str) -> str:
        return "".join(value.split())

    @staticmethod
    def _extract_by_labels(text: str, labels: list[str]) -> str | None:
        for label in labels:
            pattern = rf"{re.escape(label)}\s*[:：]\s*([^\n\r\t,，]+)"
            match = re.search(pattern, text)
            if match:
                return match.group(1).strip()
        return None

    def _extract_identity_from_table(self, table: "Table") -> tuple[str | None, str | None]:
        name: str | None = None
        sample_id: str | None = None

        normalized_name_labels = {self._normalize_label(label) for label in NAME_LABELS}
        normalized_sample_id_labels = {self._normalize_label(label) for label in SAMPLE_ID_LABELS}

        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            for idx, cell_text in enumerate(cells):
                normalized_cell_text = self._normalize_label(cell_text)

                if normalized_cell_text in normalized_name_labels:
                    name = name or self._first_non_empty_to_right(cells, idx)

                if normalized_cell_text in normalized_sample_id_labels:
                    sample_id = sample_id or self._first_non_empty_to_right(cells, idx)

                if name and sample_id:
                    return name, sample_id
        return name, sample_id

    @staticmethod
    def _first_non_empty_to_right(cells: list[str], idx: int) -> str | None:
        for cell_text in cells[idx + 1 :]:
            value = cell_text.strip()
            if value:
                return value
        return None

    def _convert_table_headers(self, document: "DocxDocument") -> None:
        for table in document.tables:
            if not table.rows:
                continue
            for cell in table.rows[0].cells:
                original = cell.text.strip()
                if original in self.table_header_mapping:
                    self._replace_cell_text(cell, self.table_header_mapping[original])

    def _convert_cell_codes(self, document: "DocxDocument") -> None:
        expected_main_headers = ["編號", "功能", "細胞解碼位點", "解碼型", "健康優勢評估", "健康優勢評分"]
        table_reports: list[dict[str, Any]] = []
        main_table_index: int | None = None
        replaced_count = 0
        unmapped_features: list[str] = []
        seen_unmapped: set[str] = set()

        for index, table in enumerate(document.tables):
            header_cells = table.rows[0].cells if table.rows else []
            headers = [cell.text.strip() for cell in header_cells]
            normalized_headers = [self._normalize_label(header) for header in headers]

            is_main_table = len(normalized_headers) >= 6 and all(
                normalized_headers[col] == expected_main_headers[col] for col in range(6)
            )
            table_reports.append(
                {
                    "table_index": index,
                    "headers": headers,
                    "is_main_table": is_main_table,
                }
            )

            if not is_main_table or main_table_index is not None:
                continue

            main_table_index = index
            for row in table.rows[1:]:
                if len(row.cells) < 3:
                    continue

                feature_name = row.cells[1].text.strip()
                old_code = row.cells[2].text.strip()
                if not feature_name or not old_code:
                    continue

                new_code = self.cell_code_mapping.get((feature_name, old_code))
                if new_code:
                    self._replace_cell_text(row.cells[2], new_code)
                    replaced_count += 1
                elif feature_name not in seen_unmapped:
                    seen_unmapped.add(feature_name)
                    unmapped_features.append(feature_name)

        self.last_cell_code_report = {
            "table_count": len(document.tables),
            "tables": table_reports,
            "main_table_index": main_table_index,
            "replaced_count": replaced_count,
            "unmapped_features": unmapped_features,
        }

    def _apply_fixed_text(self, document: "DocxDocument") -> None:
        for paragraph in document.paragraphs:
            replaced = self._replace_fixed_text(paragraph.text)
            if replaced != paragraph.text:
                paragraph.text = replaced

        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    replaced = self._replace_fixed_text(cell.text)
                    if replaced != cell.text:
                        self._replace_cell_text(cell, replaced)

    def _replace_fixed_text(self, text: str) -> str:
        replaced = text
        for old, new in self.fixed_text_mapping.items():
            replaced = replaced.replace(old, new)
        return replaced

    @staticmethod
    def _replace_cell_text(cell: "_Cell", text: str) -> None:
        if hasattr(cell, "paragraphs") and cell.paragraphs:
            cell.paragraphs[0].text = text
            for paragraph in cell.paragraphs[1:]:
                paragraph.text = ""
        else:
            cell.text = text

    @staticmethod
    def _sanitize_filename_part(value: str) -> str:
        return re.sub(r'[\\/:*?"<>|\s]+', "", value)

    def _build_output_filename(self, sample_id: str, name: str) -> str:
        sid = self._sanitize_filename_part(sample_id)
        person = self._sanitize_filename_part(name)
        return f"台-{sid}_{person}-天賦30項.docx"
