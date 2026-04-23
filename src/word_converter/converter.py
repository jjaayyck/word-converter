"""Core conversion logic for legacy Word reports."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument
    from docx.table import _Cell, Table
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
        cell_code_mapping: dict[str, str] | None = None,
        fixed_text_mapping: dict[str, str] | None = None,
    ) -> None:
        self.table_header_mapping = table_header_mapping or TABLE_HEADER_MAPPING
        self.cell_code_mapping = cell_code_mapping or CELL_CODE_MAPPING
        self.fixed_text_mapping = fixed_text_mapping or FIXED_TEXT_MAPPING

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
        full_text = "\n".join(p.text for p in document.paragraphs if p.text)
        name = self._extract_by_labels(full_text, NAME_LABELS)
        sample_id = self._extract_by_labels(full_text, SAMPLE_ID_LABELS)

        if name and sample_id:
            return name, sample_id

        for table in document.tables:
            t_name, t_id = self._extract_identity_from_table(table)
            name = name or t_name
            sample_id = sample_id or t_id
            if name and sample_id:
                return name, sample_id

        missing = []
        if not name:
            missing.append("姓名")
        if not sample_id:
            missing.append("送檢編號")
        raise ValueError(f"無法從文件中擷取欄位：{', '.join(missing)}")

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

        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            for idx, cell_text in enumerate(cells):
                if cell_text in NAME_LABELS and idx + 1 < len(cells):
                    name = name or cells[idx + 1].strip()
                if cell_text in SAMPLE_ID_LABELS and idx + 1 < len(cells):
                    sample_id = sample_id or cells[idx + 1].strip()
        return name, sample_id

    def _convert_table_headers(self, document: "DocxDocument") -> None:
        for table in document.tables:
            if not table.rows:
                continue
            for cell in table.rows[0].cells:
                original = cell.text.strip()
                if original in self.table_header_mapping:
                    self._replace_cell_text(cell, self.table_header_mapping[original])

    def _convert_cell_codes(self, document: "DocxDocument") -> None:
        code_pattern = re.compile(r"\b([A-Z]\d{2}|[A-Z]\d{2,3})\b")

        def replace_codes(text: str) -> str:
            def _repl(match: re.Match[str]) -> str:
                code = match.group(1)
                return self.cell_code_mapping.get(code, code)

            return code_pattern.sub(_repl, text)

        for paragraph in document.paragraphs:
            new_text = replace_codes(paragraph.text)
            if new_text != paragraph.text:
                paragraph.text = new_text

        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    new_text = replace_codes(cell.text)
                    if new_text != cell.text:
                        self._replace_cell_text(cell, new_text)

    def _apply_fixed_text(self, document: "DocxDocument") -> None:
        for paragraph in document.paragraphs:
            text = paragraph.text
            for old, new in self.fixed_text_mapping.items():
                text = text.replace(old, new)
            if text != paragraph.text:
                paragraph.text = text

    @staticmethod
    def _replace_cell_text(cell: "_Cell", text: str) -> None:
        if cell.paragraphs:
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
