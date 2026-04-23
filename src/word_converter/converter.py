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

    HEADER_FILL_COLOR = "00B0F0"
    HEADER_FONT_COLOR = "FFFFFF"
    ORANGE_LABEL_FILL_COLOR = "ED7D31"
    GREEN_LABEL_FILL_COLOR = "92D050"
    FONT_NAME = "微軟正黑體"

    LEGACY_MAIN_HEADERS = ["編號", "功能", "細胞解碼位點", "解碼型", "健康優勢評估", "健康優勢評分"]
    NEW_MAIN_HEADERS = ["編號", "心理天賦項目", "細胞解碼位點", "解碼型", "心理潛能優勢評估", "心理潛能優勢評分"]

    ORANGE_INFO_LABELS = {"姓名", "受測者姓名", "受檢者姓名", "出生日期"}
    GREEN_INFO_LABELS = {"送檢編號", "檢體類型"}

    def __init__(
        self,
        table_header_mapping: dict[str, str] | None = None,
        cell_code_mapping: dict[tuple[str, str], str] | None = None,
        fixed_text_mapping: dict[str, str] | None = None,
    ) -> None:
        self.table_header_mapping = table_header_mapping or TABLE_HEADER_MAPPING
        self.normalized_table_header_mapping = {
            self._normalize_label(source): target for source, target in self.table_header_mapping.items()
        }
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
        self._apply_table_styles(document)
        self._apply_page_layout(document)
        self._apply_global_font(document)

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
                normalized_original = self._normalize_label(cell.text.strip())
                replacement = self.normalized_table_header_mapping.get(normalized_original)
                if replacement:
                    self._replace_cell_text(cell, replacement)

    def _convert_cell_codes(self, document: "DocxDocument") -> None:
        table_reports: list[dict[str, Any]] = []
        main_table_index: int | None = None
        replaced_count = 0
        unmapped_features: list[str] = []
        seen_unmapped: set[str] = set()

        for index, table in enumerate(document.tables):
            header_cells = table.rows[0].cells if table.rows else []
            headers = [cell.text.strip() for cell in header_cells]
            normalized_headers = [self._normalize_label(header) for header in headers]

            is_main_table = self._is_main_table_headers(normalized_headers)
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
                    self._set_row_height(row, self.GENE_ROW_HEIGHT_EMU)
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

    def _apply_table_styles(self, document: "DocxDocument") -> None:
        for table in document.tables:
            if not table.rows:
                continue

            header_cells = table.rows[0].cells
            normalized_headers = [self._normalize_label(cell.text.strip()) for cell in header_cells]
            if self._is_main_table_headers(normalized_headers):
                self._style_main_table_header_row(header_cells)

            self._style_info_label_cells(table)

    def _style_main_table_header_row(self, header_cells: list[Any]) -> None:
        for cell in header_cells:
            self._set_cell_fill(cell, self.HEADER_FILL_COLOR)
            self._style_cell_text(cell, bold=True, font_color=self.HEADER_FONT_COLOR)

    def _style_info_label_cells(self, table: Any) -> None:
        for row in table.rows:
            for cell in row.cells:
                normalized_label = self._normalize_label(cell.text.strip())
                if normalized_label in self.ORANGE_INFO_LABELS:
                    self._set_cell_fill(cell, self.ORANGE_LABEL_FILL_COLOR)
                    self._style_cell_text(cell)
                elif normalized_label in self.GREEN_INFO_LABELS:
                    self._set_cell_fill(cell, self.GREEN_LABEL_FILL_COLOR)
                    self._style_cell_text(cell)

    @classmethod
    def _is_main_table_headers(cls, normalized_headers: list[str]) -> bool:
        if len(normalized_headers) < 6:
            return False
        legacy = cls.LEGACY_MAIN_HEADERS
        new = cls.NEW_MAIN_HEADERS
        return all(
            normalized_headers[col] in {legacy[col], new[col]}
            for col in range(6)
        )

    @classmethod
    def _style_cell_text(cls, cell: Any, bold: bool | None = None, font_color: str | None = None) -> None:
        paragraphs = getattr(cell, "paragraphs", None)
        if not paragraphs:
            return

        for paragraph in paragraphs:
            runs = getattr(paragraph, "runs", [])
            if not runs and paragraph.text:
                run = paragraph.add_run(paragraph.text)
                paragraph.text = ""
                runs = [run]
            for run in runs:
                cls._set_run_font(run, bold=bold, font_color=font_color)

    @classmethod
    def _set_run_font(cls, run: Any, bold: bool | None = None, font_color: str | None = None) -> None:
        font = getattr(run, "font", None)
        if font is None:
            return

        if bold is not None:
            font.bold = bold
        font.name = cls.FONT_NAME

        if font_color:
            from docx.shared import RGBColor

            font.color.rgb = RGBColor.from_string(font_color)

        r_pr = getattr(run, "_element", None)
        if r_pr is None:
            return

        r_fonts = run._element.rPr.rFonts if run._element.rPr is not None else None
        if r_fonts is None:
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn

            if run._element.rPr is None:
                run._element.get_or_add_rPr()
            r_fonts = OxmlElement("w:rFonts")
            run._element.rPr.append(r_fonts)
            r_fonts.set(qn("w:ascii"), cls.FONT_NAME)
            r_fonts.set(qn("w:hAnsi"), cls.FONT_NAME)
            r_fonts.set(qn("w:eastAsia"), cls.FONT_NAME)
            r_fonts.set(qn("w:cs"), cls.FONT_NAME)
        else:
            from docx.oxml.ns import qn

            r_fonts.set(qn("w:ascii"), cls.FONT_NAME)
            r_fonts.set(qn("w:hAnsi"), cls.FONT_NAME)
            r_fonts.set(qn("w:eastAsia"), cls.FONT_NAME)
            r_fonts.set(qn("w:cs"), cls.FONT_NAME)

    @staticmethod
    def _set_cell_fill(cell: Any, fill_hex: str) -> None:
        tc_pr = cell._tc.get_or_add_tcPr() if hasattr(cell, "_tc") else None
        if tc_pr is None:
            return

        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        for child in tc_pr.findall(qn("w:shd")):
            tc_pr.remove(child)

        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill_hex)
        tc_pr.append(shd)

    def _apply_global_font(self, document: "DocxDocument") -> None:
        for paragraph in document.paragraphs:
            for run in getattr(paragraph, "runs", []):
                self._set_run_font(run)

        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in getattr(cell, "paragraphs", []):
                        for run in getattr(paragraph, "runs", []):
                            self._set_run_font(run)

    def _has_disclaimer_text(self, document: "DocxDocument") -> bool:
        disclaimer_tokens = ["本報告所提供之心理天賦優勢分析", "本報告依細胞分子生物學分析及統計資料"]
        for paragraph in document.paragraphs:
            text = paragraph.text.strip()
            if any(token in text for token in disclaimer_tokens):
                return True
        return False

    def _apply_page_layout(self, document: "DocxDocument") -> None:
        sections = list(getattr(document, "sections", []))
        if not sections:
            return

        target_sections = sections
        if self._has_disclaimer_text(document) and len(sections) > 1:
            target_sections = sections[:-1]

        for section in target_sections:
            section.top_margin = self.TOP_MARGIN_EMU
            section.left_margin = self.SIDE_BOTTOM_MARGIN_EMU
            section.right_margin = self.SIDE_BOTTOM_MARGIN_EMU
            section.bottom_margin = self.SIDE_BOTTOM_MARGIN_EMU

    @staticmethod
    def _set_row_height(row: Any, height_emu: int) -> None:
        row.height = height_emu
        if hasattr(row, "height_rule"):
            row.height_rule = 2  # WD_ROW_HEIGHT_RULE.EXACTLY

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

    GENE_ROW_HEIGHT_EMU = int(1.9 * 360000)
    TOP_MARGIN_EMU = int(0.75 * 360000)
    SIDE_BOTTOM_MARGIN_EMU = int(1.0 * 360000)
