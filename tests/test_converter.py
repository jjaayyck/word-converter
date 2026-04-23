from __future__ import annotations

from dataclasses import dataclass

from src.word_converter.converter import WordReportConverter


@dataclass
class FakeCell:
    text: str


@dataclass
class FakeRow:
    cells: list[FakeCell]


@dataclass
class FakeTable:
    rows: list[FakeRow]


@dataclass
class FakeParagraph:
    text: str


@dataclass
class FakeDocument:
    paragraphs: list[FakeParagraph]
    tables: list[FakeTable]


def _build_realistic_table() -> FakeTable:
    return FakeTable(
        rows=[
            FakeRow(
                cells=[
                    FakeCell("姓   名"),
                    FakeCell("王曉明"),
                    FakeCell("性別"),
                    FakeCell("男"),
                    FakeCell("出生日期"),
                    FakeCell("1995-07-15"),
                ]
            ),
            FakeRow(
                cells=[
                    FakeCell("送檢編號"),
                    FakeCell("APT-01-00XXXX"),
                    FakeCell("APT-01-00XXXX"),
                    FakeCell("APT-01-00XXXX"),
                    FakeCell("檢體類型"),
                    FakeCell("口腔黏膜"),
                ]
            ),
        ]
    )


def test_extract_name_from_table_with_spaced_label() -> None:
    converter = WordReportConverter()
    doc = FakeDocument(paragraphs=[], tables=[_build_realistic_table()])

    name, _ = converter._extract_identity(doc)

    assert name == "王曉明"


def test_extract_sample_id_from_table() -> None:
    converter = WordReportConverter()
    doc = FakeDocument(paragraphs=[], tables=[_build_realistic_table()])

    _, sample_id = converter._extract_identity(doc)

    assert sample_id == "APT-01-00XXXX"


def test_output_filename_format() -> None:
    converter = WordReportConverter()

    output = converter._build_output_filename(sample_id="APT-01-00XXXX", name="王曉明")

    assert output == "台-APT-01-00XXXX_王曉明-天賦30項.docx"
