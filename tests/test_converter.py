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


def _build_main_table() -> FakeTable:
    return FakeTable(
        rows=[
            FakeRow(
                cells=[
                    FakeCell("編 號"),
                    FakeCell("功 能"),
                    FakeCell("細胞解碼位點"),
                    FakeCell("解碼型"),
                    FakeCell("健康優勢評估"),
                    FakeCell("健康優勢評分"),
                ]
            ),
            FakeRow(cells=[FakeCell("1"), FakeCell("運動神經"), FakeCell("CNTF"), FakeCell("AA"), FakeCell("-"), FakeCell("90")]),
            FakeRow(cells=[FakeCell("2"), FakeCell("反應速度"), FakeCell("CNTF"), FakeCell("AB"), FakeCell("-"), FakeCell("91")]),
            FakeRow(cells=[FakeCell("3"), FakeCell("專注穩定性"), FakeCell("HTR2C"), FakeCell("BB"), FakeCell("-"), FakeCell("92")]),
            FakeRow(cells=[FakeCell("4"), FakeCell("專注力"), FakeCell("HTR2C"), FakeCell("BC"), FakeCell("-"), FakeCell("93")]),
            FakeRow(cells=[FakeCell("5"), FakeCell("協調性"), FakeCell("α-actinin"), FakeCell("CC"), FakeCell("-"), FakeCell("94")]),
            FakeRow(cells=[FakeCell("6"), FakeCell("肢體靈活性"), FakeCell("α-actinin"), FakeCell("CD"), FakeCell("-"), FakeCell("95")]),
        ]
    )


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


def test_convert_cell_codes_only_on_main_table() -> None:
    converter = WordReportConverter()
    main_table = _build_main_table()
    non_main_table = FakeTable(
        rows=[
            FakeRow(cells=[FakeCell("說明"), FakeCell("CNTF")]),
            FakeRow(cells=[FakeCell("備註"), FakeCell("HTR2C")]),
        ]
    )
    doc = FakeDocument(paragraphs=[FakeParagraph("段落中 CNTF 不應被替換")], tables=[non_main_table, main_table])

    converter._convert_cell_codes(doc)

    assert main_table.rows[1].cells[2].text == "MN001"
    assert main_table.rows[2].cells[2].text == "RTS001"
    assert non_main_table.rows[0].cells[1].text == "CNTF"
    assert doc.paragraphs[0].text == "段落中 CNTF 不應被替換"
    assert converter.last_cell_code_report["table_count"] == 2
    assert converter.last_cell_code_report["main_table_index"] == 1


def test_convert_cell_codes_distinguishes_duplicate_legacy_codes() -> None:
    converter = WordReportConverter()
    main_table = _build_main_table()
    doc = FakeDocument(paragraphs=[], tables=[main_table])

    converter._convert_cell_codes(doc)

    assert main_table.rows[1].cells[2].text == "MN001"  # 運動神經 + CNTF
    assert main_table.rows[2].cells[2].text == "RTS001"  # 反應速度 + CNTF
    assert main_table.rows[3].cells[2].text == "SF001"  # 專注穩定性 + HTR2C
    assert main_table.rows[4].cells[2].text == "CCT001"  # 專注力 + HTR2C
    assert main_table.rows[5].cells[2].text == "CD001"  # 協調性 + α-actinin
    assert main_table.rows[6].cells[2].text == "PFB001"  # 肢體靈活性 + α-actinin
    assert converter.last_cell_code_report["replaced_count"] == 6
    assert converter.last_cell_code_report["unmapped_features"] == []


def test_convert_cell_codes_collects_unmapped_features() -> None:
    converter = WordReportConverter()
    main_table = _build_main_table()
    main_table.rows.append(
        FakeRow(cells=[FakeCell("7"), FakeCell("不存在功能"), FakeCell("CNTF"), FakeCell("ZZ"), FakeCell("-"), FakeCell("80")])
    )
    doc = FakeDocument(paragraphs=[], tables=[main_table])

    converter._convert_cell_codes(doc)

    assert converter.last_cell_code_report["main_table_index"] == 0
    assert converter.last_cell_code_report["replaced_count"] == 6
    assert converter.last_cell_code_report["unmapped_features"] == ["不存在功能"]


def test_apply_fixed_text_replaces_declaration_and_score_item_texts() -> None:
    converter = WordReportConverter()
    doc = FakeDocument(
        paragraphs=[
            FakeParagraph("本報告僅供參考"),
            FakeParagraph("高分項目代表先天優勢"),
            FakeParagraph("低分項目代表先天不足"),
        ],
        tables=[],
    )

    converter._apply_fixed_text(doc)

    assert doc.paragraphs[0].text == "本報告為天賦 30 項分析結果，僅供健康管理參考。"
    assert doc.paragraphs[1].text == "高分項目代表相對優勢，建議持續強化並轉化為日常表現。"
    assert doc.paragraphs[2].text == "低分項目代表目前較需補強，建議透過訓練與習慣養成逐步改善。"


def test_apply_fixed_text_replaces_text_inside_table_cells() -> None:
    converter = WordReportConverter()
    table = FakeTable(
        rows=[
            FakeRow(cells=[FakeCell("說明"), FakeCell("如有疑問請洽客服")]),
            FakeRow(cells=[FakeCell("提醒"), FakeCell("高分項目代表先天優勢")]),
        ]
    )
    doc = FakeDocument(paragraphs=[], tables=[table])

    converter._apply_fixed_text(doc)

    assert table.rows[0].cells[1].text == "如需進一步解讀，請聯繫專屬顧問或客服中心。"
    assert table.rows[1].cells[1].text == "高分項目代表相對優勢，建議持續強化並轉化為日常表現。"
