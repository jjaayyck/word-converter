from __future__ import annotations

from dataclasses import dataclass, field

from src.word_converter.converter import WordReportConverter


@dataclass
class FakeCell:
    text: str


@dataclass
class FakeRow:
    cells: list[FakeCell]
    height: int | None = None
    height_rule: int | None = None


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
    sections: list[object] = field(default_factory=list)


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


def test_convert_table_headers_replaces_main_table_titles() -> None:
    converter = WordReportConverter()
    main_table = _build_main_table()
    doc = FakeDocument(paragraphs=[], tables=[main_table])

    converter._convert_table_headers(doc)

    headers = [cell.text for cell in main_table.rows[0].cells]
    assert headers == ["編 號", "心理天賦項目", "細胞解碼位點", "解碼型", "心理潛能優勢評估", "心理潛能優勢評分"]




def test_convert_table_headers_keeps_name_label_unchanged() -> None:
    converter = WordReportConverter()
    table = FakeTable(rows=[FakeRow(cells=[FakeCell("姓名"), FakeCell("送檢編號")])])
    doc = FakeDocument(paragraphs=[], tables=[table])

    converter._convert_table_headers(doc)

    assert table.rows[0].cells[0].text == "姓名"

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


def test_convert_cell_codes_still_works_after_header_replacement() -> None:
    converter = WordReportConverter()
    main_table = _build_main_table()
    doc = FakeDocument(paragraphs=[], tables=[main_table])

    converter._convert_table_headers(doc)
    converter._convert_cell_codes(doc)

    assert main_table.rows[1].cells[2].text == "MN001"
    assert converter.last_cell_code_report["main_table_index"] == 0


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


def test_replace_recommendation_section_updates_name_and_templates() -> None:
    converter = WordReportConverter()
    main_table = _build_main_table()
    main_table.rows[1].cells[1].text = "高特質"
    main_table.rows[1].cells[4].text = "高"
    main_table.rows[2].cells[1].text = "低特質"
    main_table.rows[2].cells[4].text = "低"
    doc = FakeDocument(
        paragraphs=[
            FakeParagraph("本報告所提供之心理天賦優勢分析"),
            FakeParagraph("吳峻維"),
            FakeParagraph("舊版健康管理文案"),
        ],
        tables=[main_table],
    )

    converter._replace_recommendation_section(doc, "王曉明")

    all_text = "\n".join(p.text for p in doc.paragraphs)
    assert "吳峻維" not in all_text
    assert "心理潛能亮點建議" in all_text
    assert "Guidance to Discover Your Hidden Strengths" in all_text
    assert "王曉明" in all_text
    assert "低分項目（低特質）" in all_text


def test_apply_fixed_text_replaces_long_declaration_text() -> None:
    converter = WordReportConverter()
    old_text = (
        "o\t本報告依細胞分子生物學分析及統計資料，以口腔黏膜樣本進行檢測，僅供本次健康管理參考，"
        "無臨床診斷效力，亦不可作為醫療診斷依據。如有健康疑慮，請諮詢專業醫師。"
    )
    doc = FakeDocument(paragraphs=[FakeParagraph(old_text)], tables=[])

    converter._apply_fixed_text(doc)

    assert "本報告所提供之心理天賦優勢分析" in doc.paragraphs[0].text
    assert doc.paragraphs[0].text.startswith("o\t本報告所提供之心理天賦優勢分析")


@dataclass
class FakeSection:
    top_margin: int | None = None
    left_margin: int | None = None
    right_margin: int | None = None
    bottom_margin: int | None = None


def test_convert_cell_codes_updates_gene_row_height() -> None:
    converter = WordReportConverter()
    main_table = _build_main_table()
    doc = FakeDocument(paragraphs=[], tables=[main_table])

    converter._convert_cell_codes(doc)

    assert main_table.rows[1].height == int(1.9 * 360000)
    assert main_table.rows[1].height_rule == 2




def test_apply_page_layout_skips_last_section_when_disclaimer_exists() -> None:
    converter = WordReportConverter()
    section_a = FakeSection()
    section_b = FakeSection()
    doc = FakeDocument(
        paragraphs=[FakeParagraph("本報告所提供之心理天賦優勢分析")],
        tables=[],
        sections=[section_a, section_b],
    )

    converter._apply_page_layout(doc)

    assert section_a.top_margin == int(0.75 * 360000)
    assert section_b.top_margin is None

def test_apply_page_layout_updates_margins() -> None:
    converter = WordReportConverter()
    section = FakeSection()
    doc = FakeDocument(paragraphs=[], tables=[], sections=[section])

    converter._apply_page_layout(doc)

    assert section.top_margin == int(0.75 * 360000)
    assert section.left_margin == int(1.0 * 360000)
    assert section.right_margin == int(1.0 * 360000)
    assert section.bottom_margin == int(1.0 * 360000)
