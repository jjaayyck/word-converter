from word_converter.converter import WordReportConverter


class DummyCell:
    def __init__(self, text: str):
        self.text = text


class DummyRow:
    def __init__(self, cells: list[str]):
        self.cells = [DummyCell(text) for text in cells]


class DummyTable:
    def __init__(self, rows: list[list[str]]):
        self.rows = [DummyRow(row) for row in rows]


class DummyParagraph:
    def __init__(self, text: str):
        self.text = text


class DummyDocument:
    def __init__(self, paragraphs: list[str], tables: list[DummyTable] | None = None):
        self.paragraphs = [DummyParagraph(text) for text in paragraphs]
        self.tables = tables or []


def test_extract_identity_from_document_paragraphs() -> None:
    converter = WordReportConverter()
    document = DummyDocument([
        "這是報告",
        "姓名：王小明",
        "送檢編號：APT-01-009297",
    ])

    name, sample_id = converter._extract_identity(document)  # type: ignore[arg-type]

    assert name == "王小明"
    assert sample_id == "APT-01-009297"


def test_extract_by_labels_without_colon() -> None:
    converter = WordReportConverter()
    text = "姓名 王小美\n送檢編號 APT-01-000001"

    assert converter._extract_by_labels(text, ["姓名"]) == "王小美"
    assert converter._extract_by_labels(text, ["送檢編號"]) == "APT-01-000001"


def test_extract_identity_from_table_embedded_label_value() -> None:
    converter = WordReportConverter()
    table = DummyTable(
        [
            ["報告資訊", "其他"],
            ["姓名：林大同", "送檢編號：TW-00123"],
        ]
    )

    name, sample_id = converter._extract_identity_from_table(table)  # type: ignore[arg-type]

    assert name == "林大同"
    assert sample_id == "TW-00123"


def test_output_filename_format_and_sanitization() -> None:
    converter = WordReportConverter()

    filename = converter._build_output_filename(sample_id="APT-01/009297", name="王 小*明")

    assert filename == "台-APT-01009297_王小明-天賦30項.docx"


def test_extract_by_labels_stops_at_punctuation() -> None:
    converter = WordReportConverter()
    text = "姓名：陳小華；送檢編號：TW001"

    assert converter._extract_by_labels(text, ["姓名"]) == "陳小華"
