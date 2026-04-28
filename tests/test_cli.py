from __future__ import annotations

from pathlib import Path

from src.word_converter.cli import _collect_input_files, _partition_pending_files


def test_collect_input_files_from_directory_skips_images(tmp_path: Path) -> None:
    docx_a = tmp_path / "a.docx"
    docx_b = tmp_path / "b.DOCX"
    image = tmp_path / "cover.png"
    note = tmp_path / "note.txt"
    docx_a.write_text("a")
    docx_b.write_text("b")
    image.write_text("img")
    note.write_text("note")

    input_files, skipped_images = _collect_input_files(tmp_path)

    assert input_files == [docx_a, docx_b]
    assert skipped_images == [image]


def test_collect_input_files_with_single_image_path_returns_skip(tmp_path: Path) -> None:
    image = tmp_path / "photo.jpg"
    image.write_text("img")

    input_files, skipped_images = _collect_input_files(image)

    assert input_files == []
    assert skipped_images == [image]


class _FakeConverter:
    def __init__(self, mapping: dict[Path, Path]) -> None:
        self.mapping = mapping

    def preview_output_path(self, input_path: Path, output_dir: Path) -> Path:
        return self.mapping[input_path]


def test_partition_pending_files_separates_already_processed(tmp_path: Path) -> None:
    input_a = tmp_path / "a.docx"
    input_b = tmp_path / "b.docx"
    input_a.write_text("a")
    input_b.write_text("b")
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    output_a = output_dir / "台-APT-01_A-天賦30項.docx"
    output_b = output_dir / "台-APT-02_B-天賦30項.docx"
    output_a.write_text("done")

    converter = _FakeConverter({input_a: output_a, input_b: output_b})

    pending_files, already_processed = _partition_pending_files([input_a, input_b], output_dir, converter)  # type: ignore[arg-type]

    assert pending_files == [input_b]
    assert already_processed == [(input_a, output_a)]
