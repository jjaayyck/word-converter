from __future__ import annotations

from pathlib import Path

from src.word_converter.cli import _collect_input_files


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
