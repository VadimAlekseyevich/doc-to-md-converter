"""PDF regression fixtures are synthetic and do not require OCR or Word."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path

from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas
import pytest

from doc_to_md_converter import ConversionError, convert_document
from doc_to_md_converter.cli import main
from doc_to_md_converter.pdf_converter import convert_pdf


def jpeg(color: tuple[int, int, int] = (50, 90, 120)) -> bytes:
    buffer = BytesIO()
    Image.new("RGB", (12, 10), color).save(buffer, format="JPEG", quality=92)
    return buffer.getvalue()


def make_pdf(path: Path, jpg: bytes, *, pages: int = 2, vector: bool = True) -> None:
    canvas = Canvas(str(path), pagesize=(400, 500))
    for index in range(pages):
        if index == 0:
            canvas.drawString(35, 440, "PDF selectable text")
            canvas.drawString(35, 420, "Second line")
        canvas.drawImage(ImageReader(BytesIO(jpg)), 35, 280, width=72, height=60)
        if index == 0 and vector:
            canvas.setLineWidth(2)
            canvas.rect(30, 150, 180, 80)
        canvas.showPage()
    canvas.save()


def test_pdf_text_images_multi_page_without_duplicate_snapshot(tmp_path: Path) -> None:
    jpg = jpeg()
    path = tmp_path / "study material.pdf"
    make_pdf(path, jpg)
    result = convert_document(path)
    markdown = result.markdown_path.read_text(encoding="utf-8")
    assert "# study material" in markdown
    assert markdown.index("## Страница 1") < markdown.index("PDF selectable text") < markdown.index("## Страница 2")
    assert "Second line" in markdown
    assert result.images_count == 1  # A reused PDF image is exported once.
    assert markdown.count("(assets/images/image1.jpeg)") == 2
    assert (result.output_dir / "assets/images/image1.jpeg").read_bytes() == jpg
    assert "Визуальная копия" not in markdown
    assert not (result.output_dir / "assets/pages").exists()
    assert any("векторные элементы не переносятся" in message for message in result.warnings)
    assert any("OCR не выполняется" in message for message in result.warnings)
    assert "Готово" in result.log_path.read_text(encoding="utf-8")


def test_pdf_scanned_page_keeps_pixels_without_ocr(tmp_path: Path) -> None:
    image = Image.new("RGB", (20, 20), (125, 25, 70))
    path = tmp_path / "scan.pdf"
    canvas = Canvas(str(path), pagesize=(160, 200))
    canvas.drawImage(ImageReader(image), 0, 0, 160, 200)
    canvas.showPage()
    canvas.save()
    result = convert_pdf(path)
    markdown = result.markdown_path.read_text(encoding="utf-8")
    assert "![Изображение 1](assets/images/image1.png)" in markdown
    assert "OCR" not in markdown
    with Image.open(result.output_dir / "assets/images/image1.png") as extracted:
        assert extracted.size == image.size
        assert extracted.convert("RGB").getpixel((5, 5)) == (125, 25, 70)
    assert any("нет извлекаемого текстового слоя" in w for w in result.warnings)


def test_pdf_output_parent_auto_numbering_and_cli(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    source = tmp_path / "source.pdf"
    make_pdf(source, jpeg(), pages=1, vector=False)
    destination = tmp_path / "target dir"
    destination.mkdir()
    one = convert_pdf(source, output_parent=destination)
    two = convert_pdf(source, output_parent=destination)
    assert one.output_dir == destination / "source_md"
    assert two.output_dir == destination / "source_md_2"
    assert source.is_file()
    assert main([str(source), "--output-parent", str(destination)]) == 0
    assert "Готово:" in capsys.readouterr().out
    assert (destination / "source_md_3/index.md").is_file()
    with pytest.raises(ConversionError, match="уже существует"):
        convert_pdf(source, output_dir=one.output_dir)
    with pytest.raises(ConversionError, match="одновременно"):
        convert_pdf(source, output_dir=destination / "new", output_parent=destination)
    with pytest.raises(ConversionError, match="не найдена"):
        convert_pdf(source, output_parent=tmp_path / "missing")


def test_pdf_invalid_files_and_extension(tmp_path: Path) -> None:
    broken = tmp_path / "broken.pdf"
    broken.write_bytes(b"%PDF-1.4\ninvalid")
    with pytest.raises(ConversionError, match="PDF"):
        convert_document(broken)
    assert not (tmp_path / "broken_md").exists()
    with pytest.raises(ConversionError, match="не найден"):
        convert_document(tmp_path / "missing.pdf")
    with pytest.raises(ConversionError, match=".docx или .pdf"):
        convert_document(tmp_path / "other.txt")


def test_pdf_vector_only_page_retains_visual_snapshot(tmp_path: Path) -> None:
    path = tmp_path / "diagram.pdf"
    canvas = Canvas(str(path), pagesize=(300, 200))
    canvas.line(20, 20, 280, 180)
    canvas.showPage()
    canvas.save()
    result = convert_pdf(path)
    md = result.markdown_path.read_text(encoding="utf-8")
    assert "assets/pages/page1.png" in md
    assert (result.output_dir / "assets/pages/page1.png").stat().st_size > 0
    assert any("OCR не выполняется" in w for w in result.warnings)



def test_pdf_text_and_vector_only_does_not_duplicate_page(tmp_path: Path) -> None:
    path = tmp_path / "text-and-vector.pdf"
    canvas = Canvas(str(path), pagesize=(300, 200))
    canvas.drawString(20, 170, "Actual editable text")
    canvas.line(20, 20, 280, 120)
    canvas.showPage()
    canvas.save()
    result = convert_pdf(path)
    markdown = result.markdown_path.read_text(encoding="utf-8")
    assert "Actual editable text" in markdown
    assert "assets/pages/" not in markdown
    assert not (result.output_dir / "assets/pages").exists()
    assert any("векторные элементы не переносятся" in w for w in result.warnings)


def test_pdf_blank_page_creates_no_image_or_snapshot(tmp_path: Path) -> None:
    path = tmp_path / "blank.pdf"
    canvas = Canvas(str(path), pagesize=(300, 200))
    canvas.showPage()
    canvas.save()
    result = convert_pdf(path)
    markdown = result.markdown_path.read_text(encoding="utf-8")
    assert "Нет извлекаемого текста или изображений" in markdown
    assert "assets/pages/" not in markdown
    assert not (result.output_dir / "assets/pages").exists()
    assert result.images_count == 0
