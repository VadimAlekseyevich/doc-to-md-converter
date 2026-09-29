from __future__ import annotations

from io import BytesIO
from pathlib import Path
import struct
import zlib
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches

from doc_to_md_converter import ConversionError, convert_docx
from doc_to_md_converter.cli import main


def png(rgb: tuple[int, int, int] = (255, 0, 0)) -> bytes:
    def chunk(name: bytes, data: bytes) -> bytes:
        content = name + data
        return struct.pack(">I", len(data)) + content + struct.pack(">I", zlib.crc32(content))
    raw = b"\x00" + bytes(rgb)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">2I5B", 1, 1, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def picture(paragraph, data: bytes) -> None:
    paragraph.add_run().add_picture(BytesIO(data), width=Inches(0.2))


def hyperlink(paragraph, text: str, url: str) -> None:
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    rid = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    node = OxmlElement("w:hyperlink")
    node.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    value = OxmlElement("w:t")
    value.text = text
    run.append(value)
    node.append(run)
    paragraph._p.append(node)


def test_ordered_content_and_images_are_byte_identical(tmp_path: Path) -> None:
    doc = Document()
    doc.add_heading("Методичка", level=1)
    paragraph = doc.add_paragraph("До таблицы: ")
    paragraph.add_run("полужирный").bold = True
    paragraph.add_run(" и ")
    paragraph.add_run("курсив").italic = True
    hyperlink(paragraph, "документ", "https://example.org/some page")
    doc.add_paragraph("Пункт", style="List Bullet")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Слева"
    table.cell(0, 1).text = "Справа"
    data = png()
    picture(table.cell(1, 0).paragraphs[0], data)
    table.cell(1, 1).text = "Значение | с разделителем"
    doc.add_paragraph("После таблицы")
    picture(doc.add_paragraph("Вот изображение: "), data)
    source = tmp_path / "study.docx"
    doc.save(source)

    logs: list[str] = []
    result = convert_docx(source, progress=logs.append)
    markdown = result.markdown_path.read_text(encoding="utf-8")
    assert markdown.index("# Методичка") < markdown.index("До таблицы") < markdown.index("| Слева") < markdown.index("После таблицы")
    assert "**полужирный**" in markdown
    assert "*курсив*" in markdown
    assert "[документ](<https://example.org/some%20page>)" in markdown
    assert "- Пункт" in markdown
    assert "Значение \\| с разделителем" in markdown
    assert markdown.count("![") == 2
    assert markdown.count("assets/images/image1.png") == 2
    assert result.images_count == 1  # same embedded bytes/relation reused
    assert (result.output_dir / "assets/images/image1.png").read_bytes() == data
    assert "Готово" in result.log_path.read_text(encoding="utf-8")
    assert any("Сохранён ресурс" in line for line in logs)


def test_multiple_images_header_and_no_overwrite(tmp_path: Path) -> None:
    doc = Document()
    red, blue = png(), png((0, 0, 255))
    picture(doc.add_paragraph(), red)
    picture(doc.sections[0].header.paragraphs[0], blue)
    source = tmp_path / "lesson.docx"
    doc.save(source)
    first = convert_docx(source)
    second = convert_docx(source)
    assert first.output_dir.name == "lesson_md"
    assert second.output_dir.name == "lesson_md_2"
    assert first.images_count == 2
    markdown = first.markdown_path.read_text(encoding="utf-8")
    assert "Колонтитулы" in markdown
    assert "Верхний колонтитул" in markdown
    assert "assets/images/image1.png" in markdown
    assert "assets/images/image2.png" in markdown
    assert {path.read_bytes() for path in (first.output_dir / "assets/images").iterdir()} == {red, blue}
    with pytest.raises(ConversionError, match="уже существует"):
        convert_docx(source, output_dir=first.output_dir)


def test_orphaned_media_is_saved_and_logged(tmp_path: Path) -> None:
    doc = Document()
    doc.add_paragraph("Текст методички")
    source = tmp_path / "orphan.docx"
    doc.save(source)
    extra = png((12, 34, 56))
    with ZipFile(source, "a") as archive:
        archive.writestr("word/media/unreferenced.png", extra)
    result = convert_docx(source)
    assert result.images_count == 1
    assert (result.output_dir / "assets/images/image1.png").read_bytes() == extra
    assert any("не встречен" in warning for warning in result.warnings)
    text = result.markdown_path.read_text(encoding="utf-8")
    assert "Текст методички" in text
    assert "Дополнительные изображения" in text
    assert "assets/images/image1.png" in text


def test_unsupported_drawing_has_visible_placeholder_and_warning(tmp_path: Path) -> None:
    doc = Document()
    node = OxmlElement("w:drawing")
    doc.add_paragraph().add_run()._r.append(node)
    source = tmp_path / "shape.docx"
    doc.save(source)
    result = convert_docx(source)
    assert "Неподдерживаемый графический объект" in result.markdown_path.read_text(encoding="utf-8")
    assert any("без извлекаемой картинки" in w for w in result.warnings)


def test_reject_invalid_inputs_without_creating_export(tmp_path: Path) -> None:
    bad = tmp_path / "text.docx"
    bad.write_text("not a zip", encoding="utf-8")
    with pytest.raises(ConversionError, match="Некорректный"):
        convert_docx(bad)
    assert not (tmp_path / "text_md").exists()
    with pytest.raises(ConversionError, match="расширением"):
        convert_docx(tmp_path / "wrong.txt")
    with pytest.raises(ConversionError, match="не найден"):
        convert_docx(tmp_path / "missing.docx")


def test_cli_writes_output_and_reports_error(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    source = tmp_path / "from-cli.docx"
    doc = Document()
    doc.add_paragraph("CLI готов")
    doc.save(source)
    assert main([str(source)]) == 0
    assert "Готово:" in capsys.readouterr().out
    assert (tmp_path / "from-cli_md/index.md").exists()
    assert main([str(tmp_path / "missing.docx")]) == 1
    assert "Ошибка:" in capsys.readouterr().err


def test_vml_picture_is_extracted_from_legacy_markup(tmp_path: Path) -> None:
    doc = Document()
    image_data = png((9, 8, 7))
    ordinary = doc.add_paragraph()
    picture(ordinary, image_data)
    # Reuse the image rId in a legacy w:pict element.
    blip = ordinary._p.xpath(".//a:blip")[0]
    rid = blip.get(qn("r:embed"))
    legacy = doc.add_paragraph()
    pict = OxmlElement("w:pict")
    from lxml import etree
    v_image = etree.Element("{urn:schemas-microsoft-com:vml}imagedata")
    v_image.set(qn("r:id"), rid)
    pict.append(v_image)
    legacy.add_run()._r.append(pict)
    source = tmp_path / "legacy.docx"
    doc.save(source)

    result = convert_docx(source)
    assert result.images_count == 1
    assert result.markdown_path.read_text(encoding="utf-8").count("assets/images/image1.png") == 2
    assert (result.output_dir / "assets/images/image1.png").read_bytes() == image_data


def test_numbered_list_and_escaping(tmp_path: Path) -> None:
    doc = Document()
    doc.add_paragraph("First", style="List Number")
    doc.add_paragraph("Second", style="List Number")
    doc.add_paragraph("# not a heading")
    doc.add_paragraph("Some [brackets] *star* and |pipe|")
    source = tmp_path / "list.docx"
    doc.save(source)
    result = convert_docx(source)
    text = result.markdown_path.read_text(encoding="utf-8")
    assert "1. First" in text and "1. Second" in text
    assert "\\# not a heading" in text
    assert r"Some \[brackets\] \*star\* and \|pipe\|" in text


def test_custom_output_parent_autonames_without_overwriting(tmp_path: Path) -> None:
    doc = Document()
    doc.add_paragraph("Отдельное расположение")
    original = tmp_path / "source" / "lecture.docx"
    original.parent.mkdir()
    doc.save(original)
    destination = tmp_path / "exports with space"
    destination.mkdir()
    result1 = convert_docx(original, output_parent=destination)
    result2 = convert_docx(original, output_parent=destination)
    assert result1.output_dir == destination / "lecture_md"
    assert result2.output_dir == destination / "lecture_md_2"
    assert original.read_bytes()  # Original is still there.
    assert "Отдельное расположение" in result1.markdown_path.read_text(encoding="utf-8")
    with pytest.raises(ConversionError, match="Папка для сохранения не найдена"):
        convert_docx(original, output_parent=tmp_path / "missing")
    with pytest.raises(ConversionError, match="одновременно"):
        convert_docx(original, output_dir=tmp_path / "exact", output_parent=destination)
    assert main([str(original), "--output-parent", str(destination)]) == 0
    assert (destination / "lecture_md_3/index.md").is_file()


def test_footnotes_comments_equations_and_embedded_objects(tmp_path: Path) -> None:
    from doc_to_md_converter.converter import M, W
    from lxml import etree

    doc = Document()
    paragraph = doc.add_paragraph("Основной текст ")
    note = OxmlElement("w:footnoteReference")
    note.set(qn("w:id"), "1")
    paragraph.add_run()._r.append(note)
    comment = OxmlElement("w:commentReference")
    comment.set(qn("w:id"), "2")
    paragraph.add_run()._r.append(comment)
    math = etree.Element(f"{{{M}}}oMath", nsmap={"m": M})
    run = etree.SubElement(math, f"{{{M}}}r")
    value = etree.SubElement(run, f"{{{M}}}t")
    value.text = "x+2"
    paragraph._p.append(math)
    source = tmp_path / "rich.docx"
    doc.save(source)
    footnotes = (f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1">'
                 '<w:p><w:r><w:t>Содержимое сноски</w:t></w:r></w:p>'
                 '</w:footnote></w:footnotes>')
    comments = (f'<w:comments xmlns:w="{W}"><w:comment w:id="2" w:author="Редактор">'
                '<w:p><w:r><w:t>Содержимое комментария</w:t></w:r></w:p>'
                '</w:comment></w:comments>')
    object_data = b"original ole bytes"
    with ZipFile(source, "a") as archive:
        archive.writestr("word/footnotes.xml", footnotes)
        archive.writestr("word/comments.xml", comments)
        archive.writestr("word/embeddings/oleObject1.bin", object_data)
    result = convert_docx(source)
    md = result.markdown_path.read_text(encoding="utf-8")
    assert "Основной текст" in md
    assert "[Сноска 1]" in md and "Содержимое сноски" in md
    assert "[Комментарий 2]" in md and "Содержимое комментария" in md
    assert "Формула Word: x+2" in md
    assert "assets/objects/object1.bin" in md
    assert (result.output_dir / "assets/objects/object1.bin").read_bytes() == object_data
    assert any("формулы Word" in warning for warning in result.warnings)
    assert any("OLE" in warning for warning in result.warnings)


def test_textbox_text_is_preserved_despite_unsupported_shape(tmp_path: Path) -> None:
    from lxml import etree
    from doc_to_md_converter.converter import W
    doc = Document()
    paragraph = doc.add_paragraph("До блока. ")
    drawing = OxmlElement("w:drawing")
    textbox = etree.SubElement(drawing, f"{{{W}}}txbxContent")
    p = etree.SubElement(textbox, f"{{{W}}}p")
    r = etree.SubElement(p, f"{{{W}}}r")
    t = etree.SubElement(r, f"{{{W}}}t")
    t.text = "Важная надпись в схеме"
    paragraph.add_run()._r.append(drawing)
    source = tmp_path / "textbox.docx"
    doc.save(source)
    result = convert_docx(source)
    md = result.markdown_path.read_text(encoding="utf-8")
    assert "Важная надпись в схеме" in md
    assert "Неподдерживаемый графический объект" in md
    assert any("графического блока" in warning for warning in result.warnings)


def test_package_absolute_image_relationship_is_supported(tmp_path: Path) -> None:
    doc = Document()
    image_data = png((17, 33, 99))
    picture(doc.add_paragraph(), image_data)
    source = tmp_path / "absolute-target.docx"
    doc.save(source)
    with ZipFile(source) as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    rels = "word/_rels/document.xml.rels"
    assert b'Target="media/image1.png"' in members[rels]
    members[rels] = members[rels].replace(b'Target="media/image1.png"', b'Target="/word/media/image1.png"')
    with ZipFile(source, "w") as archive:
        for name, data in members.items():
            archive.writestr(name, data)
    result = convert_docx(source)
    assert result.images_count == 1
    assert (result.output_dir / "assets/images/image1.png").read_bytes() == image_data
    assert "assets/images/image1.png" in result.markdown_path.read_text(encoding="utf-8")


def test_private_windows_image_path_is_not_exported_as_alt_text(tmp_path: Path) -> None:
    doc = Document()
    image_data = png((31, 42, 53))
    paragraph = doc.add_paragraph()
    picture(paragraph, image_data)
    picture_properties = paragraph._p.xpath(".//wp:docPr")[0]
    picture_properties.set("descr", r"C:\Users\lizaa\Desktop\radick\4 kurs\диаграммы\UseCase_AS_IS.jpg")
    source = tmp_path / "figure.docx"
    doc.save(source)
    result = convert_docx(source)
    markdown = result.markdown_path.read_text(encoding="utf-8")
    assert r"C:\Users" not in markdown
    assert "radick" not in markdown
    assert r"![UseCase\_AS\_IS.jpg](assets/images/image1.png)" in markdown
    assert (result.output_dir / "assets/images/image1.png").read_bytes() == image_data


@pytest.mark.parametrize(
    ("description", "expected"),
    [
        ("Схема работы", "Схема работы"),
        (r"C:\Users\Public\Pictures\diagram.png", "diagram.png"),
        ("/home/user/pictures/chart.jpeg", "chart.jpeg"),
        (r"..\images\photo.png", "photo.png"),
        ("file:///C:/Users/Public/pic.jpg", "pic.jpg"),
        ("", "Изображение"),
    ],
)
def test_image_description_removes_paths_but_keeps_real_captions(description: str, expected: str) -> None:
    from doc_to_md_converter.converter import short_image_alt
    assert short_image_alt(description) == expected


def test_explorer_conversion_returns_feedback_without_keeping_gui_open(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    from doc_to_md_converter import cli

    messages: list[tuple[str, str, bool]] = []
    monkeypatch.setattr(cli, "_notify_context_result",
                        lambda title, message, *, error=False: messages.append((title, message, error)))
    doc = Document()
    doc.add_paragraph("Запуск из меню")
    source = tmp_path / "отчёт.docx"
    doc.save(source)
    assert main([str(source), "--context-menu"]) == 0
    assert (tmp_path / "отчёт_md" / "index.md").is_file()
    assert len(messages) == 1 and messages[0][2] is False
    assert str(tmp_path / "отчёт_md") in messages[0][1]
    messages.clear()
    assert main([str(tmp_path / "missing.docx"), "--context-menu"]) == 1
    assert len(messages) == 1 and messages[0][2] is True
