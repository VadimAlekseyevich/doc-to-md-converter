"""PDF to GitHub-friendly Markdown; no OCR or raster-to-text inference.

The PDF format has no reliable semantic headings/table structure. We keep page
boundaries, extract its selectable text and visible image objects in rough
reading order, and save page previews for vector drawing content.
"""
from __future__ import annotations

from hashlib import sha256
from io import BytesIO
from pathlib import Path
import shutil
import tempfile
from typing import Callable

import pypdfium2 as pdfium
from PIL import Image

from .converter import ConversionError, ConversionResult, escape


_PDF_SUFFIXES = {"JPEG": ".jpeg", "JPEG2000": ".jp2", "PNG": ".png", "TIFF": ".tiff",
                 "GIF": ".gif", "BMP": ".bmp", "WEBP": ".webp"}


def _image_extension(data: bytes) -> str:
    """Match filenames to actual extracted payload, not the source PDF name."""
    with Image.open(BytesIO(data)) as image:
        return _PDF_SUFFIXES.get(image.format or "", ".png")


def _text_fragments(obj: pdfium.PdfTextObj) -> str:
    return obj.extract().replace("\r\n", "\n").strip()


def _page_markdown(objects: list[tuple[float, float, float, str, str]]) -> str:
    """Sort page objects roughly top-to-bottom, left-to-right; join same-line text."""
    output: list[str] = []
    line: list[str] = []
    line_top: float | None = None
    line_right: float | None = None

    def flush() -> None:
        nonlocal line_top, line_right
        if line:
            output.append("".join(line).strip())
            line.clear()
        line_top = None
        line_right = None

    for top, left, right, kind, value in sorted(objects, key=lambda x: (-x[0], x[1])):
        if kind == "image":
            flush()
            output.append(value)
            continue
        if line_top is not None and abs(line_top - top) > 6:
            flush()
        if not line:
            line_top = top
            line.append(value)
        else:
            separator = (" " if line_right is not None and left - line_right > 1.5
                         and not line[-1].endswith((" ", "-", "\n"))
                         and not value.startswith((",", ".", ":", ";", "!", "?", ")")) else "")
            line.append(separator + value)
        line_right = right
    flush()
    return "\n\n".join(v for v in output if v)


def convert_pdf(
    source: str | Path,
    output_dir: str | Path | None = None,
    progress: Callable[[str], None] | None = None,
    *,
    output_parent: str | Path | None = None,
) -> ConversionResult:
    """Export PDF text and image objects without OCR, never overwriting existing output.

    JPEG / JPEG2000 image payloads are copied where PDFium allows. Other PDF
    image encodings may need a PNG fallback; vector drawings receive page previews.
    """
    original = Path(source).expanduser().resolve()
    if original.suffix.lower() != ".pdf":
        raise ConversionError("Выберите файл с расширением .pdf")
    if not original.is_file():
        raise ConversionError(f"Файл не найден: {original}")
    if output_dir is not None and output_parent is not None:
        raise ConversionError("Нельзя одновременно указывать output_dir и output_parent")
    if output_dir is None:
        parent = Path(output_parent).expanduser().resolve() if output_parent is not None else original.parent
        if not parent.is_dir():
            raise ConversionError(f"Папка для сохранения не найдена: {parent}")
        base = parent / (original.stem + "_md")
        target = base
        number = 2
        while target.exists():
            target = parent / (base.name + "_" + str(number))
            number += 1
    else:
        target = Path(output_dir).expanduser().resolve()
        if target.exists():
            raise ConversionError(f"Папка результата уже существует: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    temp = Path(tempfile.mkdtemp(prefix=".pdf-to-md-", dir=target.parent))
    logs: list[str] = []
    warnings: list[str] = []

    def log(level: str, message: str) -> None:
        line = f"[{level}] {message}"
        logs.append(line)
        if level == "WARN":
            warnings.append(message)
        if progress:
            progress(line)

    try:
        try:
            document = pdfium.PdfDocument(str(original))
        except (ValueError, RuntimeError, OSError) as exc:
            raise ConversionError(f"Некорректный, защищённый или неподдерживаемый PDF: {exc}") from exc
        try:
            images_dir = temp / "assets" / "images"
            images_dir.mkdir(parents=True, exist_ok=True)
            stored: dict[str, str] = {}
            pages_md: list[str] = []
            log("INFO", f"Чтение PDF: {len(document)} стр.")
            log("WARN", "PDF не содержит семантической структуры DOCX: порядок колонок, таблиц и стили могут отличаться.")
            for page_index in range(len(document)):
                page = document[page_index]
                objects: list[tuple[float, float, float, str, str]] = []
                text_count = 0
                image_count = 0
                vectors = False
                textpage = page.get_textpage()
                try:
                    for obj in page.get_objects(textpage=textpage):
                        try:
                            left, bottom, right, top = obj.get_bounds()
                            if isinstance(obj, pdfium.PdfTextObj):
                                raw = _text_fragments(obj)
                                if raw:
                                    text_count += 1
                                    objects.append((top, left, right, "text", escape(raw)))
                            elif isinstance(obj, pdfium.PdfImage):
                                content = BytesIO()
                                try:
                                    obj.extract(content, fb_format="png")
                                    data = content.getvalue()
                                    extension = _image_extension(data)
                                except (OSError, ValueError, RuntimeError) as exc:
                                    log("WARN", f"Страница {page_index + 1}: не удалось извлечь одно из изображений: {exc}")
                                    vectors = True
                                    continue
                                digest = sha256(data).hexdigest()
                                if digest not in stored:
                                    rel = f"assets/images/image{len(stored) + 1}{extension}"
                                    (temp / rel).write_bytes(data)
                                    stored[digest] = rel
                                    log("INFO", f"Изображение страницы {page_index + 1}: {rel}")
                                else:
                                    rel = stored[digest]
                                image_count += 1
                                objects.append((top, left, right, "image", f"![Изображение {image_count}]({rel})"))
                            else:
                                vectors = True
                        except (ValueError, RuntimeError) as exc:
                            log("WARN", f"Страница {page_index + 1}: объект PDF пропущен: {exc}")
                            vectors = True
                finally:
                    textpage.close()
                body = _page_markdown(objects)
                if not text_count:
                    log("WARN", f"Страница {page_index + 1}: нет извлекаемого текстового слоя; OCR не выполняется.")
                if vectors or not objects:
                    # Rasterize only when no stand-alone bitmap represents visual content.
                    # This preserves vector charts/lines/tables without claiming OCR.
                    previews = temp / "assets" / "pages"
                    previews.mkdir(parents=True, exist_ok=True)
                    relative = f"assets/pages/page{page_index + 1}.png"
                    bitmap = page.render(scale=1.5)
                    try:
                        bitmap.to_pil().save(temp / relative, format="PNG", optimize=True)
                    finally:
                        bitmap.close()
                    body += ("\n\n" if body else "") + f"[Визуальная копия страницы {page_index + 1}]({relative})"
                    log("INFO", f"Страница {page_index + 1}: сохранена визуальная копия {relative}")
                if not body:
                    body = "<!-- На странице нет извлекаемого текста или изображений -->"
                pages_md.append(f"## Страница {page_index + 1}\n\n{body}")
                page.close()
            (temp / "index.md").write_text("# " + escape(original.stem) + "\n\n" + "\n\n".join(pages_md) + "\n", encoding="utf-8")
            log("INFO", f"Изображений: {len(stored)}; предупреждений: {len(warnings)}")
            log("INFO", "Готово: index.md, assets/, conversion.log")
            (temp / "conversion.log").write_text("\n".join(logs) + "\n", encoding="utf-8")
            if target.exists():
                raise ConversionError(f"Папка результата уже существует: {target}")
            temp.rename(target)
            return ConversionResult(target, target / "index.md", target / "conversion.log", len(stored), tuple(warnings))
        except ConversionError:
            raise
        except Exception as exc:
            raise ConversionError(f"Не удалось обработать PDF: {exc}") from exc
        finally:
            document.close()
    finally:
        if temp.exists():
            shutil.rmtree(temp)
