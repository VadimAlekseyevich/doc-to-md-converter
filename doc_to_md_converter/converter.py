"""DOCX (OOXML) -> GitHub-readable Markdown with lossless embedded media extraction.

The XML walk is intentionally ordered: python-docx's paragraph and table lists alone
would lose the interleaving of paragraphs, tables, images and hyperlinks.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import posixpath
import re
import shutil
import tempfile
from typing import Callable
from urllib.parse import quote
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
V = "urn:schemas-microsoft-com:vml"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "r": R, "a": A, "v": V, "wp": WP, "m": M}
XML = etree.XMLParser(resolve_entities=False, no_network=True)
REL_IMAGE = "/image"


def q(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def attribute(node: etree._Element | None, name: str, ns: str = W) -> str | None:
    return node.get(q(ns, name)) if node is not None else None


def escape(text: str) -> str:
    """Escape Markdown metacharacters in normal paragraph and table content."""
    return re.sub(r"([\\`*_{}\[\]<>|])", r"\\\1", text).replace("\r", "")


def safe_url(url: str) -> str:
    # The angle-bracket destination syntax permits spaces when URL-encoded.
    return quote(url, safe="/:#?&=%+@;,$!~*'-._")


@dataclass(frozen=True)
class ConversionResult:
    output_dir: Path
    markdown_path: Path
    log_path: Path
    images_count: int
    warnings: tuple[str, ...]


class ConversionError(Exception):
    """User-visible input or conversion failure."""


class _Converter:
    def __init__(self, archive: ZipFile, output: Path, progress: Callable[[str], None] | None):
        self.archive = archive
        self.names = set(archive.namelist())
        self.output = output
        self.images_dir = output / "assets" / "images"
        self.images_dir.mkdir(parents=True, exist_ok=True)
        self.progress = progress
        self.lines: list[str] = []
        self.warnings: list[str] = []
        self.image_paths: dict[str, str] = {}  # ZIP member -> relative Markdown path
        self.rels_cache: dict[str, dict[str, tuple[str, str, str]]] = {}
        self.styles: dict[str, dict[str, object]] = {}
        self.numbering: dict[tuple[str, int], str] = {}
        self._reported: set[str] = set()
        self._load_styles()
        self._load_numbering()

    def log(self, level: str, message: str) -> None:
        line = f"[{level}] {message}"
        self.lines.append(line)
        if level == "WARN":
            self.warnings.append(message)
        if self.progress:
            self.progress(line)

    def warn_once(self, key: str, message: str) -> None:
        if key not in self._reported:
            self._reported.add(key)
            self.log("WARN", message)

    def xml(self, path: str) -> etree._Element:
        return etree.fromstring(self.archive.read(path), parser=XML)

    def _load_styles(self) -> None:
        if "word/styles.xml" not in self.names:
            return
        for style in self.xml("word/styles.xml").findall("w:style", NS):
            sid = attribute(style, "styleId")
            if not sid:
                continue
            name = style.find("w:name", NS)
            outline = style.find("w:pPr/w:outlineLvl", NS)
            num = style.find("w:pPr/w:numPr", NS)
            self.styles[sid] = {
                "name": (attribute(name, "val") or sid).lower(),
                "outline": attribute(outline, "val"),
                "num_id": attribute(num.find("w:numId", NS), "val") if num is not None else None,
                "level": attribute(num.find("w:ilvl", NS), "val") if num is not None else None,
            }

    def _load_numbering(self) -> None:
        if "word/numbering.xml" not in self.names:
            return
        root = self.xml("word/numbering.xml")
        abstract: dict[str, dict[int, str]] = {}
        for node in root.findall("w:abstractNum", NS):
            aid = attribute(node, "abstractNumId")
            if aid is None:
                continue
            levels = {}
            for lvl in node.findall("w:lvl", NS):
                index = int(attribute(lvl, "ilvl") or "0")
                fmt = lvl.find("w:numFmt", NS)
                levels[index] = attribute(fmt, "val") or "decimal"
            abstract[aid] = levels
        for num in root.findall("w:num", NS):
            nid = attribute(num, "numId")
            aid = attribute(num.find("w:abstractNumId", NS), "val")
            if nid is None or aid is None:
                continue
            for level, fmt in abstract.get(aid, {}).items():
                self.numbering[(nid, level)] = fmt
            for override in num.findall("w:lvlOverride", NS):
                level = int(attribute(override, "ilvl") or "0")
                fmt = override.find("w:lvl/w:numFmt", NS)
                if fmt is not None:
                    self.numbering[(nid, level)] = attribute(fmt, "val") or "decimal"

    def relationships(self, part: str) -> dict[str, tuple[str, str, str]]:
        if part in self.rels_cache:
            return self.rels_cache[part]
        rels_path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        relations: dict[str, tuple[str, str, str]] = {}
        if rels_path in self.names:
            for rel in self.xml(rels_path):
                if rel.tag == q(PR, "Relationship"):
                    relations[rel.get("Id", "")] = (rel.get("Type", ""), rel.get("Target", ""), rel.get("TargetMode", ""))
        self.rels_cache[part] = relations
        return relations

    def resolve(self, part: str, target: str) -> str | None:
        normalized = target.replace("\\", "/")
        # OOXML permits both relative targets and package-absolute /word/media/... URLs.
        member = (posixpath.normpath(normalized.lstrip("/")) if normalized.startswith("/") else
                  posixpath.normpath(posixpath.join(posixpath.dirname(part), normalized)))
        if member.startswith("../") or member.startswith("/") or member == "..":
            self.warn_once("unsafe-path", "Обнаружена небезопасная ссылка на ресурс внутри DOCX.")
            return None
        return member

    def store_image(self, member: str, referenced: bool = True) -> str | None:
        if member in self.image_paths:
            return self.image_paths[member]
        if member not in self.names:
            self.log("WARN", f"Встроенный ресурс не найден в DOCX: {member}")
            return None
        extension = Path(member).suffix.lower()
        # Keep original bytes and original extension, even if GitHub can't preview the format.
        if not re.fullmatch(r"\.[a-z0-9]{1,10}", extension):
            extension = ".bin"
            self.log("WARN", f"Неизвестное расширение медиафайла: {member}; сохранено как .bin")
        name = f"image{len(self.image_paths) + 1}{extension}"
        relative = f"assets/images/{name}"
        (self.images_dir / name).write_bytes(self.archive.read(member))
        self.image_paths[member] = relative
        self.log("INFO", f"Сохранён ресурс {member} → {relative}")
        if extension in {".emf", ".wmf", ".bin", ".tif", ".tiff"}:
            self.log("WARN", f"GitHub может не отображать изображение {relative}; оригинальные байты сохранены.")
        if not referenced:
            self.log("WARN", f"Медиафайл не встречен при обходе текста, но сохранён: {member}")
        return relative

    def image(self, part: str, rid: str, alt: str) -> str:
        relation = self.relationships(part).get(rid)
        if not relation:
            self.log("WARN", f"Не найдена связь изображения {rid} в {part}")
            return "[Отсутствует изображение]"
        kind, target, mode = relation
        if not kind.endswith(REL_IMAGE):
            self.log("WARN", f"Связь {rid} не является изображением: {kind}")
            return "[Неподдерживаемый ресурс]"
        if mode == "External":
            self.log("WARN", f"Внешняя картинка {target} не вложена в DOCX и не может быть сохранена автономно.")
            return f"![{escape(alt)}](<{safe_url(target)}>)"
        member = self.resolve(part, target)
        relative = self.store_image(member) if member else None
        return f"![{escape(alt)}]({relative})" if relative else "[Отсутствует изображение]"

    def draw(self, node: etree._Element, part: str) -> str:
        info = node.find(".//wp:docPr", NS)
        alt = (info.get("descr") or info.get("title") or info.get("name") or "Изображение") if info is not None else "Изображение"
        images: list[str] = []
        seen: set[str] = set()
        for blip in node.findall(".//a:blip", NS):
            rid = attribute(blip, "embed", R) or attribute(blip, "link", R)
            if rid and rid not in seen:
                seen.add(rid)
                images.append(self.image(part, rid, alt))
        for data in node.findall(".//v:imagedata", NS):
            rid = attribute(data, "id", R)
            if rid and rid not in seen:
                seen.add(rid)
                images.append(self.image(part, rid, alt))
        textbox = node.find(".//w:txbxContent", NS)
        extra = self.blocks(textbox, part) if textbox is not None else ""
        if textbox is not None:
            self.warn_once("textbox", "Текст графического блока перенесён, но исходное оформление/позиционирование не сохраняется.")
        if not images:
            self.warn_once("unsupported-drawing", "Найдены фигура, диаграмма или объект Word без извлекаемой картинки; он не может быть точно представлен в Markdown.")
            images.append("[Неподдерживаемый графический объект Word]")
        if extra:
            images.append(f"[Текст графического блока: {extra}]")
        return " ".join(images)

    def math(self, node: etree._Element) -> str:
        self.warn_once("math", "Обнаружены формулы Word: математическое оформление не переносится; текст формулы сохранён в квадратных скобках.")
        value = "".join(node.xpath(".//m:t/text()", namespaces=NS))
        return f"[Формула Word: {escape(value)}]" if value else "[Формула Word: не преобразована]"

    def formatted(self, raw: str, run: etree._Element) -> str:
        if not raw:
            return ""
        value = escape(raw)
        props = run.find("w:rPr", NS)
        if props is None:
            return value
        def enabled(tag: str) -> bool:
            element = props.find(f"w:{tag}", NS)
            return element is not None and attribute(element, "val") not in {"0", "false", "off", "none"}
        marks = ("**" if enabled("b") else "") + ("*" if enabled("i") else "")
        if enabled("strike") or enabled("dstrike"):
            marks += "~~"
        if not marks:
            return value
        # Put spaces outside markers to keep GFM emphasis well formed.
        match = re.fullmatch(r"(\s*)(.*?)(\s*)", value, flags=re.S)
        assert match is not None
        before, middle, after = match.groups()
        return before + (marks + middle + marks[::-1] if middle else "") + after if marks != "~~" else before + ("~~" + middle + "~~" if middle else "") + after

    def run(self, node: etree._Element, part: str) -> str:
        chunks: list[str] = []
        text: list[str] = []

        def flush() -> None:
            if text:
                chunks.append(self.formatted("".join(text), node))
                text.clear()

        for child in node:
            if child.tag == q(W, "t"):
                text.append(child.text or "")
            elif child.tag == q(W, "delText"):
                self.warn_once("tracked", "Найдены отслеживаемые изменения: удалённый текст пропущен.")
            elif child.tag == q(W, "tab"):
                text.append("    ")
            elif child.tag in {q(W, "br"), q(W, "cr")}:
                text.append("  \n")
            elif child.tag == q(W, "noBreakHyphen"):
                text.append("-")
            elif child.tag in {q(W, "drawing"), q(W, "pict"), q(W, "object")}:
                flush()
                chunks.append(self.draw(child, part))
            elif child.tag in {q(W, "footnoteReference"), q(W, "endnoteReference")}:
                flush()
                note_id = attribute(child, "id") or "?"
                kind = "Сноска" if child.tag == q(W, "footnoteReference") else "Концевая сноска"
                chunks.append(f"[{kind} {escape(note_id)}]")
            elif child.tag == q(W, "commentReference"):
                flush()
                chunks.append(f"[Комментарий {escape(attribute(child, 'id') or '?')}]")
            elif child.tag in {q(M, "oMath"), q(M, "oMathPara")}:
                flush()
                chunks.append(self.math(child))
            elif child.tag == q(W, "sym"):
                self.warn_once("symbols", "Word-символы из нестандартных шрифтов могут быть не воспроизведены.")
        flush()
        return "".join(chunks)

    def inline(self, node: etree._Element, part: str) -> str:
        pieces: list[str] = []
        for child in node:
            if child.tag == q(W, "r"):
                pieces.append(self.run(child, part))
            elif child.tag == q(W, "hyperlink"):
                content = self.inline(child, part)
                rid = attribute(child, "id", R)
                anchor = attribute(child, "anchor")
                rel = self.relationships(part).get(rid) if rid else None
                target = rel[1] if rel else ("#" + anchor if anchor else None)
                pieces.append(f"[{content}](<{safe_url(target)}>)" if target and content else content)
                if rid and not rel:
                    self.log("WARN", f"Не найдена связь гиперссылки {rid} в {part}")
            elif child.tag in {q(W, "sdt"), q(W, "sdtContent"), q(W, "fldSimple"), q(W, "smartTag"), q(W, "ins")}: 
                if child.tag == q(W, "ins"):
                    self.warn_once("tracked", "Документ содержит отслеживаемые изменения: взят видимый добавленный текст, удалённый пропущен.")
                pieces.append(self.inline(child, part))
            elif child.tag == q(W, "del"):
                self.warn_once("tracked", "Документ содержит отслеживаемые изменения: взят видимый добавленный текст, удалённый пропущен.")
            elif child.tag in {q(W, "drawing"), q(W, "pict"), q(W, "object")}:
                pieces.append(self.draw(child, part))
            elif child.tag in {q(M, "oMath"), q(M, "oMathPara")}:
                pieces.append(self.math(child))
        return "".join(pieces)

    def heading(self, paragraph: etree._Element) -> int | None:
        props = paragraph.find("w:pPr", NS)
        if props is None:
            return None
        outline = props.find("w:outlineLvl", NS)
        if outline is not None and (attribute(outline, "val") or "").isdigit():
            level = int(attribute(outline, "val") or "0") + 1
            if level <= 6:
                return level
        sid = attribute(props.find("w:pStyle", NS), "val") or ""
        style = self.styles.get(sid, {})
        val = style.get("outline")
        if isinstance(val, str) and val.isdigit() and int(val) < 6:
            return int(val) + 1
        style_name = str(style.get("name", sid)).lower().replace(" ", "")
        match = re.search(r"(?:heading|заголовок)([1-6])$", style_name)
        return int(match.group(1)) if match else None

    def list_prefix(self, paragraph: etree._Element) -> str:
        props = paragraph.find("w:pPr", NS)
        style_id = attribute(props.find("w:pStyle", NS), "val") if props is not None else None
        style = self.styles.get(style_id or "", {})
        num = props.find("w:numPr", NS) if props is not None else None
        num_id = attribute(num.find("w:numId", NS), "val") if num is not None else None
        level = attribute(num.find("w:ilvl", NS), "val") if num is not None else None
        num_id = num_id or style.get("num_id")
        level = level or style.get("level") or "0"
        if num_id and num_id != "0":
            try:
                index = min(int(str(level)), 8)
            except ValueError:
                index = 0
            fmt = self.numbering.get((str(num_id), index), "decimal")
            return "  " * index + ("- " if fmt in {"bullet", "none"} else "1. ")
        name = str(style.get("name", style_id or "")).replace(" ", "").lower()
        if "listbullet" in name or "списокмаркер" in name:
            return "- "
        if "listnumber" in name or "списокномер" in name:
            return "1. "
        return ""

    def paragraph(self, element: etree._Element, part: str) -> str:
        value = self.inline(element, part)
        if not value.strip():
            return ""
        level = self.heading(element)
        if level is not None:
            return "#" * level + " " + value.strip()
        prefix = self.list_prefix(element)
        # An ordinary paragraph beginning with Markdown syntax should stay literal.
        if not prefix and re.match(r"^(?:#{1,6}\s|[-+*]\s|\d+[.]\s|>\s)", value):
            value = "\\" + value
        return prefix + value

    def table(self, element: etree._Element, part: str) -> str:
        rows: list[list[str]] = []
        has_header = False
        for tr in element.findall("w:tr", NS):
            if tr.find("w:trPr/w:tblHeader", NS) is not None and not rows:
                has_header = True
            cells: list[str] = []
            for tc in tr.findall("w:tc", NS):
                content = self.blocks(tc, part).replace("\n", "<br>")
                cells.append(content or " ")
                tcpr = tc.find("w:tcPr", NS)
                span = tcpr.find("w:gridSpan", NS) if tcpr is not None else None
                merge = tcpr.find("w:vMerge", NS) if tcpr is not None else None
                if merge is not None:
                    self.warn_once("merged-table", "Markdown не поддерживает объединённые ячейки: структура таблицы упрощена.")
                if span is not None:
                    self.warn_once("merged-table", "Markdown не поддерживает объединённые ячейки: структура таблицы упрощена.")
                    cells.extend([" "] * max(0, min(int(attribute(span, "val") or "1"), 100) - 1))
            rows.append(cells)
        if not rows:
            return ""
        width = max(len(r) for r in rows)
        rows = [r + [" "] * (width - len(r)) for r in rows]
        def row(cells: list[str]) -> str:
            return "| " + " | ".join(cells) + " |"
        separator = row(["---"] * width)
        if has_header:
            return "\n".join([row(rows[0]), separator, *(row(r) for r in rows[1:])])
        # Use an empty header instead of misrepresenting the first data row as a heading.
        return "\n".join([row([" "] * width), separator, *(row(r) for r in rows)])

    def blocks(self, element: etree._Element, part: str) -> str:
        blocks: list[str] = []
        for child in element:
            if child.tag == q(W, "p"):
                value = self.paragraph(child, part)
            elif child.tag == q(W, "tbl"):
                if element.tag == q(W, "tc"):
                    self.warn_once("nested-table", "Вложенная таблица уплощена в текст ячейки; проверьте исходное оформление.")
                    # Still traverse the nested table, especially to extract its images.
                    value = self.table(child, part).replace("|", r"\|")
                else:
                    value = self.table(child, part)
            elif child.tag in {q(W, "sdt"), q(W, "sdtContent"), q(W, "customXml")}:
                value = self.blocks(child, part)
            elif child.tag == q(W, "altChunk"):
                self.warn_once("altchunk", "Обнаружен импортированный блок altChunk, который пока не конвертируется.")
                value = "[Неподдерживаемый импортированный блок Word]"
            elif child.tag in {q(M, "oMath"), q(M, "oMathPara")}:
                value = self.math(child)
            else:
                value = ""
            if value:
                blocks.append(value)
        return "\n\n".join(blocks)

    def convert(self) -> ConversionResult:
        part = "word/document.xml"
        if part not in self.names:
            raise ConversionError("В DOCX отсутствует word/document.xml")
        self.log("INFO", "Чтение структуры DOCX")
        root = self.xml(part)
        body = root.find("w:body", NS)
        if body is None:
            raise ConversionError("В документе отсутствует раздел body")
        md = self.blocks(body, part)
        appendix: list[str] = []
        # Include each distinct related header/footer in the Markdown, and extract its images.
        for rid, (kind, target, mode) in self.relationships(part).items():
            if not (kind.endswith("/header") or kind.endswith("/footer")) or mode == "External":
                continue
            member = self.resolve(part, target)
            if not member or member not in self.names:
                self.log("WARN", f"Не найден колонтитул {target}")
                continue
            content = self.blocks(self.xml(member), member)
            if content:
                label = "Верхний колонтитул" if kind.endswith("/header") else "Нижний колонтитул"
                appendix.append(f"### {label} ({rid})\n\n{content}")
        if appendix:
            md += "\n\n## Колонтитулы\n\n" + "\n\n".join(appendix)
        for filename, tag, heading in (
            ("word/footnotes.xml", "footnote", "Сноски"),
            ("word/endnotes.xml", "endnote", "Концевые сноски"),
            ("word/comments.xml", "comment", "Комментарии"),
        ):
            if filename not in self.names:
                continue
            items: list[str] = []
            for element in self.xml(filename).findall(f"w:{tag}", NS):
                identifier = attribute(element, "id") or "?"
                if identifier.startswith("-") or attribute(element, "type") in {"separator", "continuationSeparator"}:
                    continue
                content = self.blocks(element, filename)
                if content:
                    author = attribute(element, "author") if tag == "comment" else None
                    title = (f"Комментарий {identifier}" if tag == "comment" else
                             f"{'Сноска' if tag == 'footnote' else 'Концевая сноска'} {identifier}")
                    if author:
                        title += f" — {escape(author)}"
                    items.append(f"### {title}\n\n{content}")
            if items:
                md += f"\n\n## {heading}\n\n" + "\n\n".join(items)
        # Preserve orphaned media too; put their links in an appendix so every
        # extracted image is reachable from index.md even if its original location
        # is not representable (e.g. in a footnote or an unsupported Word object).
        extras: list[str] = []
        for member in sorted(self.names):
            if member.startswith("word/media/") and not member.endswith("/") and member not in self.image_paths:
                relative = self.store_image(member, referenced=False)
                if relative:
                    extras.append(f"![{escape(Path(member).name)}]({relative})")
        if extras:
            md += "\n\n## Дополнительные изображения\n\n" + "\n\n".join(extras)
        if any(name.startswith("word/charts/") for name in self.names):
            self.warn_once("charts", "Найдены диаграммы Word: их данные/редактируемое представление не преобразуются в Markdown (проверьте снимки диаграмм).")
        embeddings = sorted(name for name in self.names if name.startswith("word/embeddings/") and not name.endswith("/"))
        if embeddings:
            self.warn_once("ole", "Встроенные OLE-объекты сохранены как исходные файлы; их содержимое не преобразовано в Markdown.")
            directory = self.output / "assets" / "objects"
            directory.mkdir(parents=True, exist_ok=True)
            links = []
            for index, member in enumerate(embeddings, start=1):
                suffix = Path(member).suffix.lower()
                if not re.fullmatch(r"\.[a-z0-9]{1,10}", suffix):
                    suffix = ".bin"
                relative = f"assets/objects/object{index}{suffix}"
                (self.output / relative).write_bytes(self.archive.read(member))
                links.append(f"- [{escape(Path(member).name)}]({relative})")
                self.log("INFO", f"Сохранён встроенный объект: {member} → {relative}")
            md += "\n\n## Встроенные объекты Word\n\n" + "\n".join(links)
        if not md.strip():
            self.log("WARN", "В документе не найдено поддерживаемого текстового содержимого.")
            md = "<!-- Нет поддерживаемого текстового содержимого -->"
        (self.output / "index.md").write_text(md.rstrip() + "\n", encoding="utf-8")
        self.log("INFO", f"Изображений/медиафайлов: {len(self.image_paths)}; предупреждений: {len(self.warnings)}")
        self.log("INFO", "Готово: index.md и assets/images/; лог: conversion.log")
        (self.output / "conversion.log").write_text("\n".join(self.lines) + "\n", encoding="utf-8")
        return ConversionResult(self.output, self.output / "index.md", self.output / "conversion.log", len(self.image_paths), tuple(self.warnings))


def convert_docx(
    source: str | Path,
    output_dir: str | Path | None = None,
    progress: Callable[[str], None] | None = None,
    *,
    output_parent: str | Path | None = None,
) -> ConversionResult:
    """Convert once. `output_parent` places an auto-named folder inside a chosen directory.

    `output_dir` names an exact new folder. Never overwrite any existing folder.
    """
    original = Path(source).expanduser().resolve()
    if original.suffix.lower() != ".docx":
        raise ConversionError("Выберите файл с расширением .docx")
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
        suffix = 2
        while target.exists():
            target = base.with_name(f"{base.name}_{suffix}")
            suffix += 1
    else:
        target = Path(output_dir).expanduser().resolve()
        if target.exists():
            raise ConversionError(f"Папка результата уже существует: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    temp = Path(tempfile.mkdtemp(prefix=".docx-to-md-", dir=target.parent))
    try:
        try:
            # Validate that this is an actual Word package, not a renamed arbitrary ZIP.
            Document(str(original))
            with ZipFile(original) as archive:
                result = _Converter(archive, temp, progress).convert()
        except (BadZipFile, KeyError, PackageNotFoundError, etree.XMLSyntaxError, ValueError) as exc:
            raise ConversionError(f"Некорректный или неподдерживаемый DOCX: {exc}") from exc
        if target.exists():
            raise ConversionError(f"Папка результата уже существует: {target}")
        temp.rename(target)
        return ConversionResult(target, target / "index.md", target / "conversion.log", result.images_count, result.warnings)
    finally:
        if temp.exists():
            shutil.rmtree(temp)
