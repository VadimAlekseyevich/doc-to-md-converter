"""CLI entry point; launch the desktop window when no input path is supplied."""

from __future__ import annotations

import argparse
import sys

from .converter import ConversionError, convert_docx


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Конвертация DOCX в Markdown с исходными изображениями")
    parser.add_argument("file", nargs="?", help="Путь к файлу .docx (без аргумента открывается окно)")
    destination = parser.add_mutually_exclusive_group()
    destination.add_argument("--output", help="Точный путь к новой папке результата (не перезаписывается)")
    destination.add_argument("--output-parent", help="Существующая папка, внутри которой создать <имя_docx>_md")
    args = parser.parse_args(argv)
    if not args.file:
        if args.output or args.output_parent:
            parser.error("--output и --output-parent требуют указания DOCX-файла")
        try:
            from .gui import launch
            launch()
        except (ImportError, RuntimeError) as exc:
            print(f"Не удалось запустить GUI: {exc}", file=sys.stderr)
            return 1
        return 0

    try:
        result = convert_docx(args.file, output_dir=args.output, output_parent=args.output_parent, progress=print)
    except (ConversionError, OSError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    print(f"Готово: {result.markdown_path}")
    return 0
