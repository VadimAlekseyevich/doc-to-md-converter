"""CLI entry point; launch the desktop window when no input path is supplied."""

from __future__ import annotations

import argparse
import sys

from .converter import ConversionError, convert_docx


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Конвертация DOCX в Markdown с исходными изображениями")
    parser.add_argument("file", nargs="?", help="Путь к файлу .docx (без аргумента открывается окно)")
    parser.add_argument("--output", help="Путь к новой папке результата (не перезаписывается)")
    args = parser.parse_args(argv)
    if not args.file:
        if args.output:
            parser.error("--output требует указания DOCX-файла")
        try:
            from .gui import launch
            launch()
        except (ImportError, RuntimeError) as exc:
            print(f"Не удалось запустить GUI: {exc}", file=sys.stderr)
            return 1
        return 0

    try:
        result = convert_docx(args.file, output_dir=args.output, progress=print)
    except (ConversionError, OSError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    print(f"Готово: {result.markdown_path}")
    return 0
