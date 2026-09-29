"""CLI entry point; launch the desktop window when no input path is supplied."""

from __future__ import annotations

import argparse
import sys

from .converter import ConversionError, convert_document


def _notify_context_result(title: str, message: str, *, error: bool = False) -> None:
    """Show a short-lived dialog for Explorer integration; no resident process."""
    try:
        import tkinter as tk
        from tkinter import messagebox
    except ImportError as exc:
        print(f"{title}: {message} (диалог недоступен: {exc})",
              file=sys.stderr if error else sys.stdout)
        return
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        print(f"{title}: {message} (диалог недоступен: {exc})",
              file=sys.stderr if error else sys.stdout)
        return
    try:
        root.withdraw()
        root.attributes("-topmost", True)
        if error:
            messagebox.showerror(title, message, parent=root)
        else:
            messagebox.showinfo(title, message, parent=root)
    finally:
        root.destroy()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Конвертация DOCX и PDF в Markdown с изображениями, без OCR")
    parser.add_argument("file", nargs="?", help="Путь к файлу .docx или .pdf (без аргумента открывается окно)")
    parser.add_argument("--context-menu", action="store_true", help=argparse.SUPPRESS)
    destination = parser.add_mutually_exclusive_group()
    destination.add_argument("--output", help="Точный путь к новой папке результата (не перезаписывается)")
    destination.add_argument("--output-parent", help="Существующая папка, внутри которой создать <имя_docx>_md")
    args = parser.parse_args(argv)
    if not args.file:
        if args.output or args.output_parent or args.context_menu:
            parser.error("--output, --output-parent и --context-menu требуют указания DOCX-файла")
        try:
            from .gui import launch
            launch()
        except (ImportError, RuntimeError) as exc:
            print(f"Не удалось запустить GUI: {exc}", file=sys.stderr)
            return 1
        return 0

    try:
        result = convert_document(args.file, output_dir=args.output, output_parent=args.output_parent, progress=print)
    except (ConversionError, OSError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        if args.context_menu:
            _notify_context_result("DOCX/PDF → Markdown: ошибка", str(exc), error=True)
        return 1
    print(f"Готово: {result.markdown_path}")
    if args.context_menu:
        _notify_context_result("DOCX/PDF → Markdown", f"Конвертация завершена.\n\nРезультат:\n{result.output_dir}")
    return 0
