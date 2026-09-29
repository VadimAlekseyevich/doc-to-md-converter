"""Minimal Tkinter interface with non-blocking conversion and a final log pane."""

from __future__ import annotations

from pathlib import Path
from queue import Empty, Queue
from threading import Thread
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.scrolledtext import ScrolledText

from .converter import convert_document


def launch() -> None:
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        raise RuntimeError("нет доступного графического дисплея; используйте CLI с путём к DOCX") from exc

    root.title("DOCX / PDF → Markdown")
    root.geometry("780x480")
    root.minsize(600, 360)
    path = tk.StringVar()
    output_parent = tk.StringVar()
    events: Queue[tuple[str, str]] = Queue()

    outer = ttk.Frame(root, padding=16)
    outer.pack(fill="both", expand=True)
    ttk.Label(outer, text="Исходный файл DOCX или PDF:").pack(anchor="w")
    selector = ttk.Frame(outer)
    selector.pack(fill="x", pady=(5, 10))
    entry = ttk.Entry(selector, textvariable=path)
    entry.pack(side="left", fill="x", expand=True)

    def choose() -> None:
        selected = filedialog.askopenfilename(filetypes=[("Документы DOCX и PDF", ("*.docx", "*.pdf")), ("Word DOCX", "*.docx"), ("PDF", "*.pdf")])
        if selected:
            path.set(selected)

    browse = ttk.Button(selector, text="Обзор…", command=choose)
    browse.pack(side="left", padx=(8, 0))

    ttk.Label(outer, text="Куда сохранить результат (пусто = рядом с DOCX):").pack(anchor="w")
    destination = ttk.Frame(outer)
    destination.pack(fill="x", pady=(5, 10))
    output_entry = ttk.Entry(destination, textvariable=output_parent)
    output_entry.pack(side="left", fill="x", expand=True)

    def choose_output() -> None:
        source = Path(path.get().strip().strip('"')).expanduser()
        initial = output_parent.get().strip() or str(source.parent if source.is_file() else Path.home())
        selected = filedialog.askdirectory(initialdir=initial, mustexist=True)
        if selected:
            output_parent.set(selected)

    output_browse = ttk.Button(destination, text="Обзор…", command=choose_output)
    output_browse.pack(side="left", padx=(8, 0))
    output_reset = ttk.Button(destination, text="Рядом с DOCX", command=lambda: output_parent.set(""))
    output_reset.pack(side="left", padx=(8, 0))
    ttk.Label(outer, text="Логи конвертации:").pack(anchor="w")
    log = ScrolledText(outer, height=15, wrap="word", state="disabled")
    log.pack(fill="both", expand=True, pady=(5, 10))

    def append(message: str) -> None:
        log.configure(state="normal")
        log.insert("end", message + "\n")
        log.see("end")
        log.configure(state="disabled")

    def worker(file_name: str, parent_name: str) -> None:
        try:
            result = convert_document(file_name, output_parent=parent_name or None,
                                  progress=lambda msg: events.put(("log", msg)))
            events.put(("done", f"Готово: {result.output_dir}"))
        except Exception as exc:  # Show unexpected failures in the GUI as well.
            events.put(("error", f"Ошибка: {exc}"))

    def start() -> None:
        file_name = path.get().strip().strip('"')
        parent_name = output_parent.get().strip().strip('"')
        if not file_name:
            append("Ошибка: укажите путь к DOCX или PDF.")
            return
        log.configure(state="normal")
        log.delete("1.0", "end")
        log.configure(state="disabled")
        button.configure(state="disabled")
        browse.configure(state="disabled")
        entry.configure(state="disabled")
        output_entry.configure(state="disabled")
        output_browse.configure(state="disabled")
        output_reset.configure(state="disabled")
        Thread(target=worker, args=(file_name, parent_name), daemon=True).start()

    def poll() -> None:
        try:
            while True:
                kind, message = events.get_nowait()
                append(message)
                if kind in {"done", "error"}:
                    button.configure(state="normal")
                    browse.configure(state="normal")
                    entry.configure(state="normal")
                    output_entry.configure(state="normal")
                    output_browse.configure(state="normal")
                    output_reset.configure(state="normal")
        except Empty:
            pass
        root.after(100, poll)

    button = ttk.Button(outer, text="Конвертировать", command=start)
    button.pack(anchor="e")
    root.after(100, poll)
    root.mainloop()
