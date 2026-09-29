"""Minimal Tkinter interface with non-blocking conversion and a final log pane."""

from __future__ import annotations

from queue import Empty, Queue
from threading import Thread
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.scrolledtext import ScrolledText

from .converter import convert_docx


def launch() -> None:
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        raise RuntimeError("нет доступного графического дисплея; используйте CLI с путём к DOCX") from exc

    root.title("DOCX → Markdown")
    root.geometry("760x420")
    root.minsize(580, 300)
    path = tk.StringVar()
    events: Queue[tuple[str, str]] = Queue()

    outer = ttk.Frame(root, padding=16)
    outer.pack(fill="both", expand=True)
    ttk.Label(outer, text="Исходный файл DOCX:").pack(anchor="w")
    selector = ttk.Frame(outer)
    selector.pack(fill="x", pady=(5, 10))
    entry = ttk.Entry(selector, textvariable=path)
    entry.pack(side="left", fill="x", expand=True)

    def choose() -> None:
        selected = filedialog.askopenfilename(filetypes=[("Word document", "*.docx")])
        if selected:
            path.set(selected)

    browse = ttk.Button(selector, text="Обзор…", command=choose)
    browse.pack(side="left", padx=(8, 0))
    ttk.Label(outer, text="Логи конвертации:").pack(anchor="w")
    log = ScrolledText(outer, height=15, wrap="word", state="disabled")
    log.pack(fill="both", expand=True, pady=(5, 10))

    def append(message: str) -> None:
        log.configure(state="normal")
        log.insert("end", message + "\n")
        log.see("end")
        log.configure(state="disabled")

    def worker(file_name: str) -> None:
        try:
            result = convert_docx(file_name, progress=lambda msg: events.put(("log", msg)))
            events.put(("done", f"Готово: {result.output_dir}"))
        except Exception as exc:  # Show unexpected failures in the GUI as well.
            events.put(("error", f"Ошибка: {exc}"))

    def start() -> None:
        file_name = path.get().strip().strip('"')
        if not file_name:
            append("Ошибка: укажите путь к DOCX.")
            return
        log.configure(state="normal")
        log.delete("1.0", "end")
        log.configure(state="disabled")
        button.configure(state="disabled")
        browse.configure(state="disabled")
        entry.configure(state="disabled")
        Thread(target=worker, args=(file_name,), daemon=True).start()

    def poll() -> None:
        try:
            while True:
                kind, message = events.get_nowait()
                append(message)
                if kind in {"done", "error"}:
                    button.configure(state="normal")
                    browse.configure(state="normal")
                    entry.configure(state="normal")
        except Empty:
            pass
        root.after(100, poll)

    button = ttk.Button(outer, text="Конвертировать", command=start)
    button.pack(anchor="e")
    root.after(100, poll)
    root.mainloop()
