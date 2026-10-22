"""Finestra della guida contestuale, aperta dal pulsante ⓘ Guida."""

import tkinter as tk
from tkinter import messagebox, ttk

from . import guide, theme
from .widgets import open_path


class GuideWindow(tk.Toplevel):
    """Testo selezionabile, un capitolo per schermata, più il manuale PDF."""

    _instance = None

    @classmethod
    def open(cls, app, key):
        current = cls._instance
        if current is not None:
            try:
                alive = current.winfo_exists()
            except tk.TclError:
                alive = False
            if alive:
                current.show(key)
                current.deiconify()
                current.lift()
                current.focus_set()
                return
        cls._instance = cls(app, key)

    def __init__(self, app, key):
        super().__init__(app)
        self.app = app
        self._filling = False
        self.title("Guida di BiblioHub")
        self.configure(background=theme.BG)
        self.transient(app)
        self.bind("<Escape>", lambda _e: self.destroy())

        scale = max(app.winfo_fpixels("1i") / 96, 1)
        width, height = int(760 * scale), int(640 * scale)
        self.minsize(int(560 * scale), int(420 * scale))
        self.update_idletasks()
        x = app.winfo_rootx() + max((app.winfo_width() - width) // 2, 0)
        y = app.winfo_rooty() + max((app.winfo_height() - height) // 3, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")

        bar = ttk.Frame(self, padding=(16, 12, 16, 0))
        bar.pack(fill="x")
        ttk.Label(bar, text="Guida", style="Section.TLabel").pack(side="left")
        self.menu = ttk.Combobox(bar, state="readonly", width=32)
        self.menu.pack(side="right")
        self._labels = {chapter.menu: chapter.key for chapter in guide.CHAPTERS}
        self.menu["values"] = list(self._labels)
        self.menu.bind("<<ComboboxSelected>>", self._on_pick)

        card = ttk.Frame(self, style="Card.TFrame", padding=12)
        card.pack(fill="both", expand=True, padx=16, pady=12)

        self.text = tk.Text(
            card, wrap="word", relief="flat", borderwidth=0, padx=12, pady=8,
            background=theme.CARD, foreground=theme.TEXT,
            font=(theme.FONT_FAMILY, 11), cursor="xterm",
            spacing1=2, spacing3=6,
        )
        scroll = ttk.Scrollbar(card, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        self.text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.text.bind("<Key>", self._readonly_key)
        self.text.tag_configure("h1", font=(theme.FONT_FAMILY, 18, "bold"),
                                foreground=theme.PRIMARY, spacing3=10)
        self.text.tag_configure("h2", font=(theme.FONT_FAMILY, 13, "bold"),
                                foreground=theme.ACCENT, spacing1=10, spacing3=4)
        self.text.tag_configure("label", font=(theme.FONT_FAMILY, 11, "bold"),
                                foreground=theme.PRIMARY, spacing1=8, spacing3=2)
        self.text.tag_configure("body", font=(theme.FONT_FAMILY, 11))
        self.text.tag_configure("item", font=(theme.FONT_FAMILY, 11), lmargin1=8, lmargin2=28)

        actions = ttk.Frame(self, padding=(16, 0, 16, 14))
        actions.pack(fill="x")
        ttk.Button(actions, text="Apri il manuale PDF", command=self.open_manual).pack(side="left")
        ttk.Button(actions, text="Chiudi", command=self.destroy).pack(side="right")

        self.show(key)

    def _on_pick(self, _event=None):
        if self._filling:
            return
        self.show(self._labels.get(self.menu.get(), "home"))

    def show(self, key):
        page = guide.chapter(key)
        self._filling = True
        try:
            self.menu.set(page.menu)
        finally:
            self._filling = False
        self._render(page)
        self.text.focus_set()

    def _render(self, page):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("end", page.title + "\n", "h1")
        for block in page.blocks:
            if block.kind == "h2":
                self.text.insert("end", block.text + "\n", "h2")
            elif block.kind == "p":
                self.text.insert("end", block.text + "\n", "body")
            elif block.kind in {"steps", "expect", "notes", "list"}:
                if block.text:
                    self.text.insert("end", block.text + "\n", "label")
                for index, item in enumerate(block.items, 1):
                    mark = f"{index}.  " if block.kind == "steps" else "•  "
                    self.text.insert("end", mark + item + "\n", "item")
        self.text.yview_moveto(0)

    def _readonly_key(self, event):
        control = event.state & 0x4
        if control and event.keysym.lower() in {"c", "a", "insert"}:
            if event.keysym.lower() == "a":
                self.text.tag_add("sel", "1.0", "end")
                return "break"
            return None
        if event.keysym in {"Left", "Right", "Up", "Down", "Home", "End",
                            "Prior", "Next", "Tab", "Shift_L", "Shift_R"}:
            return None
        return "break"

    def open_manual(self):
        path = guide.MANUAL_PDF
        if not path.is_file():
            messagebox.showwarning(
                "Guida di BiblioHub",
                f"Il manuale PDF non è stato trovato:\n{path}",
                parent=self,
            )
            return
        open_path(path)
