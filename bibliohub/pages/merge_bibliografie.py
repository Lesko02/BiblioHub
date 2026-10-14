import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from ..tools import merge_bibliografie as merge
from ..widgets import Card, PathField, ToolPage

XLSX = [("Excel", "*.xlsx"), ("Tutti i file", "*.*")]

MODES = [
    ("Solo differenze", merge.MODE_DIFF),
    ("Solo in comune", merge.MODE_COMMON),
]


class MergeBibliografiePage(ToolPage):
    key = "merge"
    title = "Merge bibliografie"
    version = merge.VERSION
    description = (
        "Confronta due file dello Scraper con le stesse intestazioni di colonna. "
        "Guarda solo il testo delle celle, non la formattazione, e tratta "
        "«Rossi» e «rossi» come lo stesso valore. Periodo ed Edizione restano "
        "nel risultato e distinguono due adozioni diverse."
    )

    def build(self):
        files = Card(self.body, title="File")
        files.pack(fill="x")
        self.left = PathField(
            files, 1, "Estrazione precedente (.xlsx)", filetypes=XLSX,
            hint="Anno vecchio. Nel risultato le righe solo qui sono Left.")
        self.right = PathField(
            files, 3, "Estrazione corrente (.xlsx)", filetypes=XLSX,
            hint="Anno nuovo. Nel risultato le righe solo qui sono Right.")
        self.output = PathField(
            files, 5, "File di output (.xlsx)", mode="save", filetypes=XLSX,
            default_ext=".xlsx")

        options = Card(self.body, title="Cosa mostrare")
        options.pack(fill="x", pady=(12, 0))
        self.mode_var = tk.StringVar(value=self.settings.get(self.key, "mode", merge.MODE_DIFF))
        radios = ttk.Frame(options, style="Card.TFrame")
        radios.grid(row=1, column=0, columnspan=3, sticky="w")
        for text, value in MODES:
            rb = ttk.Radiobutton(radios, text=text, value=value, variable=self.mode_var)
            rb.pack(side="left", padx=(0, 28))
            self.register_busy(rb)
        ttk.Label(
            options,
            text="Solo differenze: testi presenti in una sola estrazione.  "
                 "Solo in comune: testi già presenti nell'estrazione precedente.",
            style="CardMuted.TLabel",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 8))

        filter_row = ttk.Frame(options, style="Card.TFrame")
        filter_row.grid(row=3, column=0, columnspan=3, sticky="w")
        self.drop_short = tk.BooleanVar(value=self.settings.get(self.key, "drop_short", True))
        drop = ttk.Checkbutton(filter_row, text="Scarta descrizioni sotto i",
                               variable=self.drop_short)
        drop.pack(side="left")
        self.min_chars = tk.StringVar(value=str(self.settings.get(self.key, "min_chars", 3)))
        spin = ttk.Spinbox(filter_row, from_=1, to=50, width=4, textvariable=self.min_chars)
        spin.pack(side="left", padx=(8, 6))
        ttk.Label(filter_row, text="caratteri", style="Card.TLabel").pack(side="left")

        self.add_action("▶  Confronta", self.start, accent=True)
        self.register_busy(self.left, self.right, self.output, drop, spin)

    def _min_chars(self):
        raw = self.min_chars.get().strip()
        try:
            value = int(raw)
        except ValueError:
            return None
        if value < 1:
            return None
        return value

    def start(self):
        if not (self.validate_input_file(self.left) and self.validate_input_file(self.right)):
            return
        output = self.output.get()
        if not output:
            messagebox.showwarning(self.title, "Specifica il file di output.", parent=self)
            return
        if Path(output).suffix.lower() != ".xlsx":
            output = str(Path(output).with_suffix(".xlsx"))
            self.output.set(output)

        min_chars = self._min_chars()
        if self.drop_short.get() and min_chars is None:
            messagebox.showwarning(self.title, "Indica un numero di caratteri maggiore di zero.",
                                   parent=self)
            return

        mode = self.mode_var.get()
        drop_short = bool(self.drop_short.get())
        left, right = self.left.get(), self.right.get()

        self.settings.set(self.key, "mode", mode)
        self.settings.set(self.key, "drop_short", drop_short)
        self.settings.set(self.key, "min_chars", min_chars or 3)
        self.settings.save()

        def task(log, progress, should_stop):
            Path(output).parent.mkdir(parents=True, exist_ok=True)
            return merge.compare_bibliografie(
                left, right, output,
                mode=mode,
                drop_short=drop_short,
                min_chars=min_chars or 3,
                log=log,
            )

        def done(result):
            if result:
                self.set_result(result["output"])

        mode_label = "solo differenze" if mode == merge.MODE_DIFF else "solo in comune"
        self.run_task(
            task, done,
            start_message=(f"◀ Precedente: {left}\n"
                           f"▶ Corrente:   {right}\n"
                           f"📂 Output:    {output}\n"
                           f"Modalità: {mode_label}\n"),
        )
