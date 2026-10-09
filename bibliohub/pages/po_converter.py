import tkinter as tk
from pathlib import Path
from tkinter import ttk

from ..tools import po_converter
from ..widgets import Card, PathField, ToolPage

FORMATS = [("CSV", "csv"), ("Excel", "excel"), ("Entrambi (CSV + Excel)", "both")]


class POConverterPage(ToolPage):
    key = "po"
    title = "PO Line → Excel/CSV"
    version = po_converter.VERSION
    description = ("Legge l'export Excel delle linee d'ordine di Alma e produce un CSV (separatore ;) "
                   "e/o un Excel con Titolo, Autore, Editore, Anno, ISBN, Prezzo di listino "
                   "(sconto 15% reintegrato), POL, PO, Data di invio e Note di ricezione.")

    def build(self):
        files = Card(self.body, title="File")
        files.pack(fill="x")
        self.input = PathField(files, 1, "Excel PO di partenza",
                               filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")])
        self.output = PathField(
            files, 2, "File di output", mode="save", default_ext="",
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx"), ("Tutti i file", "*.*")],
            hint="Estensione facoltativa. Se vuoto: <nome input>_estratto nella stessa cartella")

        fmt = Card(self.body, title="Formato di conversione")
        fmt.pack(fill="x", pady=(12, 0))
        self.format_var = tk.StringVar(value=self.settings.get(self.key, "format", "both"))
        radios = ttk.Frame(fmt, style="Card.TFrame")
        radios.grid(row=1, column=0, columnspan=3, sticky="w")
        for text, value in FORMATS:
            rb = ttk.Radiobutton(radios, text=text, value=value, variable=self.format_var)
            rb.pack(side="left", padx=(0, 24))
            self.register_busy(rb)

        self.add_action("▶  Converti", self.start, accent=True)
        self.register_busy(self.input, self.output)

    def start(self):
        if not self.validate_input_file(self.input):
            return
        src = self.input.get()
        dst = self.output.get()
        if not dst:
            dst = str(Path(src).with_name(f"{Path(src).stem}_estratto"))
            self.output.set(dst)

        choice = self.format_var.get()
        self.settings.set(self.key, "format", choice)
        self.settings.save()
        formats = po_converter.formats_from_choice(choice)

        def task(log, progress, should_stop):
            rows, written = po_converter.convert_file(src, dst, formats)
            log(f"✅ Convertite {len(rows)} righe")
            for path in written:
                log(f"   → {path}")
            return written

        def done(written):
            if written:
                self.set_result(written[-1] if len(written) == 1 else str(Path(written[0]).parent))

        self.run_task(task, done, start_message=f"📥 Input: {src}\n")
