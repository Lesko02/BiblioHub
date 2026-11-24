import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .. import theme
from ..tools import topografie
from ..widgets import Card, PathField, ToolPage, documents_dir

XLSX = [("Excel", "*.xlsx"), ("Tutti i file", "*.*")]


class SourceList:
    """Elenco di Excel dello Scraper, con aggiunta e rimozione."""

    def __init__(self, parent, row):
        self.paths: list[str] = []
        ttk.Label(parent, text="Estrazioni dello Scraper", style="Card.TLabel").grid(
            row=row, column=0, sticky="nw", padx=(0, 12), pady=4)

        holder = ttk.Frame(parent, style="Card.TFrame")
        holder.grid(row=row, column=1, columnspan=2, sticky="ew", pady=4)
        holder.columnconfigure(0, weight=1)
        self.listbox = tk.Listbox(
            holder, height=5, activestyle="none", exportselection=False,
            selectmode="extended", font=(theme.FONT_FAMILY, 9),
            relief="flat", highlightthickness=1,
            highlightbackground=theme.BORDER, highlightcolor=theme.ACCENT,
        )
        scroll_y = ttk.Scrollbar(holder, orient="vertical", command=self.listbox.yview)
        scroll_x = ttk.Scrollbar(holder, orient="horizontal", command=self.listbox.xview)
        self.listbox.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.listbox.grid(row=0, column=0, sticky="ew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        buttons = ttk.Frame(parent, style="Card.TFrame")
        buttons.grid(row=row + 1, column=1, columnspan=2, sticky="w", pady=(0, 2))
        self.add_button = ttk.Button(buttons, text="Aggiungi…", command=self._add)
        self.add_button.pack(side="left")
        self.remove_button = ttk.Button(buttons, text="Rimuovi", command=self._remove)
        self.remove_button.pack(side="left", padx=(8, 0))

        ttk.Label(
            parent,
            text="Uno o più Excel, anche di anni diversi. Servono Insegnamento e Descrizione.",
            style="CardMuted.TLabel",
        ).grid(row=row + 2, column=1, columnspan=2, sticky="w", pady=(0, 4))
        self.widgets = (self.listbox, self.add_button, self.remove_button)
        self.next_row = row + 3

    def get_paths(self) -> list[str]:
        return list(self.paths)

    def set_paths(self, paths) -> None:
        self.paths = []
        self.listbox.delete(0, "end")
        for path in paths:
            self._append(str(path))

    def _append(self, path: str) -> None:
        if path in self.paths:
            return
        self.paths.append(path)
        self.listbox.insert("end", path)
        self.listbox.see("end")

    def _add(self) -> None:
        initial = self.paths[-1] if self.paths else ""
        initialdir = str(Path(initial).parent) if initial and Path(initial).parent.is_dir() else str(documents_dir())
        chosen = filedialog.askopenfilenames(
            title="Aggiungi estrazioni dello Scraper",
            initialdir=initialdir,
            filetypes=XLSX,
        )
        for path in chosen:
            self._append(str(Path(path)))

    def _remove(self) -> None:
        selected = list(self.listbox.curselection())
        for index in reversed(selected):
            del self.paths[index]
            self.listbox.delete(index)


class TopografiePage(ToolPage):
    key = "topografie"
    title = "Topografie"
    version = topografie.VERSION
    description = (
        "Confronta uno o più Excel dello Scraper con il registro degli esemplari "
        "d'esame. Ogni citazione riceve un punteggio in base a quanto il titolo "
        "del registro compare nella citazione. Sopra la soglia, l'esemplare va "
        "nel risultato insieme al corso in cui è stato trovato."
    )

    def build(self):
        files = Card(self.body, title="File")
        files.pack(fill="x")
        self.sources = SourceList(files, 1)
        saved_sources = self.settings.get(self.key, "sources", [])
        if isinstance(saved_sources, list):
            self.sources.set_paths(path for path in saved_sources if isinstance(path, str))

        self.registro = PathField(
            files, self.sources.next_row, "Registro topografico (.xlsx)", filetypes=XLSX,
            initial=self.settings.get(self.key, "registro", ""),
            hint="Export con la colonna Title (Complete). Ogni riga è un esemplare.")
        self.output = PathField(
            files, self.sources.next_row + 2, "File di output (.xlsx)", mode="save",
            filetypes=XLSX, default_ext=".xlsx",
            initial=self.settings.get(self.key, "output", ""))

        options = Card(self.body, title="Confronto")
        options.pack(fill="x", pady=(12, 0))
        self.threshold_var = tk.StringVar(value=str(self._saved_int("threshold", topografie.DEFAULT_THRESHOLD)))
        self.gap_var = tk.StringVar(value=str(self._saved_int("gap", topografie.DEFAULT_GAP)))
        self.min_var = tk.StringVar(value=str(self._saved_int("min_chars", topografie.DEFAULT_MIN_CHARS)))
        self.max_var = tk.StringVar(value=str(self._saved_int("max_chars", topografie.DEFAULT_MAX_CHARS)))

        threshold_row = ttk.Frame(options, style="Card.TFrame")
        threshold_row.grid(row=1, column=0, columnspan=3, sticky="w")
        ttk.Label(threshold_row, text="Soglia", style="Card.TLabel").pack(side="left")
        self.threshold_spin = ttk.Spinbox(
            threshold_row, from_=0, to=100, width=5, textvariable=self.threshold_var)
        self.threshold_spin.pack(side="left", padx=(8, 8))
        ttk.Label(threshold_row, text="su 100.  Di serie 80.", style="CardMuted.TLabel").pack(side="left")

        gap_row = ttk.Frame(options, style="Card.TFrame")
        gap_row.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(gap_row, text="Scarto", style="Card.TLabel").pack(side="left")
        self.gap_spin = ttk.Spinbox(gap_row, from_=0, to=40, width=5, textvariable=self.gap_var)
        self.gap_spin.pack(side="left", padx=(8, 8))
        ttk.Label(gap_row, text="punti fra il primo e il secondo titolo.  Di serie 8.",
                  style="CardMuted.TLabel").pack(side="left")

        length_row = ttk.Frame(options, style="Card.TFrame")
        length_row.grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(length_row, text="Citazioni sotto i", style="Card.TLabel").pack(side="left")
        self.min_spin = ttk.Spinbox(length_row, from_=0, to=400, width=5, textvariable=self.min_var)
        self.min_spin.pack(side="left", padx=(8, 6))
        ttk.Label(length_row, text="caratteri, o sopra i", style="Card.TLabel").pack(side="left")
        self.max_spin = ttk.Spinbox(length_row, from_=1, to=5000, width=6, textvariable=self.max_var)
        self.max_spin.pack(side="left", padx=(8, 6))
        ttk.Label(length_row, text="caratteri, non si confrontano.", style="Card.TLabel").pack(side="left")

        hint = ttk.Label(
            options,
            text="Sopra la soglia il titolo va in Libri. Va in Da rivedere se il secondo titolo "
                 "è entro lo scarto, se il punteggio resta fino a 10 punti sotto la soglia, "
                 "o se il titolo è breve e il cognome non è nella citazione.",
            style="CardMuted.TLabel", wraplength=680, justify="left",
        )
        hint.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 2))

        self.add_action("▶  Abbina", self.start, accent=True)
        self.register_busy(
            *self.sources.widgets, self.registro, self.output,
            self.threshold_spin, self.gap_spin, self.min_spin, self.max_spin,
        )

    def _saved_int(self, key, default):
        try:
            return int(self.settings.get(self.key, key, default))
        except (TypeError, ValueError):
            return default

    def _read_int(self, var, low, high):
        raw = var.get().strip().replace(",", ".")
        try:
            number = float(raw)
        except ValueError:
            return None
        if not number.is_integer():
            return None
        value = int(number)
        if value < low or value > high:
            return None
        return value

    def start(self):
        sources = self.sources.get_paths()
        if not sources:
            messagebox.showwarning(self.title, "Aggiungi almeno un Excel dello Scraper.", parent=self)
            return
        missing = [path for path in sources if not Path(path).is_file()]
        if missing:
            messagebox.showwarning(self.title, "File non trovato:\n" + "\n".join(missing), parent=self)
            return
        if not self.validate_input_file(self.registro):
            return
        output = self.output.get()
        if not output:
            messagebox.showwarning(self.title, "Specifica il file di output.", parent=self)
            return
        if Path(output).suffix.lower() != ".xlsx":
            output = str(Path(output).with_suffix(".xlsx"))
            self.output.set(output)

        threshold = self._read_int(self.threshold_var, 0, 100)
        gap = self._read_int(self.gap_var, 0, 100)
        min_chars = self._read_int(self.min_var, 0, 1000)
        max_chars = self._read_int(self.max_var, 1, 10000)
        if None in (threshold, gap, min_chars, max_chars):
            messagebox.showwarning(
                self.title,
                "Controlla soglia, scarto e caratteri: servono numeri interi nei limiti indicati.",
                parent=self)
            return
        if min_chars > max_chars:
            messagebox.showwarning(
                self.title, "Il minimo di caratteri non può superare il massimo.", parent=self)
            return

        registro = self.registro.get()
        self.settings.set(self.key, "sources", sources)
        self.settings.set(self.key, "registro", registro)
        self.settings.set(self.key, "output", output)
        self.settings.set(self.key, "threshold", threshold)
        self.settings.set(self.key, "gap", gap)
        self.settings.set(self.key, "min_chars", min_chars)
        self.settings.set(self.key, "max_chars", max_chars)
        self.settings.save()

        def task(log, progress, should_stop):
            return topografie.match_topografie(
                sources, registro, output,
                threshold=threshold, gap=gap,
                min_chars=min_chars, max_chars=max_chars,
                log=log, progress=progress, should_stop=should_stop,
            )

        def done(result):
            if result and result.get("output"):
                self.set_result(result["output"])

        names = "\n".join(f"· {Path(path).name}" for path in sources)
        self.run_task(
            task, done, stoppable=True,
            start_message=(f"Estrazioni:\n{names}\n"
                           f"Registro: {registro}\n"
                           f"Output:   {output}\n"
                           f"Soglia {threshold}, scarto {gap}, "
                           f"caratteri {min_chars}–{max_chars}\n"),
        )
