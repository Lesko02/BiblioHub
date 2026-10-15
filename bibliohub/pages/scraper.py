import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from ..tools import scraper
from ..widgets import Card, PathField, ToolPage, documents_dir

XLSX = [("Excel", "*.xlsx"), ("Tutti i file", "*.*")]


class ScraperPage(ToolPage):
    key = "scraper"
    title = "Scraper Insegnamenti"
    version = scraper.VERSION
    description = ("Scraping: raccoglie da ogni insegnamento dell'elenco A-Z docente e materiale "
                   "didattico, una riga per voce bibliografica. La colonna Ordinamento indica "
                   "la tabella di provenienza (per esempio a.a. 2026/27 (FAA)). Le edizioni non "
                   "attive sono evidenziate in giallo.  Merge: confronta due estrazioni e riporta "
                   "solo le righe diverse, indicando il file di origine.")

    def build(self):
        self.tabs = ttk.Notebook(self.body)
        self.tabs.pack(fill="x")

        # --- Scraping ---
        tab = Card(self.tabs)
        self.tabs.add(tab, text="  Scraping  ")

        ttk.Label(tab, text="URL elenco A-Z", style="Card.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 12), pady=4)
        self.url_var = tk.StringVar(value=self.settings.get(self.key, "url", scraper.DEFAULT_URL))
        self.url_entry = ttk.Entry(tab, textvariable=self.url_var)
        self.url_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=4)
        ttk.Label(tab, text="Es: https://mediazione.cdl.unimi.it/it/insegnamenti/"
                            "elenco-insegnamenti-z", style="CardMuted.TLabel").grid(
            row=1, column=1, columnspan=2, sticky="w", pady=(0, 4))

        self.scrape_out = PathField(
            tab, 2, "File di output (.xlsx)", mode="save", filetypes=XLSX,
            default_ext=".xlsx",
            hint="Se vuoto: Elenco_<corso>_updated.xlsx nell'ultima cartella usata")

        # --- Merge ---
        tab_m = Card(self.tabs)
        self.tabs.add(tab_m, text="  Merge  ")
        self.merge_left = PathField(tab_m, 0, "File LEFT (.xlsx)", filetypes=XLSX)
        self.merge_right = PathField(tab_m, 1, "File RIGHT (.xlsx)", filetypes=XLSX)
        self.merge_out = PathField(tab_m, 2, "File di output (.xlsx)", mode="save",
                                   filetypes=XLSX, default_ext=".xlsx")
        ttk.Label(
            tab_m,
            text="Questo merge distingue maiuscole e minuscole e non riporta "
                 "Periodo ed Edizione. Per il confronto nuovo usa Merge bibliografie.",
            style="CardMuted.TLabel",
            wraplength=720,
            justify="left",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))

        self.run_button = self.add_action("▶  Avvia", self.start, accent=True)
        self.register_busy(self.url_entry, self.scrape_out, self.merge_left,
                           self.merge_right, self.merge_out)

    def _set_busy(self, busy, stoppable=False):
        super()._set_busy(busy, stoppable)
        for tab_id in self.tabs.tabs():
            self.tabs.tab(tab_id, state="disabled" if busy else "normal")
        if busy:
            self.tabs.tab(self.tabs.select(), state="normal")

    def start(self):
        if self.tabs.index(self.tabs.select()) == 0:
            self.start_scraping()
        else:
            self.start_merge()

    def start_scraping(self):
        url = self.url_var.get().strip()
        if not url.startswith("http"):
            messagebox.showwarning(self.title, "Inserisci un URL valido (deve iniziare con http).",
                                   parent=self)
            return

        output = self.scrape_out.get()
        if not output:
            folder = Path(self.settings.get(self.key, "output_dir", "") or documents_dir())
            output = str(folder / scraper.default_output_name(url))
            self.scrape_out.set(output)
        if Path(output).suffix.lower() != ".xlsx":
            output = str(Path(output).with_suffix(".xlsx"))
            self.scrape_out.set(output)

        self.settings.set(self.key, "url", url)
        self.settings.set(self.key, "output_dir", str(Path(output).parent))
        self.settings.save()

        def task(log, progress, should_stop):
            Path(output).parent.mkdir(parents=True, exist_ok=True)
            return scraper.run_scraping(url, output, log=log, progress=progress,
                                        should_stop=should_stop)

        def done(result):
            if result:
                self.set_result(result["output"])

        self.run_task(task, done, stoppable=True,
                      start_message=f"🌐 URL: {url}\n📂 Output: {output}\n")

    def start_merge(self):
        if not (self.validate_input_file(self.merge_left)
                and self.validate_input_file(self.merge_right)):
            return
        output = self.merge_out.get()
        if not output:
            messagebox.showwarning(self.title, "Specifica il file di output.", parent=self)
            return
        if Path(output).suffix.lower() != ".xlsx":
            output = str(Path(output).with_suffix(".xlsx"))
            self.merge_out.set(output)

        left, right = self.merge_left.get(), self.merge_right.get()

        def task(log, progress, should_stop):
            return scraper.compare_and_merge_excel(left, right, output, log=log)

        def done(differences):
            if differences:
                self.set_result(output)

        self.run_task(task, done, start_message=f"🔀 LEFT:  {left}\n🔀 RIGHT: {right}\n"
                                                f"📂 Output: {output}\n")
