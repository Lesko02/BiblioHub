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
    description = ("Scraping: raccoglie docente e bibliografia dall'elenco A-Z. Anno accademico "
                   "e ordinamento si scelgono come sul sito. Ogni a capo della pagina resta una riga; "
                   "un elenco a) a- o 1) 1- 1. nella stessa riga viene diviso ancora. "
                   "L'anno di pubblicazione è quello in fondo alla citazione. "
                   "Le edizioni non attive sono in giallo.  Merge: solo le righe diverse.")

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

        self._filling = False
        self._year_ids = {}
        ttk.Label(tab, text="Anno accademico", style="Card.TLabel").grid(
            row=4, column=0, sticky="w", padx=(0, 12), pady=4)
        self.year_var = tk.StringVar()
        self.year_combo = ttk.Combobox(tab, textvariable=self.year_var, state="readonly")
        self.year_combo.grid(row=4, column=1, sticky="ew", pady=4)
        self.year_combo.bind("<<ComboboxSelected>>", self._on_year_selected)
        self.load_button = ttk.Button(tab, text="Carica elenco", command=self.load_options)
        self.load_button.grid(row=4, column=2, sticky="e", padx=(8, 0), pady=4)

        ttk.Label(tab, text="Ordinamento", style="Card.TLabel").grid(
            row=5, column=0, sticky="w", padx=(0, 12), pady=4)
        self.ord_var = tk.StringVar(value=scraper.ALL_ORDINAMENTI)
        self.ord_combo = ttk.Combobox(tab, textvariable=self.ord_var, state="readonly")
        self.ord_combo.grid(row=5, column=1, columnspan=2, sticky="ew", pady=4)
        self.ord_combo["values"] = (scraper.ALL_ORDINAMENTI,)
        ttk.Label(
            tab,
            text="Carica l'elenco per vedere gli anni e gli ordinamenti della pagina. "
                 "«Tutti gli ordinamenti» scarica ogni tabella.",
            style="CardMuted.TLabel",
        ).grid(row=6, column=1, columnspan=2, sticky="w", pady=(0, 4))

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
                           self.merge_right, self.merge_out, self.load_button)

    def _set_busy(self, busy, stoppable=False):
        selected = self.tabs.select()
        super()._set_busy(busy, stoppable)
        for tab_id in self.tabs.tabs():
            keep = (not busy) or tab_id == selected
            self.tabs.tab(tab_id, state="normal" if keep else "disabled")
        combo_state = "disabled" if busy else "readonly"
        self.year_combo.configure(state=combo_state)
        self.ord_combo.configure(state=combo_state)

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

        year_id = self._year_ids.get(self.year_var.get().strip(), "")
        ordinamento = self.ord_var.get().strip()
        self.settings.set(self.key, "url", url)
        self.settings.set(self.key, "output_dir", str(Path(output).parent))
        self.settings.set(self.key, "year", self.year_var.get().strip())
        self.settings.set(self.key, "ordinamento", ordinamento)
        self.settings.save()

        def task(log, progress, should_stop):
            Path(output).parent.mkdir(parents=True, exist_ok=True)
            return scraper.run_scraping(url, output, log=log, progress=progress,
                                        should_stop=should_stop, year_id=year_id,
                                        ordinamento=ordinamento)

        def done(result):
            if result:
                self.set_result(result["output"])

        year_label = self.year_var.get().strip() or "quello della pagina"
        ord_label = ordinamento or scraper.ALL_ORDINAMENTI
        self.run_task(task, done, stoppable=True,
                      start_message=f"🌐 URL: {url}\n📅 Anno: {year_label}\n"
                                    f"📚 Ordinamento: {ord_label}\n📂 Output: {output}\n")

    def _on_year_selected(self, _event=None):
        if self._filling or self.running:
            return
        if self._selected_year_id() == getattr(self, "_loaded_year_id", None):
            return
        self.load_options()

    def _selected_year_id(self):
        return self._year_ids.get(self.year_var.get().strip(), "")

    def load_options(self):
        url = self.url_var.get().strip()
        if not url.startswith("http"):
            messagebox.showwarning(self.title, "Inserisci un URL valido (deve iniziare con http).",
                                   parent=self)
            return
        year_id = self._selected_year_id()

        def task(log, progress, should_stop):
            return scraper.fetch_page_choices(url, year_id=year_id, log=log)

        self.run_task(task, self._apply_choices, start_message=f"🌐 URL: {url}\n")

    def _apply_choices(self, result):
        if not result:
            return
        years = result["years"]
        ordinamenti = result["ordinamenti"]
        self._year_ids = {label: value for label, value in years}
        labels = [label for label, _ in years]
        saved_year = self.settings.get(self.key, "year", "")
        selected_label = next((label for label, value in years if value == result["selected"]), "")
        current = self.year_var.get().strip()
        if current not in self._year_ids:
            current = saved_year if saved_year in self._year_ids else selected_label

        saved_ord = self.settings.get(self.key, "ordinamento", scraper.ALL_ORDINAMENTI)
        ord_values = [scraper.ALL_ORDINAMENTI, *ordinamenti]
        current_ord = self.ord_var.get().strip()
        if current_ord not in ord_values:
            current_ord = saved_ord if saved_ord in ord_values else scraper.ALL_ORDINAMENTI

        self._filling = True
        try:
            self.year_combo["values"] = labels
            self.year_var.set(current if current in self._year_ids else (labels[0] if labels else ""))
            self.ord_combo["values"] = ord_values
            self.ord_var.set(current_ord)
        finally:
            self._filling = False
        self._loaded_year_id = self._selected_year_id()
        if not years:
            self.log.write("Nessun menu «Anno accademico» in questa pagina.")
        if not ordinamenti:
            self.log.write("Nessun ordinamento trovato: verrà scaricata l'intera pagina.")

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
