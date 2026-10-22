import sys
import tkinter as tk
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import messagebox, ttk

from tkcalendar import DateEntry

from .. import ASSETS_DIR
from ..tools import attestazioni
from ..widgets import Card, PathField, ToolPage, documents_dir, open_path

DATE_FMT = "%Y-%m-%d"


def format_euro(value: float) -> str:
    return "€ " + f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def previous_month(today: date) -> tuple[date, date]:
    last = today.replace(day=1) - timedelta(days=1)
    return last.replace(day=1), last


def current_month(today: date) -> tuple[date, date]:
    return today.replace(day=1), today


class AttestazioniPage(ToolPage):
    key = "attestazioni"
    title = "Attestazioni di pagamento"
    version = attestazioni.VERSION
    description = ("Filtra la tabella fatture per 'data_protocollo' nell'intervallo scelto e, per ogni "
                   "fornitore e buono, genera l'attestazione Word dal template, il PDF e il file "
                   "TXT con il testo dell'email di accompagnamento.")

    def build(self):
        files = Card(self.body, title="File")
        files.pack(fill="x")
        s = self.settings
        self.excel = PathField(files, 1, "Tabella Excel (.xlsx)",
                               filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")],
                               initial=s.get(self.key, "excel", ""))
        self.template = PathField(files, 2, "Template Word (.docx)",
                                  filetypes=[("Word", "*.docx"), ("Tutti i file", "*.*")],
                                  initial=s.get(self.key, "template",
                                                str(ASSETS_DIR / "Template.docx")))
        self.output = PathField(files, 3, "Cartella di output", mode="dir",
                                initial=s.get(self.key, "output",
                                              str(documents_dir() / "Attestazioni")))

        period = Card(self.body, title="Intervallo date (filtro su 'data_protocollo')")
        period.pack(fill="x", pady=(12, 0))
        row = ttk.Frame(period, style="Card.TFrame")
        row.grid(row=1, column=0, columnspan=3, sticky="w")

        default_start, default_end = previous_month(date.today())
        start = self._load_date("data_inizio", default_start)
        end = self._load_date("data_fine", default_end)

        ttk.Label(row, text="Dal", style="Card.TLabel").pack(side="left")
        self.date_start = self._date_entry(row, start)
        self.date_start.pack(side="left", padx=(8, 20))
        ttk.Label(row, text="al", style="Card.TLabel").pack(side="left")
        self.date_end = self._date_entry(row, end)
        self.date_end.pack(side="left", padx=(8, 28))

        prev_btn = ttk.Button(row, text="Mese scorso",
                              command=lambda: self._set_range(*previous_month(date.today())))
        prev_btn.pack(side="left", padx=(0, 6))
        curr_btn = ttk.Button(row, text="Mese corrente",
                              command=lambda: self._set_range(*current_month(date.today())))
        curr_btn.pack(side="left")

        self.pdf_var = tk.BooleanVar(value=s.get(self.key, "pdf", sys.platform == "win32"))
        pdf_check = ttk.Checkbutton(period, text="Genera anche i PDF (richiede Microsoft Word)",
                                    variable=self.pdf_var)
        pdf_check.grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))
        if sys.platform != "win32":
            self.pdf_var.set(False)
            pdf_check.configure(state="disabled")

        self.add_action("▶  Genera attestazioni", self.start, accent=True)
        self.add_action("🔍  Verifica dati", self.preview)
        self.add_action("📂  Apri cartella output", self.open_output)
        self.register_busy(self.excel, self.template, self.output, self.date_start,
                           self.date_end, prev_btn, curr_btn, pdf_check)

    def _date_entry(self, parent, value: date) -> DateEntry:
        return DateEntry(parent, width=12, date_pattern="dd/mm/yyyy", locale="it_IT",
                         firstweekday="monday", year=value.year, month=value.month,
                         day=value.day, background="#366092", foreground="white",
                         headersbackground="#E4ECF7", selectbackground="#366092")

    def _load_date(self, key, default: date) -> date:
        raw = self.settings.get(self.key, key)
        try:
            return datetime.strptime(raw, DATE_FMT).date() if raw else default
        except ValueError:
            return default

    def _set_range(self, start: date, end: date):
        self.date_start.set_date(start)
        self.date_end.set_date(end)

    def open_output(self):
        folder = Path(self.output.get() or documents_dir())
        if not folder.is_dir():
            messagebox.showinfo(self.title, f"La cartella non esiste ancora:\n{folder}", parent=self)
            return
        open_path(folder)

    def _collect(self):
        """Valida il form e ritorna (excel, template, output, inizio, fine) o None."""
        if not self.validate_input_file(self.excel):
            return None
        if not self.validate_input_file(self.template):
            return None
        output = self.output.get()
        if not output:
            messagebox.showwarning(self.title, "Seleziona la cartella di output.", parent=self)
            return None
        start, end = self.date_start.get_date(), self.date_end.get_date()
        if start > end:
            messagebox.showwarning(self.title, "La data di inizio non può essere successiva "
                                               "alla data di fine.", parent=self)
            return None

        s = self.settings
        s.set(self.key, "excel", self.excel.get())
        s.set(self.key, "template", self.template.get())
        s.set(self.key, "output", output)
        s.set(self.key, "data_inizio", start.strftime(DATE_FMT))
        s.set(self.key, "data_fine", end.strftime(DATE_FMT))
        s.set(self.key, "pdf", bool(self.pdf_var.get()))
        s.save()
        return self.excel.get(), self.template.get(), Path(output), start, end

    def _header(self, excel, start, end):
        return (f"📊 Tabella: {excel}\n"
                f"📅 Intervallo: {start:%d/%m/%Y} → {end:%d/%m/%Y}\n")

    def preview(self):
        values = self._collect()
        if not values:
            return
        excel, _template, _output, start, end = values

        def task(log, progress, should_stop):
            df = attestazioni.carica_dati(excel, start.strftime(DATE_FMT), end.strftime(DATE_FMT))
            if df.empty:
                log("Nessuna riga trovata nell'intervallo di date indicato.")
                return
            gruppi = df.groupby(["fornitore", "numero_buono"])
            log(f"Righe nell'intervallo: {len(df)}  →  attestazioni da generare: {gruppi.ngroups}\n")
            for (fornitore, buono), righe in gruppi:
                log(f"• {fornitore} — buono {attestazioni.as_text(buono)}: {len(righe)} fatture, "
                    f"totale {format_euro(righe['totale'].sum())}")

        self.run_task(task, start_message=self._header(excel, start, end))

    def start(self):
        values = self._collect()
        if not values:
            return
        excel, template, output, start, end = values
        make_pdf = bool(self.pdf_var.get())

        def task(log, progress, should_stop):
            df = attestazioni.carica_dati(excel, start.strftime(DATE_FMT), end.strftime(DATE_FMT))
            if df.empty:
                log("Nessuna riga trovata nell'intervallo di date indicato.")
                return 0
            count = attestazioni.genera_documenti(df, template, output, make_pdf=make_pdf,
                                                  log=log, progress=progress,
                                                  should_stop=should_stop)
            log(f"\n✅ Generazione completata: {count} attestazioni in {output}")
            return count

        def done(count):
            if count:
                self.set_result(output)

        self.run_task(task, done, stoppable=True,
                      start_message=self._header(excel, start, end)
                      + f"📄 Template: {template}\n📂 Output: {output}\n")
