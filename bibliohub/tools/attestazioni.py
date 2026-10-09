"""
Generatore "Attestazioni di pagamento" (v3.0).

A partire da una tabella Excel di fatture/buoni, filtrata per intervallo
di date su "data_protocollo", genera per ogni coppia fornitore/buono:
  - un documento Word (.docx) dal template docxtpl
  - il relativo PDF (tramite Microsoft Word)
  - un file TXT con il testo dell'email di accompagnamento
"""

from contextlib import nullcontext
from datetime import datetime
from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate, RichText

VERSION = "3.0"

FONT = "Calibri"

MESI_IT = {
    1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile", 5: "Maggio",
    6: "Giugno", 7: "Luglio", 8: "Agosto", 9: "Settembre",
    10: "Ottobre", 11: "Novembre", 12: "Dicembre",
}

REQUIRED_COLUMNS = [
    "fornitore", "number", "Sistema", "identificativo", "data_protocollo",
    "totale", "buoni_di_consegna", "numero_buono", "data_buono", "CIG", "pagamento",
]

WD_FORMAT_PDF = 17


def carica_dati(excel_path, data_inizio, data_fine) -> pd.DataFrame:
    df = pd.read_excel(excel_path)

    mancanti = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if mancanti:
        raise ValueError("Colonne mancanti nella tabella Excel: " + ", ".join(mancanti))

    df["data_protocollo"] = pd.to_datetime(df["data_protocollo"], dayfirst=False, errors="coerce")

    inizio = pd.to_datetime(data_inizio)
    fine = pd.to_datetime(data_fine)

    mask = (df["data_protocollo"] >= inizio) & (df["data_protocollo"] <= fine)
    return df.loc[mask].copy()


def etichetta_sistema(valore) -> str:
    valore = str(valore).strip()
    if valore.upper() == "UGOV":
        return "UGOV"
    return valore.lower()


def richiede_modello_iva(righe: pd.DataFrame) -> bool:
    if "modello IVA" not in righe.columns:
        return False
    return righe["modello IVA"].astype(str).str.strip().str.lower().eq("si").any()


def _format_data_buono(value):
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%d/%m/%Y")
    return str(value)


def costruisci_contesto(fornitore: str, righe: pd.DataFrame) -> dict:
    fatture = []
    for _, riga in righe.iterrows():
        fatture.append({
            "number": riga["number"],
            "identificativo": riga["identificativo"],
            "data_protocollo": riga["data_protocollo"].strftime("%d/%m/%Y"),
            "totale": f'{riga["totale"]:.2f}'.replace(".", ","),
            "sistema": etichetta_sistema(riga["Sistema"]),
        })

    buoni = []
    visti = set()
    for _, riga in righe.iterrows():
        chiave = (riga["numero_buono"], riga["data_buono"], riga["pagamento"])
        if chiave not in visti:
            visti.add(chiave)
            buoni.append({
                "numero_buono": riga["numero_buono"],
                "data_buono": _format_data_buono(riga["data_buono"]),
                "pagamento": riga["pagamento"],
            })

    fatture_rt = RichText()
    for i, f in enumerate(fatture):
        fatture_rt.add("n. ", font=FONT)
        fatture_rt.add(f["number"], bold=True, font=FONT)
        fatture_rt.add(f' ({f["sistema"]} ', font=FONT)
        fatture_rt.add(f["identificativo"], bold=True, font=FONT)
        fatture_rt.add(" del ", font=FONT)
        fatture_rt.add(f["data_protocollo"], bold=True, font=FONT)
        fatture_rt.add(")", font=FONT)
        if i < len(fatture) - 1:
            fatture_rt.add("; ", font=FONT)

    fornitore_rt = RichText()
    fornitore_rt.add(fornitore, bold=True, font=FONT)

    cig_rt = RichText()
    cig_rt.add(f'CIG {righe["CIG"].iloc[0]}', bold=True, font=FONT)

    importi_rt = RichText()
    for i, f in enumerate(fatture):
        importi_rt.add(f["totale"], bold=True, font=FONT)
        if i < len(fatture) - 1:
            importi_rt.add(" + ", font=FONT)

    buoni_rt = RichText()
    for i, b in enumerate(buoni):
        buoni_rt.add(f'{b["pagamento"]} del buono ', font=FONT)
        buoni_rt.add(str(b["numero_buono"]), bold=True, font=FONT)
        buoni_rt.add(" del ", font=FONT)
        buoni_rt.add(str(b["data_buono"]), bold=True, font=FONT)
        if i < len(buoni) - 1:
            buoni_rt.add("; ", font=FONT)

    n_ddt = int(pd.to_numeric(righe["buoni_di_consegna"], errors="coerce").fillna(0).sum())
    serve_modello_iva = richiede_modello_iva(righe)

    allegati_rt = RichText()
    allegati_rt.add(str(n_ddt), bold=True, font=FONT)
    allegati_rt.add(" DDT", font=FONT)
    if serve_modello_iva:
        allegati_rt.add(" + modello IVA", font=FONT)

    return {
        "fornitore": fornitore,
        "fatture": fatture,
        "buoni": buoni,
        "n_fatture": len(fatture),
        "n_allegati": n_ddt,
        "modello_iva": serve_modello_iva,
        "fatture_rt": fatture_rt,
        "fornitore_rt": fornitore_rt,
        "cig_rt": cig_rt,
        "importi_rt": importi_rt,
        "buoni_rt": buoni_rt,
        "allegati_rt": allegati_rt,
    }


def genera_testo_email(contesto: dict, righe: pd.DataFrame) -> str:
    mese_riferimento = MESI_IT[righe["data_protocollo"].min().month]

    riga_ddt = f'- {contesto["n_allegati"]} DDT'
    if contesto["modello_iva"]:
        riga_ddt += " + Modello IVA"

    return (
        "Buongiorno,\n"
        "In allegato:\n"
        f'- Lettera di Autorizzazione Pagamento fatture {contesto["fornitore"]}\n'
        f'- {contesto["n_fatture"]} fatture\n'
        f'{riga_ddt} spedizione per il Mese di {mese_riferimento}.\n'
        "Saluti cordiali"
    )


class WordPdfConverter:
    """
    Converte .docx in PDF usando un'istanza privata di Microsoft Word,
    aperta una sola volta per tutto il lotto (non interferisce con Word già aperto).
    Va usata nello stesso thread in cui viene creata.
    """

    def __enter__(self):
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            self._word = win32com.client.DispatchEx("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = 0
        except Exception:
            pythoncom.CoUninitialize()
            raise
        return self

    def convert(self, docx_path: Path) -> Path:
        pdf_path = docx_path.with_suffix(".pdf")
        doc = self._word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        try:
            doc.SaveAs(str(pdf_path.resolve()), FileFormat=WD_FORMAT_PDF)
        finally:
            doc.Close(0)
        return pdf_path

    def __exit__(self, *exc):
        import pythoncom

        try:
            self._word.Quit()
        except Exception:
            pass
        self._word = None
        pythoncom.CoUninitialize()
        return False


def _gruppi(df: pd.DataFrame):
    for fornitore, righe_fornitore in df.groupby("fornitore"):
        for numero_buono, righe_buono in righe_fornitore.groupby("numero_buono"):
            yield fornitore, numero_buono, righe_buono


def genera_documenti(df: pd.DataFrame, template_path, output_dir, make_pdf=True,
                     log=print, progress=None, should_stop=None):
    """
    Genera DOCX (+ PDF) + TXT per ogni coppia fornitore/buono.
    Ritorna il numero di attestazioni generate.
    """
    should_stop = should_stop or (lambda: False)
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True, parents=True)

    gruppi = list(_gruppi(df))
    total = len(gruppi)
    generated = 0

    if make_pdf:
        log("Avvio di Microsoft Word per la conversione PDF...")

    with (WordPdfConverter() if make_pdf else nullcontext()) as converter:
        for fornitore, numero_buono, righe_buono in gruppi:
            if should_stop():
                log("⏹️ Generazione interrotta dall'utente.")
                break

            contesto = costruisci_contesto(fornitore, righe_buono)

            doc = DocxTemplate(str(template_path))
            doc.render(contesto)
            if "Normal" in doc.docx.styles:
                doc.docx.styles["Normal"].font.name = FONT

            data_str = righe_buono["data_protocollo"].min().strftime("%Y%m%d")
            nome_base = f"Attestazione_{fornitore}_buono{numero_buono}_{data_str}"

            path_docx = output_dir / f"{nome_base}.docx"
            doc.save(path_docx)
            log(f"📄 Generato: {path_docx.name}")

            if converter:
                path_pdf = converter.convert(path_docx)
                log(f"📕 Generato PDF: {path_pdf.name}")

            path_txt = output_dir / f"{nome_base}_mail.txt"
            path_txt.write_text(genera_testo_email(contesto, righe_buono), encoding="utf-8")
            log(f"✉️ Generato TXT: {path_txt.name}")

            generated += 1
            if progress:
                progress(generated, total)

    return generated
