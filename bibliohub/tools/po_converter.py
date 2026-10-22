"""
Conversione PO Alma -> CSV/Excel (versione 2.0).

Campi in uscita:
  Titolo, Autore, Editore, Anno  <- colonna B (descrizione copia)
  ISBN                           <- Numero standard
  Prezzo                         <- Prezzo / 0.85 (sconto 15% da reintegrare), formato euro
  POL                            <- Riferimento della linea d'ordine
  PO                             <- colonna PO
  Data di Invio                  <- Data di invio
  Note di Ricezione              <- Nota di ricezione/attivazione

Uso da riga di comando:
  python -m bibliohub.tools.po_converter <input.xlsx> <output> [csv|excel|both]
"""

import csv
import re
import sys
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime, date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from ..values import as_text

VERSION = "2.1"

COPYRIGHT = "\u00a9"

FIELDNAMES = [
    "Titolo",
    "Autore",
    "Editore",
    "Anno",
    "ISBN",
    "Prezzo",
    "POL",
    "PO",
    "Data di Invio",
    "Note di Ricezione",
]

CITTA = {
    "milano", "roma", "torino", "firenze", "napoli", "bologna", "venezia",
    "padova", "genova", "palermo", "catania", "bari", "verona", "pisa",
    "siena", "parma", "modena", "perugia", "trieste", "brescia", "dueville",
    "caltanissetta", "trento", "bolzano", "udine",
    "ravenna", "rimini", "ancona", "pescara", "salerno", "lecce", "messina",
    "cagliari", "sassari", "aosta", "como", "bergamo", "vicenza", "treviso",
}

# Prezzo di listino = prezzo scontato / 0.85 (ottenuto sperimentalmente)
PRICE_FACTOR = Decimal("1.1764705882")


def clean_text(value):
    text = as_text(value)
    if not text:
        return ""
    text = text.replace("\u2019", "'").replace("\u2018", "'")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u00a0", " ")
    text = text.replace("\ufffd", "")
    text = re.sub(r"<([^>]*)>", r"\1", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def norm_header(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def find_column(df, *candidates):
    """Trova una colonna per nome normalizzato (funziona su template vecchio e nuovo)."""
    lookup = {norm_header(c): c for c in df.columns}
    for cand in candidates:
        key = norm_header(cand)
        if key in lookup:
            return lookup[key]
        for header_key, original in lookup.items():
            if key and key in header_key:
                return original
    raise KeyError("Colonna non trovata: " + ", ".join(candidates))


def is_place(token):
    t = token.strip()
    if not t:
        return False
    if t.startswith("[") and t.endswith("]"):
        inner = t[1:-1].strip().lower()
        if inner.startswith("a cura"):
            return False
        return True
    return t.lower().strip("[]") in CITTA


def is_isbn_token(token):
    solo_cifre = re.sub(r"\D", "", token)
    return token.upper() == "ISBN" or (
        len(solo_cifre) >= 10 and re.fullmatch(r"[\d\-]+", token)
    )


def is_year_token(token):
    cand = token.replace(" ", "")
    return re.fullmatch(r"(?:" + COPYRIGHT + r")?(19|20)\d{2}", cand) is not None


def parse_descrizione(raw):
    """
    Formato tipico colonna B:
      Titolo / Autore[; note], [Citta,] Editore, [©]Anno, ISBN, ISBN
    Autore puo' contenere virgole (Rigoni, Stern Mario) o piu' nomi.
    """
    desc = clean_text(raw)

    if " / " in desc:
        titolo, resto = desc.split(" / ", 1)
    else:
        titolo, resto = desc, ""

    titolo = titolo.strip()
    pezzi = [p.strip() for p in resto.split(",") if p.strip()]

    while pezzi and is_isbn_token(pezzi[-1]):
        pezzi.pop()

    anno = ""
    if pezzi and is_year_token(pezzi[-1]):
        anno = re.sub(r"\D", "", pezzi[-1])
        pezzi.pop()

    editore = ""
    if pezzi:
        editore = pezzi.pop()

    if pezzi and is_place(pezzi[-1]):
        pezzi.pop()

    autore = ", ".join(pezzi)
    # Tieni solo il nome prima di "; postfazione/introduzione..."
    if " ; " in autore:
        primo, extra = autore.split(" ; ", 1)
        if extra.lower().startswith(("postfazione", "introduzione", "prefazione", "con ")):
            autore = primo.strip()

    return titolo, clean_text(autore), clean_text(editore), anno


def parse_prezzo_listino(value):
    """Estrae il numero da '12.75 EUR' oppure da un float Excel."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return Decimal(str(value))
    text = clean_text(value).upper().replace("EUR", "").replace("€", "").strip()
    text = text.replace(" ", "").replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    if not text:
        return None
    return Decimal(text)


def format_euro(listino):
    """Prezzo scontato riportato a listino, arrotondato, es. € 15,00"""
    if listino is None:
        return ""
    lordo = (listino * PRICE_FACTOR).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    intero, decimali = f"{lordo:.2f}".split(".")
    intero_it = f"{int(intero):,}".replace(",", ".")
    return f"€ {intero_it},{decimali}"


def format_data(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    text = clean_text(value)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text[:19], fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return text


def extract_rows(input_path):
    df = pd.read_excel(input_path)

    col_pol = find_column(df, "Riferimento della linea d'ordine")
    col_desc = find_column(df, "Visualizzazione della descrizione della copia")
    col_isbn = find_column(df, "Numero standard")
    col_po = find_column(df, "PO")
    try:
        col_prezzo = find_column(df, "Prezzo totale")
    except KeyError:
        col_prezzo = find_column(df, "Prezzo")
    col_data = find_column(df, "Data di invio")
    try:
        col_note = find_column(df, "Nota di ricezione/attivazione")
    except KeyError:
        col_note = None

    rows = []
    for _, row in df.iterrows():
        titolo, autore, editore, anno = parse_descrizione(row[col_desc])
        listino = parse_prezzo_listino(row[col_prezzo])
        rows.append({
            "Titolo": titolo,
            "Autore": autore,
            "Editore": editore,
            "Anno": anno,
            "ISBN": clean_text(row[col_isbn]),
            "Prezzo": format_euro(listino),
            "POL": clean_text(row[col_pol]),
            "PO": clean_text(row[col_po]),
            "Data di Invio": format_data(row[col_data]),
            "Note di Ricezione": clean_text(row[col_note]) if col_note else "",
        })
    return rows


def write_csv(rows, output_path):
    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def write_excel(rows, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "PO estratto"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="366092")

    ws.append(FIELDNAMES)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True)

    for r_idx, row in enumerate(rows, start=2):
        for c_idx, name in enumerate(FIELDNAMES, start=1):
            cell = ws.cell(r_idx, c_idx, str(row[name]))
            cell.number_format = "@"

    for col_idx in range(1, len(FIELDNAMES) + 1):
        letter = get_column_letter(col_idx)
        max_len = max((len(str(c.value)) for c in ws[letter] if c.value is not None), default=0)
        ws.column_dimensions[letter].width = min(max_len + 3, 55)

    wb.save(output_path)


def output_base_for(output_path):
    base = Path(output_path)
    if base.suffix.lower() in {".csv", ".xlsx", ".xls"}:
        base = base.with_suffix("")
    return base


def convert_file(input_path, output_base, formats):
    """formats: iterabile con 'csv' e/o 'excel'. Ritorna (righe, percorsi_scritti)."""
    rows = extract_rows(input_path)
    base = output_base_for(output_base)

    written = []
    if "csv" in formats:
        path = base.with_suffix(".csv")
        write_csv(rows, path)
        written.append(str(path))
    if "excel" in formats:
        path = base.with_suffix(".xlsx")
        write_excel(rows, path)
        written.append(str(path))
    return rows, written


def formats_from_choice(choice):
    choice = (choice or "both").lower()
    if choice == "csv":
        return ["csv"]
    if choice in ("excel", "xlsx"):
        return ["excel"]
    return ["csv", "excel"]


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    fmt = sys.argv[3] if len(sys.argv) >= 4 else "both"
    result, written = convert_file(sys.argv[1], sys.argv[2], formats_from_choice(fmt))
    print(f"OK: {len(result)} righe")
    for path in written:
        print(" ->", path)
