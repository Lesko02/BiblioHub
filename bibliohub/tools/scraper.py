"""
Scraper insegnamenti UNIMI (v6.2) + Merge di due estrazioni.

Scraping: visita la pagina "Elenco insegnamenti A-Z" di un corso di laurea
UNIMI, apre ogni insegnamento e ne estrae, per ciascuna edizione, docente e
materiale didattico (una riga Excel per ogni voce bibliografica).

Merge: confronta due file prodotti dallo scraping e restituisce solo le
righe che differiscono, indicando l'origine (Left / Right).
"""

import re
import time
from urllib.parse import urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

VERSION = "6.3"

DEFAULT_URL = "https://informatica.cdl.unimi.it/it/insegnamenti/elenco-insegnamenti-z"

BULLET_PATTERN = re.compile(r'^\s*(?:[·•◦●○\-–—*]+|o(?![A-Za-z]))\s*')
YEAR_PATTERN = re.compile(r'\b(200[0-9]|201[0-9]|202[0-6])\b')
ISBN_PATTERN = re.compile(r'\b(97[89]\d{10}|97[89]-\d{2}-\d{5}-\d{2}-\d|97[89]-\d{10})\b')

SCRAPING_COLUMNS = ['Insegnamento', 'Periodo', 'Edizione',
                    'Docente', 'Descrizione', 'Data', 'ISBN']
MERGE_COMPARE_COLUMNS = ['Docente', 'Descrizione', 'Data', 'ISBN']
MERGE_COLUMNS = ['Insegnamento', 'Docente', 'Descrizione', 'Data', 'ISBN', 'Origine']

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid")


def default_output_name(page_url):
    domain_first_part = urlparse(page_url).netloc.split('.')[0]
    safe_name = re.sub(r'[^a-zA-Z0-9\-_]', '_', domain_first_part)
    return f"Elenco_{safe_name}_updated.xlsx"


# ╔══════════════════════════════════════════════════════════╗
# ║              FUNZIONI SCRAPING                           ║
# ╚══════════════════════════════════════════════════════════╝

def fetch_course_list(page_url, log=print, session=None):
    """
    Visita la pagina elenco insegnamenti A-Z di Unimi e raccoglie
    tutti i corsi dalle tabelle 'table.table-insegnamenti'.
    Ritorna una lista di tuple (nome_corso, url_corso, periodo).
    Il campo 'periodo' viene letto dalla colonna 'Periodo di
    distribuzione' della tabella; se assente resta stringa vuota.
    """
    http = session or requests
    log(f"🌐 Recupero elenco corsi da: {page_url}")
    try:
        response = http.get(page_url, timeout=20)
        response.raise_for_status()
    except Exception as e:
        log(f"❌ Errore nel recupero della pagina: {e}")
        return []

    soup = BeautifulSoup(response.text, 'html.parser')
    tables = soup.find_all('table',
                           class_=lambda c: c and 'table-insegnamenti' in c)

    if not tables:
        log("⚠️ Nessuna tabella insegnamenti trovata nella pagina.")
        return []

    log(f"✅ Trovate {len(tables)} tabella/e insegnamenti.")

    courses = []
    for table in tables:
        headers = []
        thead = table.find('thead')
        if thead:
            headers = [th.get_text(strip=True).lower()
                       for th in thead.find_all(['th', 'td'])]

        periodo_col_idx = None
        for i, h in enumerate(headers):
            if 'periodo' in h:
                periodo_col_idx = i
                break

        tbody = table.find('tbody')
        if not tbody:
            continue

        for row in tbody.find_all('tr'):
            tds = row.find_all('td')
            td_name = row.find('td',
                class_=lambda c: c and 'views-field-title' in c)
            if not td_name:
                continue
            a_tag = td_name.find('a')
            if not a_tag:
                continue
            name = a_tag.get_text(strip=True)
            href = a_tag.get('href', '').strip()
            if href.startswith('/'):
                parsed = urlparse(page_url)
                href = f"{parsed.scheme}://{parsed.netloc}{href}"
            if not href:
                continue

            periodo = ''
            if periodo_col_idx is not None and periodo_col_idx < len(tds):
                periodo = tds[periodo_col_idx].get_text(strip=True)

            # Fallback: cerca td con classe che contiene 'erogazione'
            if not periodo:
                td_per = row.find('td',
                    class_=lambda c: c and 'erogazione' in c.lower())
                if td_per:
                    periodo = td_per.get_text(strip=True)

            courses.append((name, href, periodo))

    log(f"Trovati {len(courses)} corsi nell'elenco.")
    return courses


def _extract_edition_data(edition_header, edition_content_div):
    """Estrae i dati da un singolo pannello edizione (header h3 + contenuto div)."""
    nome_edizione = edition_header.get_text(strip=True) if edition_header else 'Edizione unica'

    noteanno = edition_content_div.find(
        'span', class_='noteanno',
        string=lambda t: t and 'Edizione non attiva' in t)
    is_inactive = noteanno is not None

    docente = 'Docente non trovato'
    person_div = edition_content_div.find('div', class_='field--name-uof-person')
    if person_div:
        item = person_div.find('div', class_='field--item')
        if item:
            name_tag = item.find('a')
            docente = (name_tag.get_text(strip=True)
                       if name_tag else item.get_text(strip=True))

    if is_inactive:
        material = 'Edizione non attiva'
    else:
        material = 'Materiale non trovato'
        mat_divs = edition_content_div.find_all(
            'div', class_=lambda c: c and (
                'af-program-yes-module-material' in c or
                ('af-program-yes-material' in c and
                 'af-program-yes-module-material' not in c)
            )
        )
        # parts va accumulato fuori dal ciclo: dentro, ogni blocco
        # sovrascriveva il precedente e restava solo l'ultimo modulo.
        parts = []
        for mat_div in mat_divs:
            inner = mat_div.find('div')
            if not inner:
                continue
            strong = inner.find('strong')
            if strong:
                strong.decompose()
            text = inner.get_text(separator='\n', strip=True)
            if text:
                parts.append(text)
        if parts:
            material = '\n\n'.join(parts)

    return {
        'edizione':    nome_edizione,
        'docente':     docente,
        'material':    material,
        'is_inactive': is_inactive,
    }


def scrape_course_data(url, session=None):
    """
    Recupera la pagina del corso ed estrae, per ciascuna edizione trovata:
    nome edizione (es. 'Edizione unica', 'Cognomi A-K'), docente, materiale
    e flag di edizione non attiva.

    Ritorna: (success, lista_di_dict) con chiavi
    'edizione', 'docente', 'material', 'is_inactive'.
    """
    if not url or not isinstance(url, str) or not url.startswith('http'):
        return False, []

    http = session or requests
    try:
        response = http.get(url, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
    except Exception:
        return False, []

    accordion = soup.find(
        'div', class_=lambda c: c and
        'field--name-dynamic-block-fieldugov-of-af-edizioni' in c)

    editions_data = []

    if accordion:
        edition_headers = accordion.find_all(
            'h3', class_='js-views-accordion-group-header')
        edition_contents = accordion.find_all(
            'div', class_='ui-accordion-content')

        if edition_headers and edition_contents:
            for header, content in zip(edition_headers, edition_contents):
                editions_data.append(_extract_edition_data(header, content))
        else:
            # Fallback: itera sui views-row figli diretti del view-content
            # (funziona anche quando JS non ha ancora iniettato le classi accordion)
            view_content = accordion.find('div', class_='view-content')
            if view_content:
                for vrow in view_content.find_all(
                        'div', class_='views-row', recursive=False):
                    hdr = vrow.find('h3')
                    body = vrow.find('div', class_='ui-accordion-content') or vrow
                    editions_data.append(_extract_edition_data(hdr, body))

    # Fallback finale: nessun accordion → tratta tutta la pagina come unica edizione
    if not editions_data:
        editions_data.append(_extract_edition_data(None, soup))

    return True, editions_data


def _material_rows(insegnamento, periodo, edition):
    """Trasforma il materiale di un'edizione in righe Excel (una per voce)."""
    material = edition['material']
    lines = material.split('\n') if isinstance(material, str) else [str(material)]

    rows = []
    for line in lines:
        cleaned = BULLET_PATTERN.sub('', line.strip()).strip()
        if not cleaned:
            continue
        if 'http://' in cleaned or 'https://' in cleaned:
            continue
        year_match = YEAR_PATTERN.search(cleaned)
        isbn_match = ISBN_PATTERN.search(cleaned)
        rows.append({
            'Insegnamento': insegnamento,
            'Periodo':      periodo,
            'Edizione':     edition['edizione'],
            'Docente':      edition['docente'],
            'Descrizione':  cleaned,
            'Data':         year_match.group(1) if year_match else '',
            'ISBN':         isbn_match.group(0) if isbn_match else '',
        })
    return rows


def _autofit_columns(ws, max_width=50):
    for col in ws.columns:
        longest = max((len(str(cell.value)) for cell in col if cell.value is not None),
                      default=0)
        ws.column_dimensions[col[0].column_letter].width = min(longest + 2, max_width)


def _style_header(ws):
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL


def auto_adjust_columns_and_style(filename, inactive_rows):
    """
    Aggiusta larghezze colonne, applica stile intestazione,
    evidenzia in giallo le righe 'Edizione non attiva'.
    inactive_rows: set di indici 0-based delle righe dati.
    """
    wb = load_workbook(filename)
    ws = wb.active
    _autofit_columns(ws)
    _style_header(ws)

    desc_column = None
    for idx, cell in enumerate(ws[1]):
        if cell.value == 'Descrizione':
            desc_column = idx + 1
            break

    if desc_column and inactive_rows:
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        yellow_font = Font(bold=True)
        for row_idx in inactive_rows:
            cell = ws.cell(row=row_idx + 2, column=desc_column)
            cell.fill = yellow_fill
            cell.font = yellow_font

    wb.save(filename)


def run_scraping(page_url, output_file, log=print, progress=None, should_stop=None,
                 delay=0.20):
    """
    Esegue lo scraping completo e salva l'Excel.
    Se interrotto tramite should_stop(), salva comunque i dati raccolti fino a quel punto.
    Ritorna un dict riassuntivo, oppure None se non sono stati trovati corsi.
    """
    should_stop = should_stop or (lambda: False)

    with requests.Session() as session:
        courses = fetch_course_list(page_url, log, session)
        if not courses:
            log("Nessun corso trovato. Controlla l'URL e riprova.")
            return None

        output_rows = []
        inactive_output_rows = set()
        processed = 0
        stopped = False

        for idx, (insegnamento, course_url, periodo) in enumerate(courses):
            if should_stop():
                stopped = True
                break

            log(f"🔍 [{idx + 1}/{len(courses)}] {insegnamento}")
            success, editions = scrape_course_data(course_url, session)

            if not success:
                log("   ⏭️ Saltato: errore nel recupero della pagina")
            else:
                for ed in editions:
                    if ed['is_inactive']:
                        log(f"   ⚠️ [{ed['edizione']}] Edizione non attiva")
                        inactive_output_rows.add(len(output_rows))
                        output_rows.append({
                            'Insegnamento': insegnamento,
                            'Periodo':      periodo,
                            'Edizione':     ed['edizione'],
                            'Docente':      ed['docente'],
                            'Descrizione':  'Edizione non attiva',
                            'Data':         '',
                            'ISBN':         '',
                        })
                    else:
                        output_rows.extend(_material_rows(insegnamento, periodo, ed))

            processed = idx + 1
            if progress:
                progress(processed, len(courses))
            time.sleep(delay)

    df_output = pd.DataFrame(output_rows, columns=SCRAPING_COLUMNS)
    df_output.to_excel(output_file, index=False)
    auto_adjust_columns_and_style(output_file, inactive_output_rows)

    if stopped:
        log(f"\n⏹️ Interrotto dall'utente dopo {processed}/{len(courses)} corsi "
            "(dati parziali salvati).")
    else:
        log(f"\n✅ Completato! Elaborati {len(courses)} corsi → "
            f"{len(output_rows)} righe di descrizione.")
    log(f"📂 Output salvato in '{output_file}'")
    log(f"⚠️ {len(inactive_output_rows)} edizioni non attive (evidenziate in giallo)")

    return {
        'courses': len(courses),
        'processed': processed,
        'rows': len(output_rows),
        'inactive': len(inactive_output_rows),
        'stopped': stopped,
        'output': output_file,
    }


# ╔══════════════════════════════════════════════════════════╗
# ║              FUNZIONE MERGE                              ║
# ╚══════════════════════════════════════════════════════════╝

def compare_and_merge_excel(left_path, right_path, output_path, log=print):
    """Ritorna il numero di differenze trovate (0 = file identici, nessun output scritto)."""
    log(f"Lettura file di sinistra: {left_path}")
    df_left = pd.read_excel(left_path)

    log(f"Lettura file di destra: {right_path}")
    df_right = pd.read_excel(right_path)

    for df in (df_left, df_right):
        for col in MERGE_COMPARE_COLUMNS:
            if col in df.columns:
                df[col] = df[col].fillna('').astype(str)

    output_rows = []
    processed_courses = set()

    log("Confronto dei dati in corso...")

    for course in df_left['Insegnamento'].unique():
        processed_courses.add(course)
        left_records = df_left[df_left['Insegnamento'] == course][MERGE_COMPARE_COLUMNS].to_dict('records')
        right_records = df_right[df_right['Insegnamento'] == course][MERGE_COMPARE_COLUMNS].to_dict('records')

        left_tuples = [tuple(r.values()) for r in left_records]
        right_tuples = [tuple(r.values()) for r in right_records]
        set_left, set_right = set(left_tuples), set(right_tuples)

        for row, tup in zip(left_records, left_tuples):
            if tup not in set_right:
                output_rows.append({**row, 'Insegnamento': course, 'Origine': 'Left'})

        for row, tup in zip(right_records, right_tuples):
            if tup not in set_left:
                output_rows.append({**row, 'Insegnamento': course, 'Origine': 'Right'})

    for course in df_right['Insegnamento'].unique():
        if course in processed_courses:
            continue
        for _, row in df_right[df_right['Insegnamento'] == course].iterrows():
            output_rows.append({**row[MERGE_COMPARE_COLUMNS].to_dict(),
                                'Insegnamento': course, 'Origine': 'Right'})

    if not output_rows:
        log("Nessuna differenza trovata. I file sono identici.")
        return 0

    df_out = pd.DataFrame(output_rows, columns=MERGE_COLUMNS)
    df_out.to_excel(output_path, index=False)

    wb = load_workbook(output_path)
    _style_header(wb.active)
    _autofit_columns(wb.active)
    wb.save(output_path)

    log(f"✅ Merge completato. {len(output_rows)} differenze → {output_path}")
    return len(output_rows)
