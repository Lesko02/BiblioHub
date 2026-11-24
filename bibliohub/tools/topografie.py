"""
Topografie (v1.0).

Confronta uno o più Excel dello Scraper con un registro di esemplari d'esame
(export Alma: una riga per copia, titolo in «Title (Complete)»).

La citazione non ha una forma fissa. Il punteggio, da 0 a 100, misura quante
parole del titolo di catalogo compaiono nella citazione: le parole rare pesano
di più. L'autore, quando il titolo lo riporta dopo « / », entra per una quota
più piccola. Ogni citazione è confrontata solo con i titoli che condividono
una sua parola poco comune. Il ciclo è volutamente a un solo filo.

Soglia: l'abbinamento entra nel foglio Libri.
Scarto: se il secondo titolo è ancora entro questi punti, la citazione va
in «Da rivedere» invece che in Libri.
Sotto la soglia, fino a 10 punti in meno, la citazione resta in «Da rivedere».
Un titolo di meno di cinque parole, se il cognome non compare nella citazione,
va in «Da rivedere» con il motivo Autore assente.

Si può limitare il registro alle copie inventariate prima di una data.
Un anno da solo vale il 1° gennaio di quell'anno. Le copie senza data restano.
"""

import math
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from ..values import as_text

VERSION = "1.1"

DEFAULT_THRESHOLD = 80
DEFAULT_GAP = 8
DEFAULT_MIN_CHARS = 25
DEFAULT_MAX_CHARS = 400
NEAR_BAND = 10

TITLE_COLUMN = "Title (Complete)"
YEAR_COLUMN = "Publication Date"
INVENTORY_COLUMN = "Inventory Date (calendar)"

SCRAPER_REQUIRED = ["Insegnamento", "Descrizione"]
SCRAPER_FIELDS = [
    "Insegnamento", "Ordinamento", "Periodo", "Edizione",
    "Docente", "Descrizione", "Data", "ISBN",
]

# Intestazioni Alma ricopiate quando ci sono. La chiave è la colonna sorgente.
COPY_SOURCES = [
    ("Publication Date", "Data di pubblicazione"),
    ("Barcode", "Barcode"),
    ("Inventory Number", "Inventario"),
    ("Permanent Call Number", "Collocazione"),
    ("Library Name", "Biblioteca"),
    ("Location Name", "Ubicazione"),
    ("Location Code", "Codice ubicazione"),
    ("MMS Id", "MMS Id"),
    ("Dewey Classification", "Dewey"),
]
TEXT_COLUMNS = {
    "Barcode", "Inventario", "MMS Id", "ISBN", "Codice ubicazione",
}

# Una parola presente in più di tanti titoli non basta, da sola, a scegliere
# i candidati: la citazione non ha un aggancio distintivo.
RARE_TOKEN_DF = 40
COMMON_TOKEN_DF = 60
TOKEN_PROBE = 6
CANDIDATE_CAP = 80

TITLE_WEIGHT = 0.82
AUTHOR_WEIGHT = 0.18
YEAR_BONUS_EXACT = 4
YEAR_BONUS_CLOSE = 2
SINGLE_TOKEN_FACTOR = 0.45
# Titolo breve coperto per intero, ma senza nessuna parola dell'autore:
# non va in Libri (per esempio «articolo di giornale» contro «Il giornale»,
# o lo stesso titolo firmato da un'altra persona).
AUTHORLESS_MIN_TOKENS = 5
MAX_REFS_IN_CELL = 40

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
HIGH_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
MID_FILL = PatternFill(start_color="D6E3F0", end_color="D6E3F0", fill_type="solid")
LOW_FILL = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
REVIEW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
WRAP = Alignment(wrap_text=True, vertical="top")
TOP = Alignment(vertical="top")

_YEAR = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
_INITIAL = re.compile(r"(?<![0-9a-z])([a-z])\.")
_APOSTROPHES = str.maketrans({"’": "'", "‘": "'", "ʼ": "'", "`": "'", "´": "'"})

# Parole della formulazione di responsabilità, non dell'autore.
# Restano nel titolo: «Introduzione alla linguistica» non deve perderle.
AUTHOR_NOISE = frozenset("""
    traduzione tradotto traduttore prefazione postfazione introduzione
    disegni contributo contributi consulenza progetto direzione prologo
    edited preface introduction translated translation presentazione
    illustrazione illustrazioni curatore curatori appendice
""".split())

STOP = frozenset("""
    il lo la i gli le un uno una di del dello della dei degli delle e ed a ad da in con
    per su tra fra che non come nel nella nei negli sul sui sulle dall dalle
    the of and by for to an from as at or nor not
    et al cura presso vol volume volumes ed eds edition edizione pp pag pagina pagine
    isbn doi http https www com org
    bibliografia obbligatoria consigliata riferimento riferimenti testo testi parte parti
    cap capp capitolo capitoli
    new press university
""".split())

LIBRI_COLUMNS = [
    "Titolo", "Data di pubblicazione", "Anno", "Barcode", "Inventario", "Collocazione",
    "Biblioteca", "Ubicazione", "Codice ubicazione", "MMS Id", "Dewey",
    "Punteggio", "Riferimenti (n)", "Riferimenti",
]
DETTAGLIO_COLUMNS = [
    "File", "Foglio", "Riga", "Insegnamento", "Ordinamento", "Periodo", "Edizione",
    "Docente", "Descrizione", "Data", "ISBN",
    "Punteggio", "Secondo punteggio", "Scarto",
    "Titolo", "Data di pubblicazione", "Anno citazione", "Anno registro",
    "Barcode", "Inventario", "Collocazione", "Biblioteca", "Ubicazione", "MMS Id",
]
REVIEW_COLUMNS = [
    "File", "Foglio", "Riga", "Insegnamento", "Docente", "Descrizione", "Data",
    "Anno citazione", "Motivo", "Punteggio", "Titolo proposto", "Anno registro",
    "Collocazione", "Barcode", "Inventario",
    "Secondo punteggio", "Secondo titolo", "Seconda collocazione",
]
FILE_COLUMNS = [
    "File", "Righe", "Corte", "Lunghe", "Poche parole", "Senza parola rara",
    "Abbinate", "Da rivedere", "Senza abbinamento",
]


@dataclass
class Copy:
    values: dict[str, str]
    year: int | None


@dataclass
class Work:
    title: str
    title_tokens: list[str]
    author_tokens: list[str]
    copies: list[Copy]

    @property
    def years(self) -> set[int]:
        return {copy.year for copy in self.copies if copy.year is not None}


@dataclass
class FileStats:
    name: str
    path: str
    sheet: str
    rows: int = 0
    short: int = 0
    long: int = 0
    thin: int = 0
    bland: int = 0
    matched: int = 0
    review: int = 0


@dataclass
class Citation:
    origin: FileStats
    excel_row: int
    values: dict[str, str]
    words: list[str]
    word_set: set[str]
    joins: set[str]
    long_words: list[str]
    initials: set[str]
    year: int | None


@dataclass
class Outcome:
    citation: Citation
    work: Work
    score: float
    second_score: float
    second_work: Work | None
    reason: str = ""


@dataclass
class MatchResult:
    libri: list[dict] = field(default_factory=list)
    dettaglio: list[dict] = field(default_factory=list)
    review: list[dict] = field(default_factory=list)
    files: list[FileStats] = field(default_factory=list)
    registro_path: str = ""
    registro_rows: int = 0
    works: int = 0
    skipped_works: int = 0
    copies_matched: int = 0
    titles_matched: int = 0
    comparisons: int = 0
    threshold: float = DEFAULT_THRESHOLD
    gap: float = DEFAULT_GAP
    min_chars: int = DEFAULT_MIN_CHARS
    max_chars: int = DEFAULT_MAX_CHARS
    stopped: bool = False
    output: str = ""
    inventory_before: str = ""
    excluded_copies: int = 0
    undated_copies: int = 0


def _fold(text: str) -> str:
    folded = unicodedata.normalize("NFKD", str(text).casefold())
    folded = "".join(ch for ch in folded if not unicodedata.combining(ch))
    return folded.translate(_APOSTROPHES)


def _analyze(text: str) -> tuple[list[str], set[str]]:
    """Parole utili e iniziali puntate (U. Eco) di un testo già letto."""
    folded = _fold(text)
    initials = {match.group(1) for match in _INITIAL.finditer(folded)}
    cleaned = []
    for ch in folded:
        cleaned.append(ch if ch.isalnum() else " ")
    words = []
    for token in "".join(cleaned).split():
        if len(token) < 3 or token in STOP or token.isdigit():
            continue
        words.append(token)
    return words, initials


def _years(text: str) -> list[int]:
    return [int(match) for match in _YEAR.findall(text)]


def _last_year(text: str) -> int | None:
    found = _years(text)
    return found[-1] if found else None


def parse_cutoff(text: str) -> date:
    """
    Data limite dell'inventario.
    Un anno da solo («2024») vale il 1° gennaio di quell'anno.
    """
    raw = str(text).strip()
    if re.fullmatch(r"\d{4}", raw):
        year = int(raw)
        if not 1900 <= year <= 2100:
            raise ValueError(
                f"Anno non valido: {raw}. Usa un anno tra il 1900 e il 2100, "
                "oppure una data (01/01/2024)."
            )
        return date(year, 1, 1)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    raise ValueError(
        f"Data non valida: {raw}. Usa un anno (2024) oppure un giorno (01/01/2024)."
    )


def _inventory_date(value) -> date | None:
    try:
        if value is None or pd.isna(value):
            return None
    except TypeError:
        if value is None:
            return None
    if isinstance(value, datetime):
        try:
            found = value.date()
        except (ValueError, OSError):
            return None
        if isinstance(found, date) and not isinstance(found, datetime):
            return found
        return None
    if isinstance(value, date):
        return value
    text = as_text(value)
    if not text:
        return None
    try:
        return parse_cutoff(text)
    except ValueError:
        return None


def _split_title(raw: str) -> tuple[str, str]:
    if " / " in raw:
        left, right = raw.split(" / ", 1)
    elif "/" in raw:
        left, right = raw.split("/", 1)
    else:
        left, right = raw, ""
    return left, right


def _same_path(left, right) -> bool:
    try:
        return Path(left).resolve().as_posix().casefold() == Path(right).resolve().as_posix().casefold()
    except OSError:
        return str(left).casefold() == str(right).casefold()


def _dedupe(paths) -> list[str]:
    seen = set()
    unique = []
    for path in paths:
        try:
            key = str(Path(path).resolve()).casefold()
        except OSError:
            key = str(path).casefold()
        if key in seen:
            continue
        seen.add(key)
        unique.append(str(path))
    return unique


def _read_excel(path, log):
    log(f"Lettura: {path}")
    book = pd.ExcelFile(path)
    try:
        sheet = book.sheet_names[0]
        frame = book.parse(sheet)
        if len(book.sheet_names) > 1:
            log(f"Più fogli nel file: uso il primo, «{sheet}».")
    finally:
        book.close()
    frame.columns = [str(name).strip().lstrip("\ufeff") for name in frame.columns]
    return frame, sheet


def _load_registro(path, log, inventory_before: date | None = None) -> tuple[list[Work], dict]:
    frame, _sheet = _read_excel(path, log)
    if TITLE_COLUMN not in frame.columns:
        raise ValueError(
            "Il registro non ha la colonna «Title (Complete)».\n"
            f"{path}"
        )
    if inventory_before is not None and INVENTORY_COLUMN not in frame.columns:
        raise ValueError(
            "Per filtrare sull'inventario serve la colonna "
            f"«{INVENTORY_COLUMN}».\n{path}"
        )
    present = [(source, label) for source, label in COPY_SOURCES if source in frame.columns]
    grouped: dict[tuple[str, ...], Work] = {}
    skipped = 0
    excluded = 0
    undated = 0
    for _, row in frame.iterrows():
        missing_date = False
        if inventory_before is not None:
            inventoried = _inventory_date(row[INVENTORY_COLUMN])
            if inventoried is None:
                missing_date = True
            elif inventoried >= inventory_before:
                excluded += 1
                continue
        raw = " ".join(as_text(row[TITLE_COLUMN]).split())
        left, right = _split_title(raw)
        title_tokens, _unused = _analyze(left)
        author_tokens, _unused = _analyze(right)
        author_tokens = [token for token in author_tokens if token not in AUTHOR_NOISE]
        if not title_tokens:
            skipped += 1
            continue
        key = tuple(title_tokens + ["/"] + author_tokens)
        work = grouped.get(key)
        if work is None:
            work = Work(raw, title_tokens, author_tokens, [])
            grouped[key] = work
        year_text = as_text(row[YEAR_COLUMN]) if YEAR_COLUMN in frame.columns else ""
        values = {label: as_text(row[source]) for source, label in present}
        work.copies.append(Copy(values, _last_year(year_text)))
        if missing_date:
            undated += 1
    works = [work for work in grouped.values() if work.copies]
    if not works:
        raise ValueError(
            "Nel registro non c'è nessun titolo con parole confrontabili"
            + (" prima della data indicata." if inventory_before else ".")
            + f"\n{path}"
        )
    copies = sum(len(work.copies) for work in works)
    log(f"Registro: {len(frame)} righe, {len(works)} titoli confrontabili "
        f"({copies} esemplari), {skipped} titoli scartati perché senza parole utili.")
    if inventory_before is not None:
        log(f"Inventario prima del {inventory_before.strftime('%d/%m/%Y')}: "
            f"esclusi {excluded} esemplari da quella data in poi, "
            f"{undated} senza data restano nel confronto.")
    return works, {
        "rows": len(frame),
        "skipped_works": skipped,
        "excluded": excluded,
        "undated": undated,
    }


def _load_citation(origin: FileStats, excel_row: int, values: dict[str, str],
                   min_chars: int, max_chars: int) -> Citation | None:
    description = values.get("Descrizione", "")
    length = len(description.strip())
    origin.rows += 1
    if length < min_chars:
        origin.short += 1
        return None
    if length > max_chars:
        origin.long += 1
        return None
    words, initials = _analyze(description)
    if len(words) < 2:
        origin.thin += 1
        return None
    return Citation(
        origin=origin,
        excel_row=excel_row,
        values=values,
        words=words,
        word_set=set(words),
        joins={left + right for left, right in zip(words, words[1:])},
        long_words=[word for word in words if len(word) >= 7],
        initials=initials,
        year=_last_year(values.get("Data", "")) or _last_year(description),
    )


def _load_scrapers(paths, min_chars, max_chars, log) -> tuple[list[FileStats], list[Citation]]:
    stats = []
    citations = []
    for path in paths:
        frame, sheet = _read_excel(path, log)
        missing = [name for name in SCRAPER_REQUIRED if name not in frame.columns]
        if missing:
            raise ValueError(
                "Il file non ha le colonne minime dello Scraper "
                "(Insegnamento e Descrizione).\n"
                f"{path}\n"
                f"Colonne mancanti: {', '.join(missing)}"
            )
        origin = FileStats(Path(path).name, str(path), sheet)
        for offset, (_, row) in enumerate(frame.iterrows()):
            values = {
                name: as_text(row[name]) if name in frame.columns else ""
                for name in SCRAPER_FIELDS
            }
            citation = _load_citation(origin, offset + 2, values, min_chars, max_chars)
            if citation is not None:
                citations.append(citation)
        stats.append(origin)
        log(f"{origin.name}: {origin.rows} righe, "
            f"{origin.short} corte, {origin.long} lunghe, "
            f"{origin.thin} con troppe poche parole.")
    return stats, citations


def _within_one_edit(left: str, right: str) -> bool:
    """True se le due parole coincidono a una lettera (manca, è in più, o è diversa)."""
    if left == right:
        return True
    if abs(len(left) - len(right)) > 1:
        return False
    if len(left) == len(right):
        return sum(a != b for a, b in zip(left, right)) == 1
    if len(left) > len(right):
        left, right = right, left
    index_left = index_right = skips = 0
    while index_left < len(left) and index_right < len(right):
        if left[index_left] != right[index_right]:
            skips += 1
            if skips > 1:
                return False
            index_right += 1
            continue
        index_left += 1
        index_right += 1
    return True


def _token_covered(token: str, citation: Citation) -> bool:
    if token in citation.word_set or token in citation.joins:
        return True
    if len(token) < 8:
        return False
    for other in citation.long_words:
        if abs(len(other) - len(token)) > 1:
            continue
        if _within_one_edit(token, other):
            return True
    return False


def _title_coverage(work: Work, citation: Citation) -> list[str]:
    covered = set()
    for token in work.title_tokens:
        if _token_covered(token, citation):
            covered.add(token)
    tokens = work.title_tokens
    for left, right in zip(tokens, tokens[1:]):
        if left + right in citation.word_set:
            covered.add(left)
            covered.add(right)
    return [token for token in dict.fromkeys(tokens) if token in covered]


def _author_word_matches(token: str, citation: Citation) -> bool:
    """Il cognome conta anche con una lettera di differenza, se è abbastanza lungo."""
    if token in citation.word_set:
        return True
    if len(token) < 6:
        return False
    for other in citation.words:
        if len(other) < 6 or other[0] != token[0] or abs(len(other) - len(token)) > 1:
            continue
        if _within_one_edit(token, other):
            return True
    return False


def _matched_authors(work: Work, citation: Citation) -> set[str]:
    authors = list(dict.fromkeys(work.author_tokens))
    found = {token for token in authors if _author_word_matches(token, citation)}
    if found:
        for token in authors:
            if token not in found and token[:1] in citation.initials:
                found.add(token)
    return found


def _author_hit(work: Work, citation: Citation) -> float:
    authors = list(dict.fromkeys(work.author_tokens))
    if not authors:
        return 0.0
    return len(_matched_authors(work, citation)) / len(authors)


def _surname_confirmed(work: Work, citation: Citation) -> bool:
    """Il cognome, la parola più lunga dell'autore, deve comparire nella citazione."""
    authors = list(dict.fromkeys(work.author_tokens))
    if not authors:
        return False
    longest = max(len(token) for token in authors)
    found = _matched_authors(work, citation)
    return any(len(token) == longest and token in found for token in authors)


def _year_bonus(citation_year: int | None, catalog_years: set[int]) -> float:
    if citation_year is None or not catalog_years:
        return 0.0
    best = 0.0
    for year in catalog_years:
        diff = abs(citation_year - year)
        if diff == 0:
            return float(YEAR_BONUS_EXACT)
        if diff == 1:
            best = float(YEAR_BONUS_CLOSE)
    return best


def _score(work: Work, citation: Citation, doc_freq: dict[str, int], n_docs: int) -> tuple[float, bool]:
    """Punteggio e, se vero, il titolo è troppo breve per fidarsi senza l'autore."""
    unique_title = list(dict.fromkeys(work.title_tokens))
    covered = _title_coverage(work, citation)
    if not covered or not unique_title:
        return 0.0, False

    def idf(token: str) -> float:
        return math.log((n_docs + 1) / (doc_freq.get(token, 0) + 0.5))

    denominator = sum(idf(token) for token in unique_title)
    if denominator <= 0:
        return 0.0, False
    recall = sum(idf(token) for token in covered) / denominator
    author = _author_hit(work, citation) if work.author_tokens else 0.0
    if work.author_tokens:
        score = 100 * (TITLE_WEIGHT * recall + AUTHOR_WEIGHT * author)
    else:
        score = 100 * recall
    if len(covered) == 1:
        only = covered[0]
        if doc_freq.get(only, 1) > 3 or len(only) < 7:
            score *= SINGLE_TOKEN_FACTOR
    score += _year_bonus(citation.year, work.years)
    author_gap = (
        bool(work.author_tokens)
        and len(unique_title) < AUTHORLESS_MIN_TOKENS
        and not _surname_confirmed(work, citation)
    )
    return round(min(100.0, score), 1), author_gap


def _candidates(citation: Citation, postings: dict[str, list[int]],
                doc_freq: dict[str, int]) -> list[int] | None:
    """None se la citazione non ha una parola rara presente nel registro."""
    present = [token for token in dict.fromkeys(citation.words) if token in postings]
    present.sort(key=lambda token: (doc_freq[token], token))
    if not present or doc_freq[present[0]] > RARE_TOKEN_DF:
        return None
    chosen: list[int] = []
    seen = set()
    for token in present[:TOKEN_PROBE]:
        if doc_freq[token] > COMMON_TOKEN_DF and seen:
            break
        for index in postings[token]:
            if index not in seen:
                seen.add(index)
                chosen.append(index)
        if len(seen) >= CANDIDATE_CAP:
            break
    return chosen


def _index(works: list[Work]) -> tuple[dict[str, int], dict[str, list[int]]]:
    doc_freq: dict[str, int] = defaultdict(int)
    postings: dict[str, list[int]] = defaultdict(list)
    for index, work in enumerate(works):
        for token in set(work.title_tokens) | set(work.author_tokens):
            doc_freq[token] += 1
            postings[token].append(index)
    return doc_freq, postings


def _compare(works, citations, doc_freq, postings, threshold, gap, log, progress, should_stop):
    accepted: list[Outcome] = []
    review: list[Outcome] = []
    comparisons = 0
    total = len(citations)
    review_floor = max(0.0, threshold - NEAR_BAND)
    n_docs = len(works)
    stopped = False

    for done, citation in enumerate(citations, start=1):
        if should_stop and should_stop():
            stopped = True
            log("Interruzione richiesta: salvo i confronti già fatti.")
            break
        if progress and (done == total or done % 25 == 0):
            progress(done, total)

        chosen = _candidates(citation, postings, doc_freq)
        if not chosen:
            citation.origin.bland += 1
            continue
        comparisons += len(chosen)

        best_score = 0.0
        best_work = None
        best_author_gap = False
        second_score = 0.0
        second_work = None
        for index in chosen:
            score, author_gap = _score(works[index], citation, doc_freq, n_docs)
            if score > best_score:
                if best_work is not None and best_score > second_score:
                    second_score = best_score
                    second_work = best_work
                best_score = score
                best_work = works[index]
                best_author_gap = author_gap
            elif score > second_score:
                second_score = score
                second_work = works[index]

        if best_work is None or best_score <= 0:
            continue
        margin = best_score - second_score
        if best_score >= threshold and margin > gap and not best_author_gap:
            citation.origin.matched += 1
            accepted.append(Outcome(citation, best_work, best_score, second_score, second_work))
        elif best_score >= review_floor or (best_author_gap and best_score >= threshold):
            citation.origin.review += 1
            if best_author_gap and best_score >= threshold:
                reason = "Autore assente"
            elif best_score >= threshold:
                reason = "Scarto basso"
            else:
                reason = "Sotto soglia"
            review.append(Outcome(
                citation, best_work, best_score, second_score, second_work, reason,
            ))

    if progress and not stopped:
        progress(total, total)
    return accepted, review, comparisons, stopped


def _join_unique(work: Work, label: str, limit: int = 12) -> str:
    values = []
    seen = set()
    for copy in work.copies:
        value = copy.values.get(label, "")
        if value and value not in seen:
            seen.add(value)
            values.append(value)
    if len(values) > limit:
        extra = len(values) - limit
        return " | ".join(values[:limit]) + f" | … (+{extra})"
    return " | ".join(values)


def _reference_line(outcome: Outcome) -> str:
    citation = outcome.citation
    description = " ".join(citation.values.get("Descrizione", "").split())
    if len(description) > 160:
        description = description[:157] + "…"
    parts = [
        citation.origin.name,
        f"riga {citation.excel_row}",
        citation.values.get("Insegnamento", ""),
    ]
    teacher = citation.values.get("Docente", "")
    if teacher:
        parts.append(teacher)
    parts.append(f"{outcome.score:.0f}")
    parts.append(description)
    return " · ".join(part for part in parts if part)


def _copy_row(base: dict, copy: Copy, work: Work) -> dict:
    row = dict(base)
    row["Titolo"] = work.title
    row["Anno registro"] = "" if copy.year is None else str(copy.year)
    for label in (
        "Data di pubblicazione", "Barcode", "Inventario", "Collocazione",
        "Biblioteca", "Ubicazione", "MMS Id",
    ):
        row[label] = copy.values.get(label, "")
    return row


def _build_tables(accepted: list[Outcome], review: list[Outcome]) -> tuple[list[dict], list[dict], list[dict], int, int]:
    by_copy: dict[tuple[int, int], dict] = {}
    detail = []
    for outcome in accepted:
        work_id = id(outcome.work)
        for copy_index, copy in enumerate(outcome.work.copies):
            bucket = by_copy.get((work_id, copy_index))
            if bucket is None:
                bucket = {
                    "copy": copy,
                    "work": outcome.work,
                    "score": outcome.score,
                    "hits": [],
                }
                by_copy[(work_id, copy_index)] = bucket
            bucket["score"] = max(bucket["score"], outcome.score)
            bucket["hits"].append(outcome)

            citation = outcome.citation
            base = {
                "File": citation.origin.name,
                "Foglio": citation.origin.sheet,
                "Riga": citation.excel_row,
                "Insegnamento": citation.values.get("Insegnamento", ""),
                "Ordinamento": citation.values.get("Ordinamento", ""),
                "Periodo": citation.values.get("Periodo", ""),
                "Edizione": citation.values.get("Edizione", ""),
                "Docente": citation.values.get("Docente", ""),
                "Descrizione": citation.values.get("Descrizione", ""),
                "Data": citation.values.get("Data", ""),
                "ISBN": citation.values.get("ISBN", ""),
                "Punteggio": outcome.score,
                "Secondo punteggio": outcome.second_score or "",
                "Scarto": round(outcome.score - outcome.second_score, 1),
                "Anno citazione": "" if citation.year is None else str(citation.year),
            }
            detail.append(_copy_row(base, copy, outcome.work))

    libri = []
    for bucket in by_copy.values():
        hits = sorted(bucket["hits"], key=lambda item: (-item.score, item.citation.origin.name,
                                                        item.citation.excel_row))
        shown = hits[:MAX_REFS_IN_CELL]
        lines = [_reference_line(item) for item in shown]
        hidden = len(hits) - len(shown)
        if hidden:
            lines.append(f"… e altri {hidden} riferimenti, nel foglio Dettaglio.")
        copy = bucket["copy"]
        libri.append({
            "Titolo": bucket["work"].title,
            "Data di pubblicazione": copy.values.get("Data di pubblicazione", ""),
            "Anno": "" if copy.year is None else str(copy.year),
            "Barcode": copy.values.get("Barcode", ""),
            "Inventario": copy.values.get("Inventario", ""),
            "Collocazione": copy.values.get("Collocazione", ""),
            "Biblioteca": copy.values.get("Biblioteca", ""),
            "Ubicazione": copy.values.get("Ubicazione", ""),
            "Codice ubicazione": copy.values.get("Codice ubicazione", ""),
            "MMS Id": copy.values.get("MMS Id", ""),
            "Dewey": copy.values.get("Dewey", ""),
            "Punteggio": bucket["score"],
            "Riferimenti (n)": len(hits),
            "Riferimenti": "\n".join(lines),
        })

    review_rows = []
    for outcome in review:
        citation = outcome.citation
        second = outcome.second_work
        review_rows.append({
            "File": citation.origin.name,
            "Foglio": citation.origin.sheet,
            "Riga": citation.excel_row,
            "Insegnamento": citation.values.get("Insegnamento", ""),
            "Docente": citation.values.get("Docente", ""),
            "Descrizione": citation.values.get("Descrizione", ""),
            "Data": citation.values.get("Data", ""),
            "Anno citazione": "" if citation.year is None else str(citation.year),
            "Motivo": outcome.reason,
            "Punteggio": outcome.score,
            "Titolo proposto": outcome.work.title,
            "Anno registro": _year_list(outcome.work),
            "Collocazione": _join_unique(outcome.work, "Collocazione"),
            "Barcode": _join_unique(outcome.work, "Barcode"),
            "Inventario": _join_unique(outcome.work, "Inventario"),
            "Secondo punteggio": outcome.second_score or "",
            "Secondo titolo": "" if second is None else second.title,
            "Seconda collocazione": "" if second is None else _join_unique(second, "Collocazione"),
        })

    libri.sort(key=lambda row: (-float(row["Punteggio"]), row["Titolo"], row["Collocazione"]))
    detail.sort(key=lambda row: (row["Titolo"], -float(row["Punteggio"]), row["File"], row["Riga"]))
    review_rows.sort(key=lambda row: (-float(row["Punteggio"]), row["File"], row["Riga"]))
    return libri, detail, review_rows, len(by_copy), len({id(item.work) for item in accepted})


def _year_list(work: Work) -> str:
    years = sorted({copy.year for copy in work.copies if copy.year is not None})
    return " | ".join(str(year) for year in years)


def _style_header(ws, row_index: int) -> None:
    for cell in ws[row_index]:
        if cell.value:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(vertical="center")


def _apply_body(ws, headers, threshold: float) -> None:
    wrap_headers = {"Descrizione", "Riferimenti", "Titolo", "Titolo proposto", "Secondo titolo"}
    score_headers = {"Punteggio", "Secondo punteggio", "Scarto"}
    if ws.max_row >= 2:
        body = ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers))
    else:
        body = ()
    for row in body:
        for cell, header in zip(row, headers):
            if header in wrap_headers:
                cell.alignment = WRAP
            else:
                cell.alignment = TOP
            if header in TEXT_COLUMNS:
                cell.number_format = "@"
            elif header in score_headers and cell.value != "":
                cell.number_format = "0.0"
            if header == "Punteggio" and isinstance(cell.value, (int, float)):
                if cell.value >= 90:
                    cell.fill = HIGH_FILL
                elif cell.value >= threshold:
                    cell.fill = MID_FILL
                else:
                    cell.fill = LOW_FILL
            elif header == "Motivo":
                cell.fill = LOW_FILL if cell.value == "Sotto soglia" else REVIEW_FILL
                cell.font = Font(bold=True)
    widths = {
        "Titolo": 54, "Titolo proposto": 54, "Secondo titolo": 46,
        "Descrizione": 62, "Riferimenti": 78, "Collocazione": 28,
        "Seconda collocazione": 28, "Insegnamento": 36, "Docente": 22,
        "File": 26, "Ordinamento": 28, "Biblioteca": 28, "Ubicazione": 22,
        "Dewey": 22, "Data di pubblicazione": 18,
    }
    for index, header in enumerate(headers, start=1):
        letter = get_column_letter(index)
        ws.column_dimensions[letter].width = widths.get(header, 16)
    last_row = max(ws.max_row, 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"
    ws.freeze_panes = "A2"


def _write_table(ws, headers, records, threshold: float) -> None:
    ws.append(headers)
    for record in records:
        ws.append([record.get(header, "") for header in headers])
    _style_header(ws, 1)
    _apply_body(ws, headers, threshold)


def _write_summary(ws, result: MatchResult) -> None:
    ws["A1"] = "Voce"
    ws["B1"] = "Valore"
    _style_header(ws, 1)
    matched = sum(item.matched for item in result.files)
    reviewed = sum(item.review for item in result.files)
    rows = [
        ("Registro", result.registro_path),
        ("Righe del registro", result.registro_rows),
        ("Titoli confrontabili", result.works),
        ("Titoli scartati (senza parole utili)", result.skipped_works),
        ("Soglia", result.threshold),
        ("Scarto massimo fra primo e secondo titolo", result.gap),
        ("Fascia «Da rivedere» sotto la soglia", NEAR_BAND),
        ("Caratteri minimi", result.min_chars),
        ("Caratteri massimi", result.max_chars),
        ("Inventario precedente al", result.inventory_before or "nessun filtro"),
    ]
    if result.inventory_before:
        rows.extend([
            ("Esemplari esclusi (inventario successivo)", result.excluded_copies),
            ("Esemplari senza data di inventario", result.undated_copies),
        ])
    rows.extend([
        ("Confronti eseguiti", result.comparisons),
        ("Citazioni abbinate", matched),
        ("Titoli abbinati", result.titles_matched),
        ("Esemplari abbinati", result.copies_matched),
        ("Citazioni da rivedere", reviewed),
        ("Interrotto", "sì" if result.stopped else "no"),
    ])
    for offset, (label, value) in enumerate(rows, start=2):
        ws.cell(offset, 1, label)
        ws.cell(offset, 2, value)
        ws.cell(offset, 1).alignment = TOP
        ws.cell(offset, 2).alignment = Alignment(wrap_text=True, vertical="top")

    start = len(rows) + 3
    for column, header in enumerate(FILE_COLUMNS, start=1):
        ws.cell(start, column, header)
    _style_header(ws, start)
    for offset, stats in enumerate(result.files, start=1):
        unmatched = (
            stats.rows - stats.short - stats.long - stats.thin
            - stats.bland - stats.matched - stats.review
        )
        values = [
            stats.name, stats.rows, stats.short, stats.long, stats.thin,
            stats.bland, stats.matched, stats.review, unmatched,
        ]
        for column, value in enumerate(values, start=1):
            ws.cell(start + offset, column, value)
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 78
    for column in range(3, len(FILE_COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(column)].width = 20
    ws.freeze_panes = "A2"


def _write_workbook(path, result: MatchResult) -> None:
    wb = Workbook()
    libri = wb.active
    libri.title = "Libri"
    dettaglio = wb.create_sheet("Dettaglio")
    revisione = wb.create_sheet("Da rivedere")
    riepilogo = wb.create_sheet("Riepilogo")
    _write_table(libri, LIBRI_COLUMNS, result.libri, result.threshold)
    _write_table(dettaglio, DETTAGLIO_COLUMNS, result.dettaglio, result.threshold)
    _write_table(revisione, REVIEW_COLUMNS, result.review, result.threshold)
    _write_summary(riepilogo, result)
    for sheet, color in (
        (libri, "548235"),
        (dettaglio, "366092"),
        (revisione, "B26A00"),
        (riepilogo, "5B6777"),
    ):
        sheet.sheet_properties.tabColor = color
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def match_topografie(scraper_paths, registro_path, output_path, *,
                     threshold=DEFAULT_THRESHOLD, gap=DEFAULT_GAP,
                     min_chars=DEFAULT_MIN_CHARS, max_chars=DEFAULT_MAX_CHARS,
                     inventory_before: date | None = None,
                     log=print, progress=None, should_stop=None) -> dict:
    """
    Scrive l'Excel e ritorna un riepilogo. Il confronto è sequenziale.
    """
    threshold = float(threshold)
    gap = float(gap)
    min_chars = int(min_chars)
    max_chars = int(max_chars)
    if not 0 <= threshold <= 100:
        raise ValueError("La soglia deve essere tra 0 e 100.")
    if gap < 0:
        raise ValueError("Lo scarto non può essere negativo.")
    if min_chars < 0 or max_chars < 1 or min_chars > max_chars:
        raise ValueError("L'intervallo di caratteri non è valido.")

    scraper_paths = _dedupe(scraper_paths)
    if not scraper_paths:
        raise ValueError("Aggiungi almeno un Excel dello Scraper.")
    if not registro_path:
        raise ValueError("Manca il registro topografico.")
    if not output_path:
        raise ValueError("Manca il file di output.")
    output_path = str(output_path)
    if Path(output_path).suffix.lower() != ".xlsx":
        output_path = str(Path(output_path).with_suffix(".xlsx"))
    inputs = [registro_path, *scraper_paths]
    if any(_same_path(output_path, path) for path in inputs):
        raise ValueError(
            "Il file di output coincide con uno dei file in ingresso. Scegline un altro."
        )

    if isinstance(inventory_before, datetime):
        inventory_before = inventory_before.date()
    elif isinstance(inventory_before, str):
        inventory_before = parse_cutoff(inventory_before) if inventory_before.strip() else None
    works, registro_info = _load_registro(registro_path, log, inventory_before)
    files, citations = _load_scrapers(scraper_paths, min_chars, max_chars, log)
    doc_freq, postings = _index(works)
    log(f"Confronto di {len(citations)} citazioni con {len(works)} titoli…")

    accepted, review, comparisons, stopped = _compare(
        works, citations, doc_freq, postings, threshold, gap, log, progress, should_stop,
    )
    libri, dettaglio, review_rows, copies_matched, titles_matched = _build_tables(accepted, review)

    result = MatchResult(
        libri=libri,
        dettaglio=dettaglio,
        review=review_rows,
        files=files,
        registro_path=str(registro_path),
        registro_rows=registro_info["rows"],
        works=len(works),
        skipped_works=registro_info["skipped_works"],
        inventory_before="" if inventory_before is None else inventory_before.strftime("%d/%m/%Y"),
        excluded_copies=registro_info["excluded"],
        undated_copies=registro_info["undated"],
        copies_matched=copies_matched,
        titles_matched=titles_matched,
        comparisons=comparisons,
        threshold=threshold,
        gap=gap,
        min_chars=min_chars,
        max_chars=max_chars,
        stopped=stopped,
        output=output_path,
    )
    log("Scrittura del file…")
    _write_workbook(output_path, result)

    matched = sum(item.matched for item in files)
    reviewed = sum(item.review for item in files)
    if stopped:
        log(f"Interrotto. {matched} citazioni abbinate, {copies_matched} esemplari, "
            f"{reviewed} da rivedere → {output_path}")
    else:
        log(f"✅ {matched} citazioni abbinate, {copies_matched} esemplari, "
            f"{reviewed} da rivedere → {output_path}")
    log(f"Confronti eseguiti: {comparisons}.")
    return {
        "output": output_path,
        "matched": matched,
        "copies": copies_matched,
        "titles": titles_matched,
        "review": reviewed,
        "comparisons": comparisons,
        "stopped": stopped,
    }
