import sys
from dataclasses import dataclass, field


@dataclass(frozen=True)
class ToolInfo:
    key: str
    title: str
    version: str
    icon: str
    description: str
    page: str  # "modulo:Classe", importato solo alla prima apertura
    requires: list[str] = field(default_factory=list)


TOOLS = [
    ToolInfo(
        key="scraper",
        title="Scraper Insegnamenti",
        version="6.7",
        icon="🔎",
        description="Estrae docenti e bibliografia d'esame dagli elenchi "
                    "insegnamenti A-Z. Si scelgono anno accademico e ordinamento, "
                    "come sul sito. Include il merge originale.",
        page="bibliohub.pages.scraper:ScraperPage",
        requires=["pandas", "openpyxl", "requests", "bs4"],
    ),
    ToolInfo(
        key="po",
        title="PO Line → Excel/CSV",
        version="2.1",
        icon="📦",
        description="Converte l'export delle linee d'ordine (PO) di Alma in "
                    "CSV e/o Excel con titolo, autore, editore, anno, ISBN, "
                    "prezzo di listino, POL e date di invio.",
        page="bibliohub.pages.po_converter:POConverterPage",
        requires=["pandas", "openpyxl"],
    ),
    ToolInfo(
        key="attestazioni",
        title="Attestazioni di pagamento",
        version="3.1",
        icon="🧾",
        description="Genera le attestazioni di pagamento (Word + PDF) e il "
                    "testo email per fornitore e buono, a partire dalla "
                    "tabella fatture e dal template Word.",
        page="bibliohub.pages.attestazioni:AttestazioniPage",
        requires=["pandas", "openpyxl", "docxtpl", "tkcalendar"]
                 + (["win32com"] if sys.platform == "win32" else []),
    ),
    ToolInfo(
        key="merge",
        title="Merge bibliografie",
        version="1.1",
        icon="🔀",
        description="Confronta due estrazioni dello Scraper ignorando "
                    "maiuscole e minuscole. Mantiene Ordinamento, Periodo ed "
                    "Edizione e può mostrare solo le differenze oppure solo i testi in comune.",
        page="bibliohub.pages.merge_bibliografie:MergeBibliografiePage",
        requires=["pandas", "openpyxl"],
    ),
]
