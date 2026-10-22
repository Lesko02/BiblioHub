"""
Genera il manuale italiano da bibliohub.guide.

    python docs/build_manual.py

Scrive docs/Manuale.md e docs/Manuale_BiblioHub.pdf.
Il carattere è Source Sans 3 (SIL Open Font License), scaricato se manca.
"""

import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from fpdf import FPDF

from bibliohub import APP_SUBTITLE, APP_VERSION
from bibliohub.guide import CHAPTERS, DOCS_DIR, MANUAL_MD, MANUAL_PDF

FONT_DIR = DOCS_DIR / "fonts"
REGULAR = FONT_DIR / "SourceSans3-Regular.ttf"
BOLD = FONT_DIR / "SourceSans3-Bold.ttf"
FONT_URLS = {
    REGULAR.name: "https://github.com/adobe-fonts/source-sans/raw/release/TTF/SourceSans3-Regular.ttf",
    BOLD.name: "https://github.com/adobe-fonts/source-sans/raw/release/TTF/SourceSans3-Bold.ttf",
}
LICENSE_URL = "https://raw.githubusercontent.com/adobe-fonts/source-sans/release/LICENSE.md"


def pdf_text(text):
    """La i cerchiata del pulsante non è nel carattere del PDF."""
    return text.replace("\u24d8", "i")


def ensure_fonts():
    FONT_DIR.mkdir(parents=True, exist_ok=True)
    license_path = FONT_DIR / "LICENSE.md"
    if not license_path.is_file():
        urllib.request.urlretrieve(LICENSE_URL, license_path)
    for name, url in FONT_URLS.items():
        path = FONT_DIR / name
        if path.is_file() and path.stat().st_size > 1000:
            continue
        urllib.request.urlretrieve(url, path)


NAVY = (31, 59, 99)
ACCENT = (54, 96, 146)
INK = (30, 38, 51)
MUTED = (91, 103, 119)
RULE = (213, 220, 230)
PALE = (228, 236, 247)


def write_markdown():
    lines = [
        "# Manuale d'uso di BiblioHub",
        "",
        f"{APP_SUBTITLE}. Versione {APP_VERSION}.",
        "",
        "Testo generato da `bibliohub/guide.py`. "
        "Per rigenerarlo: `python docs/build_manual.py`.",
        "",
        "Lo stesso testo si apre nell'applicazione con il pulsante **ⓘ Guida**.",
        "",
    ]
    for chapter in CHAPTERS:
        lines.append(f"## {chapter.title}")
        lines.append("")
        for block in chapter.blocks:
            if block.kind == "h2":
                lines.append(f"### {block.text}")
                lines.append("")
            elif block.kind == "p":
                lines.append(block.text)
                lines.append("")
            else:
                if block.text:
                    lines.append(f"### {block.text}")
                    lines.append("")
                for index, item in enumerate(block.items, 1):
                    mark = f"{index}." if block.kind == "steps" else "-"
                    lines.append(f"{mark} {item}")
                lines.append("")
    MANUAL_MD.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


class Manual(FPDF):
    def header(self):
        if self.page_no() == 1:
            return
        self.set_fill_color(*NAVY)
        self.rect(0, 0, 210, 12, "F")
        self.set_xy(16, 3)
        self.set_font("Sans", "", 9)
        self.set_text_color(255, 255, 255)
        self.cell(0, 6, "BiblioHub    ·    Manuale d'uso")
        self.set_y(20)

    def footer(self):
        if self.page_no() == 1:
            return
        self.set_y(-15)
        self.set_draw_color(*RULE)
        self.set_line_width(0.2)
        self.line(16, self.get_y(), 194, self.get_y())
        self.set_y(-13)
        self.set_font("Sans", "", 9)
        self.set_text_color(*MUTED)
        self.cell(90, 8, "Università degli Studi di Milano")
        self.cell(0, 8, str(self.page_no()), align="R")

    def cover(self):
        self.add_page()
        self.set_fill_color(*NAVY)
        self.rect(0, 0, 210, 297, "F")
        self.set_fill_color(*ACCENT)
        self.rect(0, 0, 8, 297, "F")

        self.set_xy(22, 78)
        self.set_font("Sans", "", 12)
        self.set_text_color(201, 214, 234)
        self.cell(0, 8, "UNIVERSITÀ DEGLI STUDI DI MILANO",
                  new_x="LMARGIN", new_y="NEXT")

        self.ln(10)
        self.set_x(22)
        self.set_font("Sans", "B", 36)
        self.set_text_color(255, 255, 255)
        self.cell(0, 16, "BiblioHub", new_x="LMARGIN", new_y="NEXT")

        self.set_x(22)
        self.set_font("Sans", "", 18)
        self.cell(0, 12, "Manuale d'uso", new_x="LMARGIN", new_y="NEXT")

        self.ln(7)
        self.set_fill_color(123, 163, 208)
        self.set_x(22)
        self.cell(36, 1.4, "", fill=True, new_x="LMARGIN", new_y="NEXT")

        self.ln(8)
        self.set_x(22)
        self.set_font("Sans", "", 13)
        self.set_text_color(232, 238, 246)
        self.multi_cell(
            150, 7,
            "Bibliografie d'esame, linee d'ordine Alma\ne attestazioni di pagamento.",
            new_x="LMARGIN", new_y="NEXT",
        )

        self.set_xy(22, 248)
        self.set_font("Sans", "", 12)
        self.set_text_color(201, 214, 234)
        self.cell(0, 7, APP_SUBTITLE, new_x="LMARGIN", new_y="NEXT")
        self.set_x(22)
        self.cell(0, 7, f"Versione {APP_VERSION}    ·    ottobre 2026",
                  new_x="LMARGIN", new_y="NEXT")

    def h1(self, text):
        self.set_fill_color(*PALE)
        self.set_text_color(*NAVY)
        self.set_font("Sans", "B", 18)
        y = self.get_y()
        self.rect(16, y, 178, 12, "F")
        self.set_xy(20, y + 2)
        self.cell(170, 8, pdf_text(text), new_x="LMARGIN", new_y="NEXT")
        self.ln(6)

    def h2(self, text):
        self.ln(2)
        self.ensure(18)
        self.start_section(pdf_text(text), level=1)
        self.set_text_color(*ACCENT)
        self.set_font("Sans", "B", 13)
        self.multi_cell(0, 7, pdf_text(text), new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def paragraph(self, text):
        self.set_text_color(*INK)
        self.set_font("Sans", "", 11)
        self.multi_cell(0, 5.8, pdf_text(text), new_x="LMARGIN", new_y="NEXT")
        self.ln(2.4)

    def label(self, text):
        self.ln(1)
        self.ensure(12)
        self.set_text_color(*NAVY)
        self.set_font("Sans", "B", 11)
        self.cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")

    def items(self, entries, numbered):
        for index, item in enumerate(entries, 1):
            self.ensure(12)
            mark = f"{index}." if numbered else "·"
            top = self.get_y()
            self.set_text_color(*ACCENT if numbered else NAVY)
            self.set_font("Sans", "B", 11)
            self.set_xy(self.l_margin, top)
            self.cell(8, 5.8, mark)
            self.set_xy(self.l_margin + 8, top)
            self.set_text_color(*INK)
            self.set_font("Sans", "", 11)
            self.multi_cell(self.epw - 8, 5.8, pdf_text(item), new_x="LMARGIN", new_y="NEXT")
            self.ln(1.2)

    def ensure(self, height):
        if self.get_y() + height > self.page_break_trigger:
            self.add_page()


def write_pdf():
    pdf = Manual(format="A4")
    pdf.set_lang("it-IT")
    pdf.set_title("Manuale d'uso di BiblioHub")
    pdf.set_author("BiblioHub")
    pdf.set_subject(APP_SUBTITLE)
    pdf.set_creator("BiblioHub docs/build_manual.py")
    pdf.set_keywords("BiblioHub, UNIMI, bibliografie, attestazioni, Alma")
    pdf.set_display_mode(zoom="fullpage", layout="continuous")
    pdf.set_auto_page_break(True, margin=18)
    pdf.set_margins(16, 20, 16)
    pdf.add_font("Sans", "", str(REGULAR))
    pdf.add_font("Sans", "B", str(BOLD))

    pdf.cover()

    pdf.add_page()
    pdf.start_section("Indice")
    pdf.h1("Indice")
    pdf.set_font("Sans", "", 12)
    for chapter in CHAPTERS:
        pdf.ensure(12)
        pdf.set_text_color(*ACCENT)
        pdf.cell(0, 9, pdf_text(chapter.title), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    pdf.set_text_color(*MUTED)
    pdf.set_font("Sans", "", 10)
    pdf.multi_cell(
        0, 5.5,
        "Nel lettore PDF i capitoli sono nei segnalibri, nello stesso ordine.",
        new_x="LMARGIN", new_y="NEXT",
    )

    for chapter in CHAPTERS:
        pdf.add_page()
        pdf.start_section(pdf_text(chapter.title))
        pdf.h1(chapter.title)
        for block in chapter.blocks:
            if block.kind == "h2":
                pdf.h2(block.text)
            elif block.kind == "p":
                pdf.paragraph(block.text)
            else:
                if block.text:
                    pdf.label(block.text)
                pdf.items(block.items, numbered=block.kind == "steps")
                pdf.ln(1.2)

    pdf.output(str(MANUAL_PDF))


def main():
    ensure_fonts()
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    write_markdown()
    write_pdf()
    print(f"Markdown: {MANUAL_MD}")
    print(f"PDF:      {MANUAL_PDF}")


if __name__ == "__main__":
    main()
