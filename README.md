# BiblioHub

Desktop app for the library workflows of the University of Milan: scrape course bibliographies from the degree-course sites, convert Alma purchase-order lines, generate payment attestations, and compare two bibliography exports.

The user manual is in Italian: [docs/Manuale.md](docs/Manuale.md) and [docs/Manuale_BiblioHub.pdf](docs/Manuale_BiblioHub.pdf). In the app, **ⓘ Guida** at the top right opens the guide for the screen you are on, and can open the PDF.

## Run locally

Windows, Python 3.11 or newer. Microsoft Word is required only when you generate attestation PDFs.

```powershell
cd C:\Users\Lesko\Documents\BiblioHub
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
```

To start without a console, using the packages of that environment:

```powershell
.\.venv\Scripts\pythonw.exe BiblioHub.pyw
```

The first time you open a tool, BiblioHub offers to install any missing Python package. Folders and options are remembered in `%APPDATA%\BiblioHub\settings.json`.

## Avvio in locale

Windows, Python 3.11 o successivo. Microsoft Word serve solo per i PDF delle attestazioni.

```powershell
cd C:\Users\Lesko\Documents\BiblioHub
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
```

Senza finestra nera, con i pacchetti di quell'ambiente:

```powershell
.\.venv\Scripts\pythonw.exe BiblioHub.pyw
```

Alla prima apertura di uno strumento, BiblioHub può installare i pacchetti mancanti. Percorsi e opzioni restano in `%APPDATA%\BiblioHub\settings.json`.

Nell'app, **ⓘ Guida** in alto a destra spiega la schermata aperta. Il manuale completo è `docs/Manuale_BiblioHub.pdf`.

## Tools

| Tool | Version | Role |
| --- | --- | --- |
| Scraper Insegnamenti | 6.5 | Course bibliographies from the A–Z list, with academic year and ordinamento |
| PO Line → Excel/CSV | 2.1 | Alma PO lines to CSV and Excel, list price restored |
| Attestazioni di pagamento | 3.1 | Word, PDF and the accompanying email, grouped by supplier and voucher |
| Merge bibliografie | 1.1 | Case-insensitive compare of two scraper workbooks |

Application version: 1.1.0.

## Layout

```
main.py, BiblioHub.pyw     start the app
bibliohub/app.py           window, navigation, package install
bibliohub/registry.py      tools, versions, dependencies
bibliohub/pages/           screens
bibliohub/tools/           logic, independent of the GUI
bibliohub/guide.py         Italian help text
docs/                      Italian manual (Markdown and PDF)
```

After editing `bibliohub/guide.py`, regenerate the manual from the project folder:

```powershell
python docs/build_manual.py
```
