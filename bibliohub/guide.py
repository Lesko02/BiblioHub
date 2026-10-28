"""Testi della guida in italiano, usati dalla finestra ⓘ e dal manuale PDF."""

from pathlib import Path

DOCS_DIR = Path(__file__).resolve().parents[1] / "docs"
MANUAL_PDF = DOCS_DIR / "Manuale_BiblioHub.pdf"
MANUAL_MD = DOCS_DIR / "Manuale.md"


class Block:
    def __init__(self, kind, text="", items=()):
        self.kind = kind
        self.text = text
        self.items = tuple(items)


class Chapter:
    def __init__(self, key, menu, title, blocks):
        self.key = key
        self.menu = menu
        self.title = title
        self.blocks = tuple(blocks)


def P(text):
    return Block("p", text)


def H(text):
    return Block("h2", text)


def STEPS(*items):
    return Block("steps", "Come si usa", items)


def EXPECT(*items):
    return Block("expect", "Cosa aspettarsi", items)


def NOTES(*items):
    return Block("notes", "Da tenere presente", items)


def L(*items):
    return Block("list", "", items)


CHAPTERS = [
    Chapter(
        "home",
        "Panoramica",
        "Panoramica",
        [
            P("BiblioHub riunisce in una finestra gli strumenti per le bibliografie "
              "d'esame, gli ordini Alma e le attestazioni di pagamento "
              "dell'Università degli Studi di Milano."),
            P("Dalla home si apre uno strumento. Home, in alto a sinistra, riporta "
              "alla scelta. Il pulsante i con la scritta Guida, in alto a destra, apre questa guida "
              "sulla schermata in cui ti trovi. Da qui si apre anche il manuale PDF."),
            H("Durante un lavoro"),
            P("In basso vedi lo stato, la barra di avanzamento e il registro delle "
              "attività. A lavoro finito, Apri risultato apre il file o la cartella "
              "prodotti. Se l'operazione si può fermare, compare Interrompi."),
            P("Le ultime cartelle e le opzioni restano in "
              "%APPDATA%\\BiblioHub\\settings.json. Alla prima apertura di uno "
              "strumento, se manca un pacchetto Python, BiblioHub propone di installarlo."),
            H("I quattro strumenti"),
            L(
                "Scraper Insegnamenti 6.5 — bibliografie dal sito dei corsi di laurea.",
                "PO Line → Excel/CSV 2.1 — linee d'ordine Alma in CSV ed Excel.",
                "Attestazioni di pagamento 3.1 — Word, PDF e testo dell'email.",
                "Merge bibliografie 1.1 — confronto tra due estrazioni dello scraper.",
            ),
        ],
    ),
    Chapter(
        "scraper",
        "Scraper Insegnamenti",
        "Scraper Insegnamenti",
        [
            P("Legge la pagina «Elenco insegnamenti A-Z» di un corso di laurea UNIMI "
              "e scrive un Excel: una riga per ogni risorsa bibliografica, con docente, "
              "edizione, periodo e ordinamento."),
            H("Scheda Scraping"),
            STEPS(
                "Incolla l'URL della pagina elenco. Finisce di solito con "
                "elenco-insegnamenti-z. Esempio: la pagina di Informatica o di Mediazione.",
                "Scegli il file Excel di output, oppure lascialo vuoto. In quel caso "
                "il file si chiama Elenco_CORSO_updated.xlsx e finisce nell'ultima cartella usata.",
                "Premi Carica elenco. Si riempiono Anno accademico e Ordinamento, "
                "gli stessi menu del sito.",
                "Scegli l'anno. Gli ordinamenti di quell'anno vengono ricaricati.",
                "Scegli un ordinamento, oppure Tutti gli ordinamenti per scaricare ogni tabella.",
                "Premi Avvia. Nel registro compare il corso in lavorazione. "
                "Interrompi salva le righe già raccolte.",
                "Apri risultato apre l'Excel.",
            ),
            EXPECT(
                "Colonne: Insegnamento, Ordinamento, Periodo, Edizione, Docente, "
                "Descrizione, Data, ISBN.",
                "Ordinamento è il titolo della tabella sul sito, senza la parola "
                "Immatricolati. Per esempio: a.a. 2026/27 (FAA).",
                "Ogni a capo della pagina, quello del tag br, resta una riga. "
                "Dentro la stessa riga, a) a- e 1) 1- 1. (fino a 10) vengono divisi ancora.",
                "Data è l'ultimo anno tra il 1900 e il 2099 della citazione, di solito "
                "l'anno di pubblicazione in fondo. «dal 1943 al 2003, Bologna, 2006» "
                "resta 2006. Un anno come (1797) accanto all'autore resta fuori.",
                "«a) … 2015; b) … 2017» e «1. … 2.» sulla stessa riga producono una riga per voce. "
                "«a. … b.» resta una sola riga: il punto dopo la lettera può essere un autore. "
                "Due libri separati solo da un a capo restano due righe anche senza elenco.",
                "Un'edizione non attiva è una riga con Descrizione «Edizione non attiva», "
                "in giallo e in grassetto.",
                "Le celle sono testo: i numeri non escono con la coda .0.",
            ),
            NOTES(
                "Senza Carica elenco, Avvia usa l'anno già aperto nella pagina e tutti gli ordinamenti.",
                "Serve la rete. Un corso il cui indirizzo non risponde viene saltato e segnalato nel registro.",
                "Lo stesso indirizzo di scheda, se compare in più ordinamenti, viene letto una sola volta.",
                "Chiudi l'Excel di destinazione prima di rilanciare: un file aperto non si può sovrascrivere.",
            ),
            H("Scheda Merge, in questo strumento"),
            P("La seconda scheda confronta due Excel e tiene solo le righe diverse. "
              "Il confronto distingue maiuscole e minuscole. Il risultato ha Origine "
              "Left o Right e omette Periodo ed Edizione."),
            P("Per il confronto tra un anno e l'altro usa lo strumento Merge bibliografie: "
              "ignora le maiuscole. L'ordinamento resta in colonna e non entra nel confronto."),
            NOTES(
                "File LEFT e file RIGHT sono i due Excel. Se sono identici, il registro "
                "lo dice e non viene scritto un file.",
            ),
        ],
    ),
    Chapter(
        "po",
        "PO Line → Excel/CSV",
        "PO Line → Excel/CSV",
        [
            P("Converte l'export Excel delle linee d'ordine Alma in un elenco con titolo, "
              "autore, editore, anno, ISBN, prezzo di listino, POL, PO, data di invio "
              "e note di ricezione."),
            STEPS(
                "Seleziona l'Excel delle linee d'ordine esportato da Alma.",
                "Indica il file di output, oppure lascialo vuoto: il risultato si chiama "
                "NOME_estratto e resta nella stessa cartella.",
                "Scegli CSV, Excel, oppure Entrambi.",
                "Premi Converti. Apri risultato apre il file, o la cartella se ne hai chiesti due.",
            ),
            EXPECT(
                "Colonne: Titolo, Autore, Editore, Anno, ISBN, Prezzo, POL, PO, "
                "Data di Invio, Note di Ricezione.",
                "Titolo, autore, editore e anno escono dalla descrizione della copia.",
                "Il prezzo in uscita è il listino: allo sconto del 15% presente in Alma "
                "viene riapplicato il fattore 1,1764705882, arrotondato al centesimo, "
                "per esempio € 15,00.",
                "Il CSV usa il separatore punto e virgola e la codifica UTF-8 con BOM, "
                "così Excel lo apre con le colonne già divise.",
                "Con Entrambi ottieni due file con lo stesso nome, uno .csv e uno .xlsx.",
            ),
            NOTES(
                "L'Excel di Alma deve avere le colonne della descrizione copia, del numero "
                "standard, del riferimento di linea, di PO, del prezzo e della data di invio.",
                "La nota di ricezione, se la colonna c'è, finisce in Note di Ricezione.",
            ),
        ],
    ),
    Chapter(
        "attestazioni",
        "Attestazioni di pagamento",
        "Attestazioni di pagamento",
        [
            P("Dalla tabella delle fatture produce, per ogni fornitore e buono, "
              "l'attestazione in Word, il PDF e un file di testo con l'email di accompagnamento."),
            STEPS(
                "Seleziona la tabella Excel. Il modello Word è già impostato sul template "
                "incluso in BiblioHub: cambialo solo se usi un altro modello.",
                "Scegli la cartella di output. Se non esiste, viene creata.",
                "Imposta l'intervallo di date. Il filtro è la colonna data_protocollo. "
                "Mese scorso e Mese corrente compilano le date da soli.",
                "Lascia attiva Genera anche i PDF se Microsoft Word è installato.",
                "Premi Verifica dati per vedere quanti documenti usciranno, fornitore per fornitore, "
                "con il numero di fatture e il totale.",
                "Premi Genera attestazioni. Interrompi ferma i documenti ancora da fare: "
                "quelli già scritti restano nella cartella.",
                "Apri cartella output, oppure Apri risultato a fine lavoro.",
            ),
            EXPECT(
                "Per ogni coppia fornitore e numero buono escono tre file: "
                "Attestazione_FORNITORE_buonoNUMERO_AAAAMMGG.docx, "
                "lo stesso nome in PDF, e lo stesso nome con _mail.txt.",
                "La data nel nome è la data_protocollo più vecchia del gruppo.",
                "Il PDF viene creato da Word in background. Senza la casella PDF restano Word e testo.",
                "I numeri interi letti da Excel, come un buono finito in .0, escono senza decimali.",
            ),
            NOTES(
                "La tabella deve avere le colonne: fornitore, number, Sistema, identificativo, "
                "data_protocollo, totale, buoni_di_consegna, numero_buono, data_buono, CIG, pagamento.",
                "L'intervallo include il primo e l'ultimo giorno. La data di inizio deve precedere "
                "o coincidere con la data di fine.",
                "Chiudi Word e i file della cartella di output se un salvataggio dice che il file è aperto.",
                "Su un computer senza Word la casella PDF è spenta: restano i documenti Word e le email.",
            ),
        ],
    ),
    Chapter(
        "merge",
        "Merge bibliografie",
        "Merge bibliografie",
        [
            P("Confronta due Excel prodotti dallo Scraper. Guarda il testo della risorsa, "
              "non i colori, e tratta «Rossi» e «rossi» come lo stesso valore. "
              "L'ordinamento (anno accademico) resta visibile e non entra nel confronto. "
              "Periodo ed Edizione distinguono due adozioni diverse."),
            STEPS(
                "In Estrazione precedente carica l'Excel dell'anno vecchio. "
                "Nel risultato le righe solo di questo file sono Left.",
                "In Estrazione corrente carica l'Excel dell'anno nuovo. "
                "Le righe solo di questo file sono Right.",
                "Scegli il file di output. L'estensione .xlsx viene aggiunta se manca.",
                "Solo differenze elenca i testi presenti in una sola estrazione. "
                "Solo in comune elenca i testi presenti in entrambe.",
                "La casella sulle descrizioni brevi, attiva di serie, scarta le descrizioni "
                "con meno di 3 caratteri. Puoi alzare o abbassare la soglia, o spegnere il filtro.",
                "Premi Confronta, poi Apri risultato.",
            ),
            EXPECT(
                "Colonne: Insegnamento, Ordinamento, Periodo, Edizione, Docente, "
                "Descrizione, Data, ISBN, Origine.",
                "Solo differenze: Origine Left (sfondo pesca) o Right (sfondo verde).",
                "Solo in comune: Origine «In comune» (sfondo azzurro). Il testo è quello "
                "del file corrente.",
                "Il nome dell'insegnamento, se i due file differiscono solo per le maiuscole, "
                "è quello del file corrente.",
                "Il file di output viene sempre scritto, anche quando non ci sono righe da segnalare.",
            ),
            NOTES(
                "Servono le colonne dello Scraper: Insegnamento, Periodo, Edizione, Docente, "
                "Descrizione, Data, ISBN. Se manca Ordinamento, il confronto prosegue e la "
                "colonna resta vuota.",
                "Due righe sono la stessa risorsa quando coincidono, a meno delle maiuscole, "
                "Periodo, Edizione, Docente, Descrizione, Data e ISBN. "
                "Un ordinamento diverso non le rende diverse.",
                "La scheda Merge dentro lo Scraper è il confronto precedente: distingue le "
                "maiuscole e non riporta Periodo ed Edizione.",
            ),
        ],
    ),
    Chapter(
        "trouble",
        "Se qualcosa non va",
        "Se qualcosa non va",
        [
            H("Lo scraper non parte"),
            P("L'URL deve iniziare con http. Se il registro si ferma su «Nessun corso trovato», "
              "controlla di aver aperto proprio la pagina elenco A-Z e, se hai filtrato, "
              "che l'ordinamento esista per l'anno scelto."),
            H("Data vuota o anno inatteso"),
            P("Data resta vuota quando nella citazione non c'è un anno tra il 1900 e il 2099. "
              "Se l'anno è quello del titolo e più avanti c'è l'anno di stampa, in colonna "
              "finisce l'anno di stampa. Due citazioni sulla stessa riga, senza un elenco "
              "a) a- o 1) 1- 1., restano una sola riga. Un a capo della pagina le divide comunque."),
            H("File aperto"),
            P("Excel o Word tengono il file bloccato. Chiudi il documento e rilancia. "
              "Il messaggio indica il percorso che non si è potuto scrivere."),
            H("PDF delle attestazioni"),
            P("La conversione usa Microsoft Word, non una stampante PDF. Word deve essere "
              "installato sullo stesso computer. Se l'installazione di un pacchetto chiede "
              "di riavviare BiblioHub, chiudi e riapri l'applicazione."),
            H("Merge con errore di colonne"),
            P("Merge bibliografie accetta gli Excel dello Scraper. Un file con lo stesso "
              "numero di colonne ma con intestazioni diverse viene rifiutato: il registro "
              "elenca le colonne mancanti."),
        ],
    ),
    Chapter(
        "code",
        "Com'è fatto il programma",
        "Com'è fatto il programma",
        [
            P("BiblioHub è un'applicazione desktop in Python e tkinter. "
              "La logica di ogni strumento sta in bibliohub/tools e non dipende dalla finestra. "
              "Le schermate stanno in bibliohub/pages. Il registro degli strumenti, con versione "
              "e pacchetti richiesti, è bibliohub/registry.py."),
            H("File principali"),
            L(
                "main.py e BiblioHub.pyw — avvio, il secondo senza finestra nera.",
                "bibliohub/app.py — finestra, home, installazione dei pacchetti mancanti.",
                "bibliohub/widgets.py — campi file, registro attività, lavoro in background.",
                "bibliohub/guide.py — testi di questa guida. Il PDF si rigenera da qui.",
                "bibliohub/values.py — i numeri interi letti da Excel restano senza .0.",
                "bibliohub/settings.py — preferenze in %APPDATA%\\BiblioHub\\settings.json.",
                "bibliohub/assets/Template.docx — modello Word delle attestazioni.",
                "docs/Manuale.md e docs/Manuale_BiblioHub.pdf — questo manuale.",
            ),
            H("Versioni attuali"),
            L(),
            P("Dopo una modifica a bibliohub/guide.py il manuale si rigenera, dalla cartella "
              "del progetto, con: python docs/build_manual.py"),
        ],
    ),
]


def chapter(key):
    for item in CHAPTERS:
        if item.key == key:
            return item
    return CHAPTERS[0]


def _fill_versions():
    from . import APP_VERSION
    from .registry import TOOLS

    blurbs = {
        "scraper": "bibliografie dal sito dei corsi di laurea.",
        "po": "linee d'ordine Alma in CSV ed Excel.",
        "attestazioni": "Word, PDF e testo dell'email.",
        "merge": "confronto tra due estrazioni dello scraper.",
    }
    for block in chapter("home").blocks:
        if block.kind == "list" and block.items and block.items[0].startswith("Scraper"):
            block.items = tuple(
                f"{tool.title} {tool.version} — {blurbs[tool.key]}" for tool in TOOLS
            )
    for block in chapter("code").blocks:
        if block.kind == "list" and not block.items:
            block.items = (
                f"Applicazione {APP_VERSION}.",
                *(f"{tool.title} {tool.version}." for tool in TOOLS),
            )


_fill_versions()
