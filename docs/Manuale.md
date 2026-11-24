# Manuale d'uso di BiblioHub

Strumenti per i sistemi bibliotecari UNIMI. Versione 1.1.0.

Testo generato da `bibliohub/guide.py`. Per rigenerarlo: `python docs/build_manual.py`.

Lo stesso testo si apre nell'applicazione con il pulsante **ⓘ Guida**.

## Panoramica

BiblioHub riunisce in una finestra gli strumenti per le bibliografie d'esame, il confronto con il registro topografico, gli ordini Alma e le attestazioni di pagamento dell'Università degli Studi di Milano.

Dalla home si apre uno strumento. Home, in alto a sinistra, riporta alla scelta. Il pulsante i con la scritta Guida, in alto a destra, apre questa guida sulla schermata in cui ti trovi. Da qui si apre anche il manuale PDF.

### Durante un lavoro

In basso vedi lo stato, la barra di avanzamento e il registro delle attività. A lavoro finito, Apri risultato apre il file o la cartella prodotti. Se l'operazione si può fermare, compare Interrompi.

Le ultime cartelle e le opzioni restano in %APPDATA%\BiblioHub\settings.json. Alla prima apertura di uno strumento, se manca un pacchetto Python, BiblioHub propone di installarlo.

### Gli strumenti

- Scraper Insegnamenti 7.0 — bibliografie dal sito dei corsi di laurea.
- PO Line → Excel/CSV 2.1 — linee d'ordine Alma in CSV ed Excel.
- Attestazioni di pagamento 3.1 — Word, PDF e testo dell'email.
- Merge bibliografie 1.2 — confronto tra due estrazioni dello scraper.
- Topografie 1.1 — abbinamento fra bibliografie estratte e registro dei libri d'esame.

## Scraper Insegnamenti

Legge la pagina «Elenco insegnamenti A-Z» di un corso di laurea UNIMI e scrive un Excel: una riga per ogni risorsa bibliografica, con docente, edizione, periodo e ordinamento.

### Scheda Scraping

### Come si usa

1. Incolla l'URL della pagina elenco. Finisce di solito con elenco-insegnamenti-z. Esempio: la pagina di Informatica o di Mediazione.
2. Scegli il file Excel di output, oppure lascialo vuoto. In quel caso il file si chiama Elenco_CORSO_ANNO_updated.xlsx, per esempio Elenco_informatica_2026-2027_updated.xlsx, e finisce nell'ultima cartella usata. Con Primo o Secondo semestre il nome lo riporta: Elenco_informatica_2026-2027_Primo_semestre_updated.xlsx.
3. Premi Carica elenco. Si riempiono Anno accademico e Ordinamento, gli stessi menu del sito.
4. Scegli l'anno. Gli ordinamenti di quell'anno vengono ricaricati.
5. Scegli un ordinamento, oppure Tutti gli ordinamenti per scaricare ogni tabella.
6. Scegli il periodo: Primo semestre, Secondo semestre, Annuali o Qualsiasi. Conta la colonna Periodo di distribuzione dell'elenco.
7. Premi Avvia. Nel registro compare il corso in lavorazione. Interrompi salva le righe già raccolte.
8. Apri risultato apre l'Excel.

### Cosa aspettarsi

- Colonne: Insegnamento, Ordinamento, Periodo, Edizione, Docente, Descrizione, Data, ISBN.
- Ordinamento è il titolo della tabella sul sito, senza la parola Immatricolati. Per esempio: a.a. 2026/27 (FAA).
- Ogni a capo della pagina, quello del tag br, resta una riga. Dentro la stessa riga, a) a- e 1) 1- 1. (fino a 10) vengono divisi ancora.
- Data è l'ultimo anno tra il 1900 e il 2099 della citazione, di solito l'anno di pubblicazione in fondo. «dal 1943 al 2003, Bologna, 2006» resta 2006. Un anno come (1797) accanto all'autore resta fuori.
- «a) … 2015; b) … 2017» e «1. … 2.» sulla stessa riga producono una riga per voce. «a. … b.» resta una sola riga: il punto dopo la lettera può essere un autore. Due libri separati solo da un a capo restano due righe anche senza elenco.
- Un'edizione non attiva è una riga con Descrizione «Edizione non attiva», in giallo e in grassetto.
- Le celle sono testo: i numeri non escono con la coda .0.

### Da tenere presente

- Senza Carica elenco, Avvia usa l'anno già aperto nella pagina e tutti gli ordinamenti. Il nome del file prende comunque quell'anno.
- Il nome automatico si aggiorna quando cambiano anno o periodo. Un nome scritto a mano o scelto con Sfoglia resta quello.
- Con Primo semestre, Secondo semestre o Annuali i corsi degli altri periodi non vengono aperti: lo scraping è più breve. Annuali comprende i corsi annuali e quelli su più periodi. I corsi con periodo non definito o senza periodo compaiono solo con Qualsiasi.
- Serve la rete. Un corso il cui indirizzo non risponde viene saltato e segnalato nel registro.
- Lo stesso indirizzo di scheda, se compare in più ordinamenti, viene letto una sola volta.
- Le schede si scaricano dieci alla volta, con una pausa di 0,15 secondi dopo ogni scheda. L'Excel resta nello stesso ordine dell'elenco. Interrompi salva solo i corsi già scritti e annulla le schede ancora in coda.
- Chiudi l'Excel di destinazione prima di rilanciare: un file aperto non si può sovrascrivere.

### Scheda Merge, in questo strumento

La seconda scheda confronta due Excel e tiene solo le righe diverse. Il confronto distingue maiuscole e minuscole. Il risultato ha Origine Left o Right e omette Periodo ed Edizione.

Per il confronto tra un anno e l'altro usa lo strumento Merge bibliografie: ignora le maiuscole. L'ordinamento resta in colonna e non entra nel confronto.

### Da tenere presente

- File LEFT e file RIGHT sono i due Excel. Se sono identici, il registro lo dice e non viene scritto un file.

## PO Line → Excel/CSV

Converte l'export Excel delle linee d'ordine Alma in un elenco con titolo, autore, editore, anno, ISBN, prezzo di listino, POL, PO, data di invio e note di ricezione.

### Come si usa

1. Seleziona l'Excel delle linee d'ordine esportato da Alma.
2. Indica il file di output, oppure lascialo vuoto: il risultato si chiama NOME_estratto e resta nella stessa cartella.
3. Scegli CSV, Excel, oppure Entrambi.
4. Premi Converti. Apri risultato apre il file, o la cartella se ne hai chiesti due.

### Cosa aspettarsi

- Colonne: Titolo, Autore, Editore, Anno, ISBN, Prezzo, POL, PO, Data di Invio, Note di Ricezione.
- Titolo, autore, editore e anno escono dalla descrizione della copia.
- Il prezzo in uscita è il listino: allo sconto del 15% presente in Alma viene riapplicato il fattore 1,1764705882, arrotondato al centesimo, per esempio € 15,00.
- Il CSV usa il separatore punto e virgola e la codifica UTF-8 con BOM, così Excel lo apre con le colonne già divise.
- Con Entrambi ottieni due file con lo stesso nome, uno .csv e uno .xlsx.

### Da tenere presente

- L'Excel di Alma deve avere le colonne della descrizione copia, del numero standard, del riferimento di linea, di PO, del prezzo e della data di invio.
- La nota di ricezione, se la colonna c'è, finisce in Note di Ricezione.

## Attestazioni di pagamento

Dalla tabella delle fatture produce, per ogni fornitore e buono, l'attestazione in Word, il PDF e un file di testo con l'email di accompagnamento.

### Come si usa

1. Seleziona la tabella Excel. Il modello Word è già impostato sul template incluso in BiblioHub: cambialo solo se usi un altro modello.
2. Scegli la cartella di output. Se non esiste, viene creata.
3. Imposta l'intervallo di date. Il filtro è la colonna data_protocollo. Mese scorso e Mese corrente compilano le date da soli.
4. Lascia attiva Genera anche i PDF se Microsoft Word è installato.
5. Premi Verifica dati per vedere quanti documenti usciranno, fornitore per fornitore, con il numero di fatture e il totale.
6. Premi Genera attestazioni. Interrompi ferma i documenti ancora da fare: quelli già scritti restano nella cartella.
7. Apri cartella output, oppure Apri risultato a fine lavoro.

### Cosa aspettarsi

- Per ogni coppia fornitore e numero buono escono tre file: Attestazione_FORNITORE_buonoNUMERO_AAAAMMGG.docx, lo stesso nome in PDF, e lo stesso nome con _mail.txt.
- La data nel nome è la data_protocollo più vecchia del gruppo.
- Il PDF viene creato da Word in background. Senza la casella PDF restano Word e testo.
- I numeri interi letti da Excel, come un buono finito in .0, escono senza decimali.

### Da tenere presente

- La tabella deve avere le colonne: fornitore, number, Sistema, identificativo, data_protocollo, totale, buoni_di_consegna, numero_buono, data_buono, CIG, pagamento.
- L'intervallo include il primo e l'ultimo giorno. La data di inizio deve precedere o coincidere con la data di fine.
- Chiudi Word e i file della cartella di output se un salvataggio dice che il file è aperto.
- Su un computer senza Word la casella PDF è spenta: restano i documenti Word e le email.

## Merge bibliografie

Confronta due Excel prodotti dallo Scraper. Guarda il testo della risorsa, non i colori, e tratta «Rossi» e «rossi» come lo stesso valore. L'ordinamento (anno accademico) resta visibile e non entra nel confronto. Periodo ed Edizione distinguono due adozioni diverse.

### Come si usa

1. In Estrazione precedente carica l'Excel dell'anno vecchio. Nel risultato le righe solo di questo file sono Left.
2. In Estrazione corrente carica l'Excel dell'anno nuovo. Le righe solo di questo file sono Right.
3. Scegli il file di output. L'estensione .xlsx viene aggiunta se manca.
4. Solo differenze elenca i testi presenti in una sola estrazione. Solo in comune elenca i testi presenti in entrambe.
5. La casella sulle descrizioni brevi, attiva di serie, scarta le descrizioni con meno di 3 caratteri. Puoi alzare o abbassare la soglia, o spegnere il filtro.
6. Premi Confronta, poi Apri risultato.

### Cosa aspettarsi

- Colonne: Insegnamento, Ordinamento, Periodo, Edizione, Docente, Descrizione, Data, ISBN, Origine.
- Solo differenze: Origine Left (sfondo pesca) o Right (sfondo verde).
- Solo in comune: Origine «In comune» (sfondo azzurro). Il testo è quello del file corrente.
- Il nome dell'insegnamento, se i due file differiscono solo per le maiuscole, è quello del file corrente.
- Il file di output viene sempre scritto, anche quando non ci sono righe da segnalare.

### Da tenere presente

- Servono le colonne dello Scraper: Insegnamento, Periodo, Edizione, Docente, Descrizione, Data, ISBN. Se manca Ordinamento, il confronto prosegue e la colonna resta vuota.
- Due righe sono la stessa risorsa quando coincidono, a meno delle maiuscole, Periodo, Edizione, Docente, Descrizione, Data e ISBN. Un ordinamento diverso non le rende diverse.
- La scheda Merge dentro lo Scraper è il confronto precedente: distingue le maiuscole e non riporta Periodo ed Edizione.

## Topografie

Confronta uno o più Excel dello Scraper con un registro di esemplari già segnati come libri d'esame. Il registro è un export di Alma: ogni riga è una copia e il titolo sta in Title (Complete), di solito nella forma «titolo : sottotitolo / autore».

La citazione estratta non ha una forma fissa. Il punteggio, da 0 a 100, dice quanto il titolo del registro è contenuto in quella riga. Editore, città e note in più non lo abbassano: lo abbassano le parole del titolo che mancano. Ogni citazione è misurata solo contro i titoli che condividono una sua parola poco comune. Il confronto scorre in un solo filo.

### Il punteggio

Il titolo del registro si divide sulla barra: a sinistra le parole del titolo, a destra l'autore. Maiuscole, accenti e punteggiatura non contano. Si ignorano le parole più corte di tre lettere, i numeri e le parole vuote (il, di, the, of).

Ogni parola del titolo pesa in base a quanto è rara nel registro: «adversaire» pesa più di «storia». La copertura è la quota di questo peso ritrovata nella citazione. Se dopo la barra c'è un autore, il punteggio è 82% di copertura più 18% di parole dell'autore ritrovate. Senza autore, è solo la copertura.

Una sola parola trovata, se è corta o molto comune, vale il 45% del punteggio. Una parola lunga con una lettera di differenza conta lo stesso (Mahfuz e Mahfouz). L'iniziale «U.» vale Umberto solo se «Eco» c'è già. Anno uguale: +4 punti; un anno di differenza: +2. Un anno più lontano non toglie punti, perché in bibliografia c'è spesso l'anno dell'opera e a scaffale una ristampa. Il tetto è 100.

### Soglia e scarto

Di ogni citazione restano il titolo col punteggio più alto e il secondo. Lo scarto è la differenza fra i due. In Libri entra la citazione che sta sopra la soglia (di serie 80), con uno scarto maggiore di quello impostato (di serie 8), e col cognome nella citazione quando il titolo ha meno di cinque parole. Il cognome è la parola più lunga dell'autore.

Le altre citazioni vicine finiscono in Da rivedere. Scarto basso: il secondo titolo è ancora entro gli 8 punti, come due livelli dello stesso manuale. Sotto soglia: il punteggio arriva fino a 10 punti sotto la soglia. Autore assente: il titolo è breve e il cognome non c'è, per esempio lo stesso titolo firmato da un'altra persona, o «giornale» dentro «articolo di giornale». Sotto i 70 punti la citazione non viene scritta.

### Come si usa

1. In Estrazioni dello Scraper aggiungi uno o più Excel. Servono le colonne Insegnamento e Descrizione. Ordinamento, Periodo, Edizione, Docente, Data e ISBN, se ci sono, finiscono nel risultato.
2. In Registro topografico carica l'export degli esemplari. La colonna Title (Complete) è obbligatoria. Barcode, inventario, collocazione, biblioteca, ubicazione, MMS Id e Dewey vengono ricopiati quando l'intestazione c'è.
3. Di serie si confronta tutto il registro. Con la casella sull'inventario restano solo le copie inventariate prima della data: le aggiunte più recenti escono dal confronto. Un anno da solo, per esempio 2024, vale il 1° gennaio 2024. La copia inventariata proprio quel giorno resta fuori. Le righe senza data restano.
4. Scegli il file di output. Non può essere uno dei file appena caricati.
5. La soglia, di serie 80, decide quali abbinamenti vanno nel foglio Libri. Lo scarto, di serie 8, manda in Da rivedere i casi in cui il secondo titolo è ancora vicino al primo.
6. Le citazioni più corte del minimo, o più lunghe del massimo, non vengono confrontate. Di serie il minimo è 25 caratteri e il massimo 400.
7. Premi Abbina, poi Apri risultato. Interrompi ferma il confronto e salva quanto è già stato abbinato.

### Cosa aspettarsi

- Libri: un esemplare per riga, con barcode, inventario e collocazione, e i corsi in cui la citazione è stata riconosciuta. Il riferimento indica il file, la riga dell'Excel e l'insegnamento.
- Dettaglio: una riga per ogni citazione abbinata e ogni copia, con il punteggio.
- Da rivedere: citazioni sotto la soglia ma vicine (fino a 10 punti in meno), citazioni in cui i due titoli migliori sono troppo vicini, e titoli brevi in cui il cognome non compare nella citazione. Il motivo è Sotto soglia (sfondo pesca), Scarto basso o Autore assente (sfondo giallo).
- Riepilogo: quante righe sono state lette, scartate, abbinate o messe da parte, file per file. Con il filtro sull'inventario indica anche gli esemplari esclusi.
- Lo stesso titolo presente in più copie produce una riga per copia, con lo stesso punteggio.
- Il verde in colonna Punteggio è da 90 in su, l'azzurro è sopra la soglia, il pesca è sotto.

### Da tenere presente

- Un trattino vale come uno spazio, e due parole attaccate nell'altra fonte si riconoscono lo stesso.
- In Libri il punteggio di un esemplare è il più alto tra le citazioni che l'hanno riconosciuto. In Dettaglio lo scarto è quello della citazione, uguale per tutte le copie di quel titolo. In Da rivedere lo scarto è la differenza fra Punteggio e Secondo punteggio.
- Un titolo del registro senza parole utilizzabili, per esempio solo «vol. 1», non entra nel confronto. Un titolo di una sola parola resta, se quella parola è lunga e poco comune.
- Il filtro sull'inventario usa la colonna Inventory Date (calendar). Senza quella colonna, con la casella attiva, il confronto non parte.
- Si legge il primo foglio di ogni file. La colonna Riga del risultato è la riga di quel foglio, intestazione compresa.

## Se qualcosa non va

### Lo scraper non parte

L'URL deve iniziare con http. Se il registro si ferma su «Nessun corso trovato», controlla di aver aperto proprio la pagina elenco A-Z e, se hai filtrato, che l'ordinamento esista per l'anno scelto.

### Data vuota o anno inatteso

Data resta vuota quando nella citazione non c'è un anno tra il 1900 e il 2099. Se l'anno è quello del titolo e più avanti c'è l'anno di stampa, in colonna finisce l'anno di stampa. Due citazioni sulla stessa riga, senza un elenco a) a- o 1) 1- 1., restano una sola riga. Un a capo della pagina le divide comunque.

### File aperto

Excel o Word tengono il file bloccato. Chiudi il documento e rilancia. Il messaggio indica il percorso che non si è potuto scrivere.

### PDF delle attestazioni

La conversione usa Microsoft Word, non una stampante PDF. Word deve essere installato sullo stesso computer. Se l'installazione di un pacchetto chiede di riavviare BiblioHub, chiudi e riapri l'applicazione.

### Merge con errore di colonne

Merge bibliografie accetta gli Excel dello Scraper. Un file con lo stesso numero di colonne ma con intestazioni diverse viene rifiutato: il registro elenca le colonne mancanti.

### Topografie con errore di colonne

Le estrazioni devono avere almeno Insegnamento e Descrizione. Il registro deve avere Title (Complete). Il filtro sull'inventario richiede anche Inventory Date (calendar). Il messaggio indica il file e la colonna mancante. Se l'output scelto è uno dei file in ingresso, il confronto non parte.

## Com'è fatto il programma

BiblioHub è un'applicazione desktop in Python e tkinter. La logica di ogni strumento sta in bibliohub/tools e non dipende dalla finestra. Le schermate stanno in bibliohub/pages. Il registro degli strumenti, con versione e pacchetti richiesti, è bibliohub/registry.py.

### File principali

- main.py e BiblioHub.pyw — avvio, il secondo senza finestra nera.
- bibliohub/app.py — finestra, home, installazione dei pacchetti mancanti.
- bibliohub/widgets.py — campi file, registro attività, lavoro in background.
- bibliohub/guide.py — testi di questa guida. Il PDF si rigenera da qui.
- bibliohub/values.py — i numeri interi letti da Excel restano senza .0.
- bibliohub/settings.py — preferenze in %APPDATA%\BiblioHub\settings.json.
- bibliohub/assets/Template.docx — modello Word delle attestazioni.
- docs/Manuale.md e docs/Manuale_BiblioHub.pdf — questo manuale.

### Versioni attuali

- Applicazione 1.1.0.
- Scraper Insegnamenti 7.0.
- PO Line → Excel/CSV 2.1.
- Attestazioni di pagamento 3.1.
- Merge bibliografie 1.2.
- Topografie 1.1.

Dopo una modifica a bibliohub/guide.py il manuale si rigenera, dalla cartella del progetto, con: python docs/build_manual.py
