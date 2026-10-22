# BiblioHub

Programma Windows per le bibliografie d'esame, gli ordini Alma e le attestazioni di pagamento dell'Università degli Studi di Milano.

Dentro l'app, **ⓘ Guida** in alto a destra spiega la schermata aperta. Il manuale completo è `docs/Manuale_BiblioHub.pdf`.

## Scarica e avvia

**1. Installa Python, se non ce l'hai già**

Apri [python.org/downloads/windows](https://www.python.org/downloads/windows/), scarica Python 3.11 o più nuovo e avvialo. Nella prima schermata spunta **Add python.exe to PATH**, poi Install.

**2. Scarica BiblioHub**

Su questa pagina di GitHub premi il pulsante verde **Code**, poi **Download ZIP**. Apri lo ZIP e premi **Estrai tutto**.

**3. Avvia**

Apri la cartella appena estratta. Clicca una volta la barra dell'indirizzo in alto (dove c'è il percorso), scrivi `cmd` e premi Invio.

Nella finestra nera incolla queste due righe, una alla volta, e premi Invio dopo ognuna:

```bat
pip install -r requirements.txt
python main.py
```

Si apre BiblioHub. Le volte successive, dalla stessa cartella, ripeti solo `python main.py`.

Microsoft Word serve solo se generi i PDF delle attestazioni. Se una finestra dice che manca un pacchetto, accetta l'installazione.

Se al passo 3 compare «python non è riconosciuto», chiudi la finestra nera, ripeti il passo 1 controllando la spunta su PATH, poi rifai il passo 3.

## Download and run

**1. Install Python if you don't have it**

Open [python.org/downloads/windows](https://www.python.org/downloads/windows/), download Python 3.11 or newer and run it. On the first screen tick **Add python.exe to PATH**, then Install.

**2. Download BiblioHub**

On this GitHub page click the green **Code** button, then **Download ZIP**. Open the ZIP and click **Extract all**.

**3. Start**

Open the folder you just extracted. Click the address bar at the top, type `cmd` and press Enter.

In the black window paste these two lines, one at a time, and press Enter after each:

```bat
pip install -r requirements.txt
python main.py
```

BiblioHub opens. Next time, from the same folder, run only `python main.py`.

Microsoft Word is needed only for attestation PDFs. If a window says a package is missing, accept the install.

If step 3 says Python is not recognized, close the black window, repeat step 1 and check the PATH box, then do step 3 again.
