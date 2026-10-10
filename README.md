# Conti

App web per gestire entrate, uscite e saldi di casa. Funziona su iPhone, Samsung e PC, segue la dark mode del telefono e si può installare sulla schermata Home.

- Ognuno accede con il proprio account Google e ha **i propri conti**, separati.
- Ognuno può dare all'altro l'accesso **in sola lettura** (Impostazioni → Sola lettura per altri).
- **Esporta Excel** del mese o dell'anno con le formule (Controlli → Chiusura dei conti).
- **Importa o aggiorna da Excel**: se modifichi il file da PC lo ricarichi (Impostazioni → Importa o aggiorna da Excel) e l'app mostra le differenze mese per mese prima di applicarle.

Tutto gratuito: il sito sta su GitHub Pages, i dati su Firebase (piano Spark, senza carta di credito).

---

## Installazione (circa 15 minuti, una volta sola)

### 1. Crea il progetto Firebase
1. Vai su <https://console.firebase.google.com> e accedi con un account Google.
2. **Aggiungi progetto**, scegli un nome (per esempio `conti-casa`), disattiva Google Analytics, **Crea progetto**.

### 2. Attiva il login con Google
1. Nel menu a sinistra: **Build → Authentication → Inizia**.
2. Scheda **Sign-in method → Google → Abilita**, scegli l'email di supporto, **Salva**.

### 3. Crea il database
1. **Build → Firestore Database → Crea database**.
2. Posizione: una europea (per esempio `eur3` o Milano `europe-west8`).
3. Scegli **Modalità produzione** e conferma.
4. Apri la scheda **Regole**, cancella tutto, incolla il contenuto del file `firestore.rules` di questa cartella e premi **Pubblica**.

### 4. Collega l'app al progetto
1. In Firebase: icona ingranaggio → **Impostazioni progetto** → in fondo **Le tue app** → icona web `</>` → dai un nome → **Registra app** (senza Firebase Hosting).
2. Copia i valori di `firebaseConfig` e incollali nel file `js/firebase-config.js`, al posto dei segnaposto `INCOLLA_QUI`.

### 5. Pubblica su GitHub Pages
1. Su <https://github.com> crea un repository (per esempio `conti`). Può essere pubblico: nella cartella non c'è nessun dato personale, solo il programma.
2. Carica tutti i file di questa cartella nel repository (pulsante **Add file → Upload files**).
3. **Settings → Pages → Build and deployment**: Source = *Deploy from a branch*, Branch = `main`, cartella `/ (root)`, **Save**.
4. Dopo un minuto il sito è online su `https://TUONOME.github.io/conti/`.

### 6. Autorizza il sito per il login
In Firebase: **Authentication → Settings → Domini autorizzati → Aggiungi dominio** e scrivi `TUONOME.github.io` (solo il dominio, senza `https://` né `/conti`).

---

## Primo utilizzo

1. Apri il sito, **Accedi con Google**.
2. Tocca **Importa il file Excel** e scegli il tuo file `Conti_2026.xlsx`. Controlla l'anteprima e tocca **Importa**.
3. La tua ragazza fa lo stesso con il suo account e il suo file.
4. Per guardare i conti dell'altro: Impostazioni (ingranaggio) → **Sola lettura per altri** → scrivi l'email Google dell'altra persona. Lei vedrà in alto un selettore "I miei conti / Conti di …".

**Installare sul telefono**
- iPhone: apri il sito in Safari → Condividi → **Aggiungi alla schermata Home**.
- Samsung: apri il sito in Chrome → menu ⋮ → **Installa app** (o Aggiungi a schermata Home).

Se su iPhone l'accesso con Google non parte dall'icona sulla Home, aprilo una volta da Safari.

## Chiusura mensile o annuale

**Controlli → Chiusura dei conti**: scegli il mese e tocca **Scarica il mese**, oppure **Scarica tutto l'anno**.
Il file ha lo stesso schema del tuo foglio (un foglio per mese con Apertura e Chiusura, più il foglio Totale con le formule e i controlli). I grafici restano nell'app, non nel file Excel.

## Se modifichi il file da PC

1. Modifica il file Excel senza cambiare i nomi dei fogli (`Gennaio 2026`, …) né la posizione delle righe *Apertura* e *Chiusura* e delle colonne (Descrizione, Dettagli, 4 conti, Totale).
2. Nell'app: Impostazioni → **Importa o aggiorna da Excel**.
3. L'anteprima mostra per ogni mese *Uguale*, *Nuovo* oppure *+N / −N* (voci aggiunte o tolte). I mesi con differenze sono già selezionati: togli la spunta a quelli che non vuoi toccare.
4. Il mese selezionato viene **sostituito** con quello del file. Se nel frattempo avevi inserito voci nell'app per quel mese, esportale prima oppure togli la spunta.

L'import legge i valori salvati nel file, quindi salva il file in Excel prima di caricarlo.

## Provarla senza Firebase

Finché `js/firebase-config.js` ha i segnaposto, l'app parte in **modalità demo**: i dati restano solo nel browser del dispositivo. Serve per provare tutto prima di configurare Firebase. Per aprirla in locale: `python3 -m http.server` nella cartella e vai su `http://localhost:8000`.

## Come sono protetti i dati

- I valori di `firebaseConfig` non sono segreti: identificano il progetto. La protezione è nelle regole di `firestore.rules`.
- Le regole permettono di **scrivere solo nel proprio spazio**. Un'altra persona può solo **leggere**, e solo se la tua email (verificata da Google) è nell'elenco della condivisione.
- Per togliere l'accesso: Impostazioni → Sola lettura per altri → ×.

## Limiti del piano gratuito

Firebase Spark: 1 GiB di dati, 50.000 letture e 20.000 scritture al giorno. Per due persone che registrano qualche decina di voci al giorno restano molto sotto i limiti.

## File del progetto

| File | A cosa serve |
|---|---|
| `index.html`, `css/style.css` | Pagina e aspetto |
| `js/app.js` | Schermate e logica |
| `js/excel.js` | Esporta e importa Excel |
| `js/backend-firebase.js` | Login Google e database Firestore |
| `js/backend-local.js` | Modalità demo senza Firebase |
| `js/firebase-config.js` | **Da compilare** con i dati del tuo progetto |
| `firestore.rules` | Regole di accesso da incollare in Firebase |
| `manifest.webmanifest`, `sw.js`, `icons/` | Installazione sul telefono |
