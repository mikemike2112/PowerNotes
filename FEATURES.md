# Funzionalità

Inventario ricavato dal vero Word con UI Automation (`tools/inventory-ribbon.ps1` → `docs/word-ribbon-inventory.json`): Word viene avviato con finestra visibile su un documento vuoto creato dallo script, si seleziona ogni scheda del ribbon e si leggono gruppi e comandi. Nessun documento dell'utente viene aperto. Priorità: **P0** richiesto dal brief, **P1** utile, **P2** avanzato. Si spunta `[x]` solo ciò che è implementato **e** coperto da test automatici (xunit, corpus Word via COM, autotest dell'interfaccia reale).

## File (fuori dal ribbon)
- [x] **P0** Nuovo / Apri / Salva / Salva con nome (.docx nativo, salvataggio atomico) — test round-trip e autotest UI
- [x] **P0** Esporta in PDF (stessa pipeline di layout) — test PDF su tutto il corpus + autotest UI
- [x] **P0** Stampa (paginatore sul layout, finestra di stampa di sistema) — verificato con scrittura XPS dal paginatore nell'autotest UI; la stampa su carta reale non è stata provata
- [x] **P0** Annulla / Ripeti (istantanee XML, digitazione raggruppata)
- [x] **P0** Selezione, caret, frecce, Home/End, PagSu/PagGiù, Ctrl+frecce, doppio/triplo click — test sessione + autotest UI
- [ ] **P1** Tema scuro/chiaro commutabile — implementato, verificato solo con screenshot
- [ ] **P2** Salvataggio con password, firma, versioni

## Scheda Home

### Appunti
- [x] **P0** Incolla — testo semplice
- [x] **P0** Taglia — testo semplice
- [x] **P0** Copia — testo semplice
- [ ] **P1** Copia formato
- [ ] **P2** Appunti di Office...

### Carattere
- [ ] **P0** Ingrandisci carattere
- [ ] **P0** Riduci carattere
- [ ] **P1** Cancella formattazione
- [x] **P0** Grassetto
- [x] **P0** Corsivo
- [x] **P0** Sottolineato
- [x] **P0** Barrato
- [x] **P0** Pedice
- [x] **P0** Apice
- [ ] **P0** Carattere...
- [x] **P0** Colore evidenziatore testo
- [x] **P0** Colore carattere
- [x] **P0** Tipo di carattere
- [x] **P0** Dimensione carattere
- [ ] **P1** Maiuscole/minuscole
- [ ] **P2** Effetti testo e tipografia

### Paragrafo
- [x] **P0** Elenchi puntati
- [x] **P0** Elenchi numerati
- [x] **P0** Riduci rientro
- [x] **P0** Aumenta rientro
- [ ] **P1** Ordina...
- [ ] **P1** Mostra tutto
- [x] **P0** Allinea a sinistra
- [x] **P0** Allinea al centro
- [x] **P0** Allinea a destra
- [x] **P0** Giustifica
- [ ] **P1** Bordi
- [x] **P0** Paragrafo...
- [ ] **P1** Sfondo
- [x] **P0** Elenco a più livelli — livelli con Aumenta/Riduci rientro o Tab
- [x] **P0** Interlinea e spaziatura paragrafo — finestra Paragrafo e menu interlinea

### Stili
- [x] **P0** Stili — galleria: Normale, Titolo, Sottotitolo, Titolo 1-3, Citazione, Paragrafo elenco

### Modifica
- [x] **P0** Trova
- [x] **P0** Sostituisci...
- [ ] **P1** Seleziona

### Voce
- [ ] **P2** Dettatura

### Riservatezza
- [ ] **P2** Riservatezza

### Editor
- [ ] **P2** Editor

### Componenti aggiuntivi
- [ ] **P2** Add‑ins

## Scheda Inserisci

### Pagine
- [ ] **P1** Pagina vuota
- [x] **P0** Interruzione di pagina
- [ ] **P2** Frontespizio

### Tabelle
- [x] **P0** Tabella — inserimento, celle unite, righe/colonne

### Illustrazioni
- [ ] **P2** Icone
- [ ] **P2** Modelli 3D
- [ ] **P2** SmartArt...
- [ ] **P2** Grafico...
- [x] **P0** Immagini — inline PNG/JPEG/GIF/BMP
- [ ] **P1** Forme
- [ ] **P2** Schermata

### Elementi multimediali
- [ ] **P2** Video online...

### Collegamenti
- [ ] **P1** Collegamento
- [ ] **P1** Segnalibro...
- [ ] **P2** Riferimento incrociato...

### Commenti
- [ ] **P1** Commento

### Intestazione e piè di pagina
- [x] **P0** Intestazione — finestra dedicata
- [x] **P0** Piè di pagina — finestra dedicata
- [x] **P0** Numero di pagina — Pagina X di Y nel piè di pagina

### Testo
- [ ] **P2** Riga della firma
- [ ] **P2** Data e ora...
- [ ] **P2** Oggetto...
- [ ] **P1** Casella di testo
- [ ] **P2** Parti rapide
- [ ] **P2** WordArt
- [ ] **P2** Capolettera

### Simboli
- [ ] **P1** Equazione
- [ ] **P1** Simbolo

### eSignature
- [ ] **P2** Campi eSignature edizione Standard

## Scheda Progettazione

### Formattazione documento
- [ ] **P2** Set di stili
- [ ] **P2** Imposta come predefinito
- [ ] **P2** Temi
- [ ] **P2** Colori
- [ ] **P2** Tipi di carattere
- [ ] **P2** Spaziatura paragrafo
- [ ] **P2** Effetti

### Set di stili
- [ ] **P2** Set di stili

### Sfondo pagina
- [ ] **P2** Bordi pagina...
- [ ] **P2** Filigrana
- [ ] **P2** Colore pagina

## Scheda Layout

### Imposta pagina
- [x] **P0** Imposta pagina... — formato, orientamento, margini
- [x] **P0** Margini — in Imposta pagina
- [x] **P0** Orientamento
- [x] **P0** Dimensioni — A4, A5, Letter, Legal
- [ ] **P0** Colonne — solo rendering (test del corpus); nessun comando
- [ ] **P0** Interruzioni — interruzione di pagina sì; interruzione di sezione solo letta/mostrata
- [ ] **P1** Numeri di riga
- [ ] **P1** Sillabazione

### Paragrafo
- [ ] **P2** Aumenta
- [ ] **P2** Diminuisci
- [x] **P0** Paragrafo...

### Disponi
- [ ] **P2** Porta avanti
- [ ] **P2** Porta indietro
- [ ] **P2** Riquadro di selezione...
- [ ] **P2** Posizione
- [ ] **P2** Testo a capo
- [ ] **P2** Allinea
- [ ] **P2** Raggruppa
- [ ] **P2** Ruota

## Scheda Riferimenti

### Sommario
- [ ] **P2** Aggiorna sommario...
- [ ] **P1** Sommario
- [ ] **P2** Aggiungi testo

### Note a piè di pagina
- [ ] **P1** Inserisci nota a piè di pagina
- [ ] **P2** Inserisci nota di chiusura
- [ ] **P2** Nota a piè di pagina successiva
- [ ] **P2** Mostra note
- [ ] **P2** Finestra di dialogo Note a piè di pagina e di chiusura...

### Citazioni e bibliografia
- [ ] **P2** Gestisci fonti...
- [ ] **P2** Stile
- [ ] **P2** Inserisci citazione
- [ ] **P2** Bibliografia
- [ ] **P2** Cambia provider

### Didascalie
- [ ] **P2** Inserisci didascalia...
- [ ] **P2** Inserisci indice delle figure...
- [ ] **P2** Aggiorna tabella
- [ ] **P2** Riferimento incrociato...

### Indice
- [ ] **P2** Segna voce...
- [ ] **P2** Inserisci indice...
- [ ] **P2** Aggiorna indice

### Indice delle fonti
- [ ] **P2** Segna citazione...
- [ ] **P2** Inserisci indice delle fonti...
- [ ] **P2** Aggiorna tabella

## Scheda Lettere

### Crea
- [ ] **P2** Buste...
- [ ] **P2** Etichette...

### Inizia stampa unione
- [ ] **P2** Modifica elenco destinatari...
- [ ] **P2** Inizia stampa unione
- [ ] **P2** Seleziona destinatari

### Inserisci campi
- [ ] **P2** Evidenzia campi unione
- [ ] **P2** Blocco di indirizzi...
- [ ] **P2** Formula di apertura...
- [ ] **P2** Inserisci campo unione
- [ ] **P2** Corrispondenza campi...
- [ ] **P2** Aggiorna etichette
- [ ] **P2** Regole

### Anteprima risultati
- [ ] **P2** Anteprima risultati
- [ ] **P2** Primo
- [ ] **P2** Precedente
- [ ] **P2** Successivo
- [ ] **P2** Ultimo
- [ ] **P2** Trova destinatario...
- [ ] **P2** Controllo errori...
- [ ] **P2** Record

### Fine
- [ ] **P2** Finalizza e unisci

## Scheda Revisione

### Strumenti di correzione
- [ ] **P2** Editor
- [ ] **P1** Ortografia e Grammatica
- [ ] **P2** Thesaurus...
- [ ] **P1** Conteggio parole

### Comandi vocali
- [ ] **P2** Leggi ad alta voce

### Accessibilità
- [ ] **P2** Verifica accessibilità

### Lingua
- [ ] **P2** Lingua

### Commenti
- [ ] **P1** Nuovo commento
- [ ] **P2** Elimina
- [ ] **P2** Precedente
- [ ] **P2** Successivo
- [ ] **P2** Mostra commenti

### Rilevamento modifiche
- [ ] **P1** Rilevamento modifiche

### Confronta
- [ ] **P2** Confronta

### Proteggi
- [ ] **P2** Proteggi

### Input penna
- [ ] **P2** Nascondi input penna

## Scheda Visualizza

### Visualizzazioni
- [ ] **P2** Modalità di lettura
- [x] **P0** Layout di stampa — unica vista
- [ ] **P2** Layout Web
- [ ] **P2** Struttura
- [ ] **P2** Bozza

### Immersiva
- [ ] **P2** Focus
- [ ] **P2** Strumento di lettura immersiva

### Spostamento tra pagine
- [ ] **P2** Verticale
- [ ] **P2** Orizzontale

### Mostra
- [ ] **P1** Righello
- [ ] **P2** Griglia
- [ ] **P2** Riquadro di spostamento

### Zoom
- [x] **P0** Zoom... — pulsanti −/+/100% e Ctrl+rotella
- [x] **P0** 100%
- [ ] **P1** Una pagina
- [ ] **P2** Più pagine
- [ ] **P1** Larghezza pagina

### Finestra
- [ ] **P2** Nuova finestra
- [ ] **P2** Disponi tutto
- [ ] **P2** Dividi
- [ ] **P2** Affianca
- [ ] **P2** Scorrimento sincrono
- [ ] **P2** Reimposta posizione finestra
- [ ] **P2** Cambia finestra

### Macro
- [ ] **P2** Visualizza macro
- [ ] **P2** Macro

### SharePoint
- [ ] **P2** Proprietà

## Scheda Sviluppo

### Codice
- [ ] **P2** Visual Basic
- [ ] **P2** Macro
- [ ] **P2** Registra macro...
- [ ] **P2** Sospendi registrazione
- [ ] **P2** Sicurezza macro

### Componenti aggiuntivi
- [ ] **P2** Componenti aggiuntivi
- [ ] **P2** Componenti aggiuntivi di Word...
- [ ] **P2** Componenti aggiuntivi COM...

### Controlli
- [ ] **P2** Controllo contenuto RTF
- [ ] **P2** Controllo contenuto testo normale
- [ ] **P2** Controllo contenuto immagine
- [ ] **P2** Controllo contenuto raccolta blocchi predefiniti
- [ ] **P2** Controllo contenuto di tipo casella di controllo
- [ ] **P2** Controllo contenuto casella combinata
- [ ] **P2** Controllo contenuto casella di riepilogo a discesa
- [ ] **P2** Controllo contenuto Selezione data
- [ ] **P2** Controllo contenuto sezione ripetuta
- [ ] **P2** Modalità progettazione
- [ ] **P2** Proprietà
- [ ] **P2** Strumenti legacy
- [ ] **P2** Raggruppa

### Mapping
- [ ] **P2** Riquadro mapping XML

### Proteggi
- [ ] **P2** Blocca autori
- [ ] **P2** Limita modifica

### Modelli
- [ ] **P2** Modello di documento

## Scheda Guida
