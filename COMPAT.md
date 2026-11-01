# Compatibilità con Word

Tutti i numeri sotto sono misurati con il vero Word (COM, script in `tools/`) sul corpus `tests/corpus/office-made/` (37 file generati da Word, Author/LastAuthor = "Corpus", verificato da test).

## Round-trip (apri → salva → riapri in Word)

| Controllo | Esito |
|---|---|
| Corpus | 37 file |
| Salvato dal clone si apre in Word senza errore né ripristino, stesso testo, pagine, paragrafi, tabelle, immagini, sezioni (letti da Word) | **37/37 (100%)** |
| `OpenXmlValidator`: errori nuovi rispetto all'originale | 0 (i 36 originali di Word hanno 0 errori: nessun falso positivo da elencare) |
| Apri+salva senza modifiche: stesse parti e relazioni, ogni part XML canonicalizzata identica | 37/37 |
| File con modifica (14 file: testo digitato, Sostituisci tutto, grassetto su tutto) riletti da Word: modifica presente, il resto invariato | 14/14 |
| Contenuto sconosciuto (namespace estraneo in `mc:Ignorable`, attributo sconosciuto, elemento annidato, part personalizzata con relazione) dopo modifica e salvataggio | conservato |
| Salvataggio atomico (file temporaneo + sostituzione; sorgente intatto se il salvataggio fallisce) | testato |
| File malformati (non zip, XML rotto) | `DocumentLoadException` con messaggio chiaro, nessun crash |

### Controllo negativo del rilevatore
Il rilevatore è Word (apertura via COM senza errore) più un controllo di integrità del package (relazioni, parti, tipi di contenuto, `r:id`). Su quattro file volutamente corrotti: XML rotto → **Word rifiuta**; tipo di contenuto sbagliato → **Word rifiuta**; parte mancante (`styles.xml` cancellato ma referenziato) e relazione verso immagine inesistente → **Word li apre in silenzio** (tollera il danno), ma il controllo di integrità li segnala. Per questo il rilevatore è composto: Word da solo non basterebbe. I 37 file salvati e i 36 originali passano l'integrità senza falsi positivi.

## Layout: conta di pagine e posizioni

| Confronto con Word (`ComputeStatistics` e `Information(wdVerticalPositionRelativeToPage)`) | Esito |
|---|---|
| Numero di pagine uguale | **37/37** |
| Pagina e quota verticale della prima riga di ogni paragrafo entro 3 pt (file senza tabelle/immagini: 25 file) | 25/25 (la maggior parte entro 0,3 pt, per esempio `long-document`, 9 pagine) |

Regole di Word riprodotte e verificate sul corpus: spazio fra paragrafi = **massimo** fra Dopo e Prima (non somma); spaziatura contestuale; spazio prima soppresso in cima a pagina dopo salto automatico; paragrafo con la sola interruzione di sezione senza riga; l'interlinea multipla scala solo il testo; altezza riga dai metrici `hhea` del font (anche il font Symbol dei puntini); orfane/vedove; font "cloud" di Office (Aptos) letti dalla cache utente.

## Differenze visive osservate (PDF Word vs PDF clone, rasterizzati con `Windows.Data.Pdf`, `tools/CompatTool`)
Confrontate a vista 10 file (`mixed-report`, `styles-headings`, `text-formatting`, `table-merged-cells`, `list-multilevel`, `mixed-illustrated`, `page-numbers`, `tabs`, `alignment`, `long-document`). Differenze reali osservate:
- Tabelle con bordo 0,5 pt: nel PDF del clone le linee risultano un poco più scure (antialiasing del rasterizzatore).
- `mixed-illustrated`: dopo l'immagine inline il testo successivo è circa 3 px più in basso rispetto a Word (la riga dell'immagine ha un residuo di discesa diverso); nelle pagine con elenco l'accumulo arriva a ~4 px.
- Il grassetto di font senza variante Bold installata (es. Calibri Light) non è sintetizzato: resta regolare.
- `table-merged-cells` dopo la correzione dei RowSpan coincide con Word a vista.

## Limiti noti (non implementato o semplificato)
- Spaziatura carattere e kerning/legature ignorati (larghezze da `hmtx`): ±1% sulle lunghezze di riga, nessuna differenza di conta pagine nel corpus.
- Oggetti ancorati: immagini e caselle di testo sono **disegnate** alla posizione indicata (pagina, margine, paragrafo; allineamenti); il testo non scorre attorno (`wrapSquare/Tight`), mentre "sopra e sotto" riserva lo spazio solo se ancorato al paragrafo. Forme vettoriali e rotazioni non disegnate (restano nel file).
- Note a piè di pagina: contenuto conservato, non ancora impaginato.
- Righe di tabella più alte di una pagina non si spezzano. Intestazioni di tabella ripetute e stili di tabella condizionali (prima riga, prima colonna, bande) confrontati con Word su `table-style-header-repeat.docx` (stile Tabella griglia 4, 70 righe): stesse pagine (2), stessa intestazione ripetuta, ombreggiature e grassetti uguali a vista; differenza osservata: Word mette 51 righe sulla prima pagina, il clone 53 (altezza riga leggermente minore nel clone).
- Revisioni: il testo inserito è mostrato, il cancellato no; non si possono accettare/rifiutare.
- Intestazioni/piè di pagina: modifica sul foglio (doppio clic nella zona, Esc per uscire) o dalla finestra dedicata; la prima pagina diversa e pari/dispari sono lette, non create da comandi.
- Appunti: testo semplice + frammento formattato fra documenti di PowerNotes; da/verso altre applicazioni solo testo semplice.
- Grassetto/corsivo sintetizzati (doppia stampa/inclinazione) se il font non ha la variante; qualità inferiore a un vero font Bold.
- Font `.ttc` incorporati nel PDF estraendo il singolo font.
- Nessuna correzione ortografica, nessun campo oltre a PAGE/NUMPAGES.
