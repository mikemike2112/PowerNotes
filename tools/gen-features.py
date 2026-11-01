# Genera FEATURES.md dall'inventario del ribbon di Word (docs/word-ribbon-inventory.json, letto con UI Automation).
import json, io
inv = json.load(open('docs/word-ribbon-inventory.json', encoding='utf-8-sig'))
DONE = {  # comando -> nota (implementato E coperto da test)
 'Incolla': 'testo semplice', 'Taglia': 'testo semplice', 'Copia': 'testo semplice',
 'Grassetto': '', 'Corsivo': '', 'Sottolineato': '', 'Barrato': '', 'Pedice': '', 'Apice': '',
 'Colore carattere': '', 'Colore evidenziatore testo': '', 'Tipo di carattere': '', 'Dimensione carattere': '',
 'Elenchi puntati': '', 'Elenchi numerati': '', 'Elenco a più livelli': 'livelli con Aumenta/Riduci rientro o Tab',
 'Riduci rientro': '', 'Aumenta rientro': '', 'Allinea a sinistra': '', 'Allinea al centro': '', 'Allinea a destra': '', 'Giustifica': '',
 'Interlinea e spaziatura paragrafo': 'finestra Paragrafo e menu interlinea', 'Paragrafo...': '', 'Stili': 'galleria: Normale, Titolo, Sottotitolo, Titolo 1-3, Citazione, Paragrafo elenco',
 'Trova': '', 'Sostituisci...': '',
 'Interruzione di pagina': '', 'Tabella': 'inserimento, celle unite, righe/colonne', 'Immagini': 'inline PNG/JPEG/GIF/BMP',
 'Intestazione': 'finestra dedicata', 'Piè di pagina': 'finestra dedicata', 'Numero di pagina': 'Pagina X di Y nel piè di pagina',
 'Imposta pagina...': 'formato, orientamento, margini', 'Margini': 'in Imposta pagina', 'Orientamento': '', 'Dimensioni': 'A4, A5, Letter, Legal',
 'Layout di stampa': 'unica vista', '100%': '', 'Zoom...': 'pulsanti −/+/100% e Ctrl+rotella',
}
P0 = set(DONE) | {'Carattere...', 'Ingrandisci carattere', 'Riduci carattere', 'Colonne', 'Interruzioni', 'Annulla', 'Ripeti'}
P1 = {'Cancella formattazione','Mostra tutto','Sfondo','Bordi','Seleziona','Pagina vuota','Collegamento','Segnalibro...','Inserisci nota a piè di pagina','Conteggio parole','Ortografia e Grammatica','Rilevamento modifiche','Nuovo commento','Colonne','Interruzioni','Larghezza pagina','Una pagina','Righello','Sommario','Copia formato','Maiuscole/minuscole','Ordina...','Sillabazione','Numeri di riga','Casella di testo','Forme','Simbolo','Equazione','Commento'}
out = io.StringIO()
w = lambda s='': out.write(s + '\n')
w('# Funzionalità')
w()
w('Inventario ricavato dal vero Word con UI Automation (`tools/inventory-ribbon.ps1` → `docs/word-ribbon-inventory.json`): Word viene avviato con finestra visibile su un documento vuoto creato dallo script, si seleziona ogni scheda del ribbon e si leggono gruppi e comandi. Nessun documento dell\'utente viene aperto. Priorità: **P0** richiesto dal brief, **P1** utile, **P2** avanzato. Si spunta `[x]` solo ciò che è implementato **e** coperto da test automatici (xunit, corpus Word via COM, autotest dell\'interfaccia reale).')
w()
w('## File (fuori dal ribbon)')
for n, ok, pr, note in [
  ('Nuovo / Apri / Salva / Salva con nome (.docx nativo, salvataggio atomico)', True, 'P0', 'test round-trip e autotest UI'),
  ('Esporta in PDF (stessa pipeline di layout)', True, 'P0', 'test PDF su tutto il corpus + autotest UI'),
  ('Stampa (paginatore sul layout, finestra di stampa di sistema)', True, 'P0', 'verificato con scrittura XPS dal paginatore nell\'autotest UI; la stampa su carta reale non è stata provata'),
  ('Annulla / Ripeti (istantanee XML, digitazione raggruppata)', True, 'P0', ''),
  ('Selezione, caret, frecce, Home/End, PagSu/PagGiù, Ctrl+frecce, doppio/triplo click', True, 'P0', 'test sessione + autotest UI'),
  ('Tema scuro/chiaro commutabile', False, 'P1', 'implementato, verificato solo con screenshot'),
  ('Salvataggio con password, firma, versioni', False, 'P2', ''),
]:
    w(f"- [{'x' if ok else ' '}] **{pr}** {n}" + (f' — {note}' if note else ''))
for tab in inv['schede']:
    if tab not in inv: continue
    w(); w(f'## Scheda {tab}')
    for g, items in inv[tab].items():
        if g == tab: continue
        w(); w(f'### {g}')
        for it in items:
            if it in ('Altre opzioni', 'Apri') or it.startswith('Colore evidenziatore testo ') or it.startswith('Colore carattere ') or it.startswith('Sfondo RGB'): continue
            ok = it in DONE
            pr = 'P0' if it in P0 else 'P1' if it in P1 else 'P2'
            note = DONE.get(it, '')
            if not ok and it in ('Colonne',): note = 'solo rendering (test del corpus); nessun comando'
            if not ok and it in ('Interruzioni',): note = 'interruzione di pagina sì; interruzione di sezione solo letta/mostrata'
            w(f"- [{'x' if ok else ' '}] **{pr}** {it}" + (f' — {note}' if note else ''))
open('FEATURES.md', 'w', encoding='utf-8').write(out.getvalue())
