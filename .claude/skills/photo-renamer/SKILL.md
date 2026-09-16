---
name: photo-renamer
description: Rinomina in batch le foto prodotto con il relativo codice EAN, deducendo dal nome dei file quali colonne dell'Excel compongono il codice prodotto e quale colonna è l'EAN/barcode. Usa questa skill quando l'utente chiede di rinominare/abbinare foto con i codici EAN a partire da una cartella assets/<stagione>/<brand>/ (es. "rinomina le foto di borbonese per ai26", "/photo-renamer ai26 borbonese").
---

# Photo Renamer (MSAFR)

Rinomina le foto di un brand con il codice EAN del prodotto corrispondente,
leggendo i dati da un file Excel multi-foglio. Sostituisce il vecchio script
`multi_sheet_photo_renamer.py` con mappature di colonne hard-coded per brand:
qui le colonne vengono **dedotte** osservando i nomi dei file e i dati Excel,
e la deduzione viene salvata per essere riusata nelle stagioni successive.

## Struttura cartelle attesa

```
assets/<stagione>/<brand>/
├── *.xlsx              <- un solo file Excel, uno o più fogli
├── foto1.jpg
├── foto2.jpg
└── ...
```

`assets/` è nel `.gitignore`: è materiale del cliente, non va versionato.

## Invocazione

L'utente specifica **stagione** e **brand** (es. `ai26 borbonese`), corrispondenti
a `assets/ai26/borbonese/`. Se manca uno dei due, chiedilo prima di procedere:
non indovinare quale sottocartella di `assets/` usare se ce n'è più di una.

## Passo 1 — Config già nota? Riusala, ma verificala

Controlla se esiste già `.claude/skills/photo-renamer/brand_configs/<brand>.json`
(nomi file in minuscolo). Se esiste, salta al Passo 3 usando quella
configurazione — ma tratta l'esito del dry-run come verifica: se il file Excel
di questa stagione ha una struttura diversa (colonne mancanti, troppi RED FLAG),
torna al Passo 2 per ri-dedurre la configurazione e sovrascrivila.

## Passo 2 — Deduzione colonna EAN e colonne codice (solo la prima volta per brand)

Esegui:
```
python3 .claude/skills/photo-renamer/scripts/inspect_brand.py <stagione> <brand>
```

Questo stampa, senza decidere nulla: intestazioni di ogni foglio, alcune righe
complete di esempio, statistiche per colonna (% di valori solo numerici,
lunghezze più comuni, esempi di valori), e un campione di nomi foto con il
"codice candidato" (la parte prima del primo `_`/`-`).

Analizza l'output tu stesso (ragionamento, non uno script) per dedurre:

1. **Colonna EAN**: di solito il nome contiene "EAN", "BARCODE", "GTIN" o
   "CODICE A BARRE"; i valori sono quasi sempre stringhe **solo numeriche di
   lunghezza 8, 12, 13 o 14** (EAN-8/UPC-12/EAN-13/ITF-14). Usa nome colonna +
   statistiche insieme, non uno dei due da solo.

2. **Colonne codice** (`code_columns`, in ordine): il sottoinsieme di colonne
   la cui concatenazione (valori ripuliti, maiuscolo/minuscolo indifferente)
   riproduce il "codice candidato" visto nei nomi foto. Sono in genere campi
   tipo Modello/Codice/Parte/Colore/Variante — **non** descrizioni testuali
   libere, indirizzi, clienti, pesi, prezzi o tariffe doganali. Verifica
   l'ipotesi concatenando a mano 3-5 righe campione e confrontando col codice
   dei file (puoi scrivere un piccolo snippet Python usando
   `xlsx_reader.read_workbook(...)` per farlo velocemente, invece di fidarti
   a occhio).

   Tieni conto che per alcuni prodotti non tutti i campi sono valorizzati:
   è normale e lo script gestisce già il matching parziale (colonne vuote
   vengono saltate nella concatenazione). Non serve che tu forzi un match
   perfetto al 100% delle righe.

3. Se non riesci a determinare un sottoinsieme di colonne che spieghi la
   maggioranza dei codici visti nei file, chiedi conferma all'utente prima
   di proseguire piuttosto che indovinare.

Salva la configurazione dedotta:
```
python3 .claude/skills/photo-renamer/scripts/save_brand_config.py <brand> \
  --ean-column "NOME_COLONNA_EAN" \
  --code-columns "Colonna1,Colonna2,Colonna3" \
  --notes "breve nota su come hai dedotto lo schema"
```

## Passo 3 — Dry-run e controllo dei RED FLAG

```
python3 .claude/skills/photo-renamer/scripts/rename_photos.py <stagione> <brand> --dry-run
```

Lo script:
- prova prima un match **esatto** (codice concatenato == token iniziale del
  nome file);
- se non trova nulla, ripiega su un match **parziale per sottostringa**
  (come il vecchio script);
- **RED FLAG**: se un match parziale/generico produce più di 7 file
  candidati per lo stesso EAN, NON rinomina nulla per quell'EAN e lo segna
  in un report `RED_FLAG_da_rivedere_*.csv` — troppi candidati con colonne
  parziali è quasi sempre segno che le `code_columns` scelte sono troppo
  generiche (regex/colonne del codice non abbastanza specifiche).

Leggi il riepilogo stampato a schermo:
- Se **ci sono RED FLAG**: guarda `RED_FLAG_da_rivedere_*.csv` in
  `assets/<stagione>/<brand>/reports/`, capisci quale colonna aggiuntiva
  serve per disambiguare (es. una colonna Taglia/Variante mancante), aggiorna
  la config col Passo 2 e ripeti il dry-run finché i RED FLAG sono spariti o
  comunque ridotti a casi genuinamente ambigui da segnalare all'utente.
- Se il numero di foto "rinominabili" è molto basso rispetto al totale delle
  foto trovate, probabilmente la config è sbagliata (colonne invertite,
  colonna EAN sbagliata): NON procedere al passo 4, torna al Passo 2.

## Passo 4 — Conferma ed esecuzione reale

Mostra all'utente il riepilogo del dry-run (quante foto verrebbero
rinominate, quanti EAN senza foto, eventuali RED FLAG residui) e chiedi
conferma esplicita prima di rinominare davvero i file — è un'operazione che
tocca file reali del cliente. Dopo conferma:

```
python3 .claude/skills/photo-renamer/scripts/rename_photos.py <stagione> <brand>
```

(stessa invocazione, senza `--dry-run`). Riporta all'utente il riepilogo
finale e il percorso dei report in `assets/<stagione>/<brand>/reports/`
(mapping vecchio→nuovo nome, EAN senza foto, eventuali RED FLAG non risolti).

## Note

- L'ottimizzazione JPEG (`jpegoptim`) resta opzionale e "best effort": se il
  binario non è installato viene saltata senza errori. Con `--skip-optimize`
  la si disattiva esplicitamente.
- Nessuna dipendenza esterna (niente pandas/openpyxl): la lettura xlsx è
  fatta con `xlsx_reader.py`, solo standard library.
- Le righe Excel duplicate per lo stesso EAN (tipico quando lo stesso
  prodotto compare per più clienti/filiali) vengono deduplicate
  automaticamente: vince la prima occorrenza.
