# 📸 Multi Sheet Awesome Photo Renamer (MSAFR)

MSAFR rinomina in batch le foto prodotto usando il relativo codice EAN,
ricavato da un file Excel multi-foglio. È distribuito come **skill per
Claude Code**: non serve più mantenere a mano una mappatura di colonne per
ogni brand nello script — la skill deduce da sola quale colonna è l'EAN e
quali colonne compongono il codice prodotto presente nel nome dei file,
osservando i dati.

## 🚀 Funzionalità
✅ Lettura di Excel multi-foglio senza dipendenze esterne (solo standard library)
✅ Deduzione automatica (assistita da Claude) della colonna EAN e delle colonne del codice prodotto, per brand
✅ Matching esatto con fallback parziale, con segnalazione (RED FLAG) dei match troppo generici
✅ Ottimizzazione facoltativa delle JPEG con `jpegoptim`
✅ Report CSV di rinomine, EAN senza foto e casi da rivedere manualmente

## 📂 Struttura cartelle

```
📁 Progetto (repo)
└── 📁 assets                       <-- gitignored, materiale cliente
    └── 📁 ai26                     <-- nome stagione
        └── 📁 borbonese            <-- nome brand
            ├── borbonese.xlsx      <-- Excel del brand (uno o più fogli)
            ├── foto1.jpg
            └── foto2.jpg
```

## 🔧 Utilizzo

In Claude Code, dalla root del progetto, invoca la skill indicando stagione
e brand (la sottocartella di `assets/`):

```
/photo-renamer ai26 borbonese
```

oppure semplicemente chiedendo in linguaggio naturale, es. "rinomina le foto
di borbonese per la stagione ai26". La skill:

1. La prima volta per un brand, analizza intestazioni Excel e nomi dei file
   per dedurre colonna EAN e colonne del codice prodotto, e salva la
   configurazione in `.claude/skills/photo-renamer/brand_configs/<brand>.json`
   (riusata automaticamente nelle stagioni successive).
2. Esegue un'anteprima (dry-run) e valuta eventuali RED FLAG (match troppo
   generici).
3. Dopo conferma, rinomina davvero le foto e produce i report in
   `assets/<stagione>/<brand>/reports/`.

Dettagli completi del funzionamento: `.claude/skills/photo-renamer/SKILL.md`.

### Uso manuale degli script (senza Claude)

Gli script sono normali script Python (solo standard library, nessun
`requirements.txt` necessario) e possono essere lanciati anche a mano:

```bash
python3 .claude/skills/photo-renamer/scripts/inspect_brand.py ai26 borbonese
python3 .claude/skills/photo-renamer/scripts/save_brand_config.py borbonese \
  --ean-column BARCODE --code-columns "Modello,Parte,Colore"
python3 .claude/skills/photo-renamer/scripts/rename_photos.py ai26 borbonese --dry-run
python3 .claude/skills/photo-renamer/scripts/rename_photos.py ai26 borbonese
```

### jpegoptim (opzionale)
https://github.com/tjko/jpegoptim — se non installato, l'ottimizzazione viene
semplicemente saltata.

## 📜 Licenza
GNU License
