#!/usr/bin/env python3
"""Rinomina le foto prodotto con il relativo codice EAN, deducendo la
corrispondenza foto<->riga Excel dal codice prodotto contenuto nel nome file.

Uso:
  rename_photos.py <stagione> <brand> [--dry-run] [--skip-optimize]
                    [--ean-column NOME] [--code-columns "A,B,C"]

Se --ean-column/--code-columns non sono passati, viene usata la
configurazione salvata in brand_configs/<brand>.json (vedi save_brand_config.py).

Logica di matching per ogni riga (deduplicata per EAN):
  1. match ESATTO: il token iniziale del nome file (prima del primo "_"/"-")
     coincide (case-insensitive) con la concatenazione di TUTTE le colonne
     codice non vuote per quella riga.
  2. se nessun match esatto, fallback PARZIALE: tutte le colonne codice non
     vuote compaiono come sottostringa nel nome file (matching per substring,
     come nello script originale).
  3. RED FLAG: se il match e' parziale (fallback substring, oppure riga con
     colonne codice mancanti) e i file candidati sono più di 7, non si
     rinomina nulla per quell'EAN: probabilmente le colonne scelte sono
     troppo generiche. Viene riportato in un report dedicato per revisione
     manuale, invece di rinominare a caso.
"""
import argparse
import csv
import os
import subprocess
import sys
import time
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brand_config import load_config
from discovery import code_token, find_excel, list_photos, require_brand_dir
from xlsx_reader import read_workbook

RED_FLAG_THRESHOLD = 7
MAX_SIZE_MB = 1
JPEG_QUALITY = 85


def optimize_jpeg_image(file_path):
    if os.path.getsize(file_path) / (1024 * 1024) <= MAX_SIZE_MB:
        return None
    try:
        process = subprocess.run(
            ["jpegoptim", f"--max={JPEG_QUALITY}", "--strip-all", file_path],
            capture_output=True, text=True,
        )
        return process.returncode == 0
    except FileNotFoundError:
        return None  # jpegoptim non installato: ottimizzazione facoltativa, si salta


def optimize_images_in_folder(folder_path):
    optimized = 0
    for filename in list_photos(folder_path):
        result = optimize_jpeg_image(os.path.join(folder_path, filename))
        if result:
            optimized += 1
    if optimized:
        print(f"Ottimizzate {optimized} immagini (>{MAX_SIZE_MB}MB) con jpegoptim.")


def resolve_config(brand, cli_ean_column, cli_code_columns):
    saved = load_config(brand) or {}
    ean_column = cli_ean_column or saved.get("ean_column")
    code_columns = cli_code_columns or saved.get("code_columns")

    if not ean_column or not code_columns:
        sys.exit(
            f"Configurazione mancante per il brand '{brand}'.\n"
            f"Esegui prima inspect_brand.py per individuare le colonne, poi "
            f"save_brand_config.py per salvarle (oppure passa --ean-column e --code-columns)."
        )
    return ean_column, code_columns


def validate_columns(workbook, ean_column, code_columns):
    all_columns = set()
    for records in workbook.values():
        if records:
            all_columns.update(records[0].keys())

    missing = [c for c in [ean_column, *code_columns] if c not in all_columns]
    if missing:
        sys.exit(
            f"Colonne non trovate nel file Excel: {missing}\n"
            f"Colonne disponibili: {sorted(all_columns)}"
        )


def iter_unique_ean_rows(workbook, ean_column):
    """Itera tutte le righe di tutti i fogli, deduplicando per EAN (prima occorrenza vince)."""
    seen = set()
    for sheet_name, records in workbook.items():
        for row in records:
            ean = row.get(ean_column)
            if ean in (None, ""):
                continue
            ean = str(ean).strip()
            if ean in seen:
                continue
            seen.add(ean)
            yield sheet_name, ean, row


def build_code(row, code_columns):
    """Ritorna (codice_concatenato, colonne_usate, is_partial)."""
    segments = []
    for col in code_columns:
        val = row.get(col)
        if val not in (None, ""):
            segments.append((col, str(val).strip()))
    full_code = "".join(v for _, v in segments)
    is_partial = len(segments) < len(code_columns)
    return full_code, [c for c, _ in segments], is_partial


def find_candidates(full_code, segments, remaining_files):
    """Ritorna (candidati, tipo_match)."""
    if not full_code:
        return [], None

    exact = [f for f in remaining_files if code_token(f).lower() == full_code.lower()]
    if exact:
        return sorted(exact), "esatto"

    if segments:
        values = [v.lower() for _, v in segments]
        partial = [f for f in remaining_files if all(v in f.lower() for v in values)]
        if partial:
            return sorted(partial), "parziale(substring)"

    return [], None


def write_csv_report(path, fieldnames, rows):
    if not rows:
        return None
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    return path


def run(season, brand, ean_column, code_columns, dry_run, skip_optimize):
    brand_dir = require_brand_dir(season, brand)
    excel_path = find_excel(brand_dir)
    workbook = read_workbook(excel_path)
    validate_columns(workbook, ean_column, code_columns)

    print(f"Cartella: {brand_dir}")
    print(f"Excel: {os.path.basename(excel_path)} (fogli: {list(workbook.keys())})")
    print(f"Colonna EAN: {ean_column!r}  |  Colonne codice: {code_columns}")
    if dry_run:
        print(">>> MODALITA' DRY-RUN: nessun file verra' rinominato <<<")

    if not skip_optimize and not dry_run:
        optimize_images_in_folder(brand_dir)

    photos = list_photos(brand_dir)
    print(f"Foto trovate: {len(photos)}")
    remaining_files = set(photos)

    renamed_rows = []
    unmatched_rows = []
    red_flag_rows = []
    renamed_count = 0

    for sheet_name, ean, row in iter_unique_ean_rows(workbook, ean_column):
        full_code, used_columns, is_partial = build_code(row, code_columns)
        candidates, match_type = find_candidates(full_code, [(c, row.get(c)) for c in used_columns], remaining_files)

        if not candidates:
            unmatched_rows.append({
                "foglio": sheet_name,
                "ean": ean,
                "codice_atteso": full_code,
                "colonne_usate": ",".join(used_columns),
                "motivo": "nessuna foto corrispondente",
            })
            continue

        is_generic_match = is_partial or match_type != "esatto"
        if is_generic_match and len(candidates) > RED_FLAG_THRESHOLD:
            red_flag_rows.append({
                "foglio": sheet_name,
                "ean": ean,
                "codice_atteso": full_code,
                "colonne_usate": ",".join(used_columns),
                "tipo_match": match_type,
                "n_candidati": len(candidates),
                "esempio_candidati": ";".join(candidates[:10]),
            })
            continue

        for i, old_name in enumerate(candidates, start=1):
            ext = os.path.splitext(old_name)[1]
            new_name = f"{ean}_{i}{ext}"
            renamed_rows.append({
                "vecchio_nome": old_name,
                "nuovo_nome": new_name,
                "ean": ean,
                "foglio": sheet_name,
                "tipo_match": match_type,
                "colonne_usate": ",".join(used_columns),
            })
            if not dry_run:
                os.rename(os.path.join(brand_dir, old_name), os.path.join(brand_dir, new_name))
            remaining_files.discard(old_name)
            renamed_count += 1

    reports_dir = os.path.join(brand_dir, "reports")
    os.makedirs(reports_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    prefix = "dryrun_" if dry_run else ""

    renamed_report = write_csv_report(
        os.path.join(reports_dir, f"{prefix}rinominati_{brand}_{ts}.csv"),
        ["vecchio_nome", "nuovo_nome", "ean", "foglio", "tipo_match", "colonne_usate"],
        renamed_rows,
    )
    unmatched_report = write_csv_report(
        os.path.join(reports_dir, f"{prefix}non_trovati_{brand}_{ts}.csv"),
        ["foglio", "ean", "codice_atteso", "colonne_usate", "motivo"],
        unmatched_rows,
    )
    red_flag_report = write_csv_report(
        os.path.join(reports_dir, f"{prefix}RED_FLAG_da_rivedere_{brand}_{ts}.csv"),
        ["foglio", "ean", "codice_atteso", "colonne_usate", "tipo_match", "n_candidati", "esempio_candidati"],
        red_flag_rows,
    )

    print()
    print("=== Riepilogo ===")
    verb = "rinominabili (dry-run)" if dry_run else "rinominati"
    print(f"Foto {verb}: {renamed_count}")
    print(f"Foto rimaste senza corrispondenza EAN->file: {len(remaining_files)}")
    print(f"EAN senza foto corrispondente: {len(unmatched_rows)}" + (f" -> {unmatched_report}" if unmatched_report else ""))
    print(f"RED FLAG (match troppo generico, >{RED_FLAG_THRESHOLD} candidati): {len(red_flag_rows)}" + (f" -> {red_flag_report}" if red_flag_report else ""))
    if renamed_report:
        print(f"Report rinomine: {renamed_report}")

    if red_flag_rows:
        print(
            "\nATTENZIONE: alcuni EAN hanno prodotto troppi candidati con un match parziale.\n"
            "Probabilmente le code_columns configurate sono troppo generiche: valuta di "
            "aggiungerne una (es. una colonna colore/taglia in più) e ripeti il dry-run."
        )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("season")
    parser.add_argument("brand")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--skip-optimize", action="store_true")
    parser.add_argument("--ean-column")
    parser.add_argument("--code-columns", help="lista separata da virgole, in ordine")
    args = parser.parse_args()

    cli_code_columns = None
    if args.code_columns:
        cli_code_columns = [c.strip() for c in args.code_columns.split(",") if c.strip()]

    ean_column, code_columns = resolve_config(args.brand, args.ean_column, cli_code_columns)

    start = time.time()
    run(args.season, args.brand, ean_column, code_columns, args.dry_run, args.skip_optimize)
    print(f"\nCompletato in {time.time() - start:.2f}s")


if __name__ == "__main__":
    main()
