#!/usr/bin/env python3
"""Stampa struttura Excel + campione nomi foto per un brand/stagione.

Non decide nulla da solo: serve come base di osservazione per chi (Claude)
deve dedurre quale colonna e' l'EAN e quali colonne compongono il codice
prodotto usato nei nomi dei file.

Uso: inspect_brand.py <stagione> <brand> [--sample N]
"""
import argparse
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from discovery import code_token, find_excel, list_photos, require_brand_dir
from xlsx_reader import read_workbook


def digit_length_profile(values):
    lengths = Counter()
    digit_count = 0
    total = 0
    for v in values:
        if v in (None, ""):
            continue
        total += 1
        s = str(v).strip()
        if s.isdigit():
            digit_count += 1
            lengths[len(s)] += 1
    return total, digit_count, lengths.most_common(3)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("season")
    parser.add_argument("brand")
    parser.add_argument("--sample", type=int, default=8, help="righe/foto di esempio da mostrare")
    args = parser.parse_args()

    brand_dir = require_brand_dir(args.season, args.brand)
    excel_path = find_excel(brand_dir)
    photos = list_photos(brand_dir)

    print(f"=== Cartella: {brand_dir} ===")
    print(f"Excel: {os.path.basename(excel_path)}")
    print(f"Foto trovate: {len(photos)}\n")

    show_n = min(args.sample * 3, len(photos))
    print(f"=== Campione nomi file foto (primi {show_n}) ===")
    for f in photos[:show_n]:
        print(f"  {f:50s} -> codice candidato: {code_token(f)}")
    print()

    workbook = read_workbook(excel_path)
    print(f"=== Fogli Excel trovati: {list(workbook.keys())} ===\n")

    for sheet_name, records in workbook.items():
        if not records:
            print(f"--- Foglio '{sheet_name}': vuoto ---\n")
            continue

        headers = list(records[0].keys())
        print(f"--- Foglio '{sheet_name}': {len(records)} righe, {len(headers)} colonne ---")
        print("Colonne:", headers, "\n")

        print(f"Prime {args.sample} righe (tutte le colonne):")
        for r in records[: args.sample]:
            print("  ", {h: r.get(h) for h in headers})
        print()

        print("Statistiche per colonna (aiutano a individuare la colonna EAN e le colonne codice):")
        for h in headers:
            values = [r.get(h) for r in records]
            total, digit_count, top_lengths = digit_length_profile(values)
            sample_vals = []
            for v in values:
                if v not in (None, "") and v not in sample_vals:
                    sample_vals.append(v)
                if len(sample_vals) >= 4:
                    break
            digit_pct = (digit_count / total * 100) if total else 0
            print(
                f"  - {h!r}: non vuoti={total}, %solo-cifre={digit_pct:.0f}%, "
                f"lunghezze più comuni={top_lengths}, esempi={sample_vals}"
            )
        print()


if __name__ == "__main__":
    main()
