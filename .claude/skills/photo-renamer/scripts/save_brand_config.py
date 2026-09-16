#!/usr/bin/env python3
"""Salva/aggiorna la configurazione (colonna EAN + colonne codice) di un brand.

Uso:
  save_brand_config.py <brand> --ean-column BARCODE --code-columns "Modello,Parte,Colore" [--notes "..."]
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brand_config import save_config


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("brand")
    parser.add_argument("--ean-column", required=True)
    parser.add_argument("--code-columns", required=True, help="lista separata da virgole, in ordine")
    parser.add_argument("--notes", default="")
    args = parser.parse_args()

    code_columns = [c.strip() for c in args.code_columns.split(",") if c.strip()]
    path = save_config(args.brand, args.ean_column, code_columns, args.notes)
    print(f"Configurazione salvata in {path}")
    print(f"  ean_column   = {args.ean_column}")
    print(f"  code_columns = {code_columns}")


if __name__ == "__main__":
    main()
