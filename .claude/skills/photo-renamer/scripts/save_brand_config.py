#!/usr/bin/env python3
"""Salva/aggiorna la configurazione (colonna EAN + colonne codice + eventuale
posizione del campione colore) di un brand.

Uso:
  save_brand_config.py <brand> --ean-column BARCODE --code-columns "Modello,Parte,Colore" [--notes "..."]
  save_brand_config.py <brand> --swatch-position last   # solo per aggiornare questo campo
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brand_config import load_config, save_config


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("brand")
    parser.add_argument("--ean-column", help="richiesto la prima volta che si configura il brand")
    parser.add_argument("--code-columns", help="lista separata da virgole, in ordine; richiesto la prima volta")
    parser.add_argument("--notes", default="")
    parser.add_argument(
        "--swatch-position",
        choices=["first", "last", "none"],
        help=(
            "'first'/'last' se in questo brand una foto per prodotto e' sempre solo un "
            "campione colore (senza borsa/gadget visibile) in quella posizione della serie; "
            "'none' se e' stato verificato che questo brand NON ha questo problema"
        ),
    )
    args = parser.parse_args()

    existing = load_config(args.brand) or {}
    ean_column = args.ean_column or existing.get("ean_column")
    code_columns = (
        [c.strip() for c in args.code_columns.split(",") if c.strip()]
        if args.code_columns
        else existing.get("code_columns")
    )

    if not ean_column or not code_columns:
        sys.exit(
            f"Configurazione incompleta per '{args.brand}': servono --ean-column e --code-columns "
            f"(almeno la prima volta)."
        )

    path = save_config(
        args.brand, ean_column, code_columns, notes=args.notes, swatch_position=args.swatch_position
    )
    saved = load_config(args.brand)
    print(f"Configurazione salvata in {path}")
    print(f"  ean_column      = {saved['ean_column']}")
    print(f"  code_columns    = {saved['code_columns']}")
    print(f"  swatch_position = {saved['swatch_position']}")


if __name__ == "__main__":
    main()
