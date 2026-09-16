"""Lettura/scrittura della configurazione persistente per brand.

La configurazione (colonna EAN + colonne che compongono il codice prodotto)
viene salvata in .claude/skills/photo-renamer/brand_configs/<brand>.json,
FUORI da assets/ (che e' materiale cliente, gitignored) cosi' viene versionata
e riutilizzata anche quando la cartella assets/ cambia da una stagione
all'altra.
"""
import json
import os

CONFIG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "brand_configs")


def config_path(brand):
    return os.path.join(CONFIG_DIR, f"{brand.lower()}.json")


def load_config(brand):
    path = config_path(brand)
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_config(brand, ean_column, code_columns, notes=""):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    path = config_path(brand)
    config = {
        "ean_column": ean_column,
        "code_columns": code_columns,
        "notes": notes,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return path
