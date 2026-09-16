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


def save_config(brand, ean_column, code_columns, notes="", swatch_position=None):
    """swatch_position: "first"/"last" se in questo brand una foto per prodotto è
    sempre solo un campione colore (senza borsa/gadget) in quella posizione della
    serie, None/"none" se e' stato verificato che non serve (o non ancora verificato)."""
    os.makedirs(CONFIG_DIR, exist_ok=True)
    path = config_path(brand)

    existing = load_config(brand) or {}
    if swatch_position is None:
        swatch_position = existing.get("swatch_position")
    elif swatch_position == "none":
        swatch_position = None

    config = {
        "ean_column": ean_column,
        "code_columns": code_columns,
        "swatch_position": swatch_position,
        "notes": notes or existing.get("notes", ""),
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return path
