"""Helper condivisi per individuare cartelle/file di una coppia (stagione, brand)."""
import glob
import os
import re
import sys

PHOTO_EXTENSIONS = (".jpg", ".jpeg")


def brand_dir_path(season, brand):
    return os.path.join("assets", season, brand)


def require_brand_dir(season, brand):
    path = brand_dir_path(season, brand)
    if not os.path.isdir(path):
        sys.exit(
            f"Cartella non trovata: {path}\n"
            f"Struttura attesa: assets/<stagione>/<brand>/ con dentro le foto .jpg e il file .xlsx"
        )
    return path


def find_excel(brand_dir):
    candidates = [
        c for c in glob.glob(os.path.join(brand_dir, "*.xlsx")) if not os.path.basename(c).startswith("~$")
    ]
    if not candidates:
        sys.exit(f"Nessun file .xlsx trovato in {brand_dir}")
    if len(candidates) > 1:
        print(f"ATTENZIONE: più file .xlsx trovati, uso il primo: {candidates}", file=sys.stderr)
    return candidates[0]


def list_photos(brand_dir):
    return sorted(f for f in os.listdir(brand_dir) if f.lower().endswith(PHOTO_EXTENSIONS))


def code_token(filename):
    """Estrae il token 'codice' candidato dal nome file (parte prima del primo separatore)."""
    stem = os.path.splitext(filename)[0]
    return re.split(r"[_\-\s]", stem, maxsplit=1)[0]
