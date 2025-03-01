import os
import pandas as pd
import time
import sys
import csv
import subprocess
from datetime import datetime
# This is the: MULTI SHEET AWESOME PHOTO RENAMER [MSAFR for friends]
# Constants
DEFAULT_REPORTS_SUBDIR = "reports"
DEFAULT_PHOTOS_SUBDIR = "photoes"
DEFAULT_EXCELS_SUBDIR = "excels"
FILE_EXTENSION = ".jpg"

# Brand column mappings
BRAND_COLUMN_MAPPINGS = {
    "guess": ["Model", "Part", "Color"],
    "liujo": ["Modello", "Parte", "Colore"],
    "furla": ["Modello", "Parte", "Colore", "TipoVariante"],
    "alviero": ["Linea", "Modello", "Tessuto", "Colore"],
    "brand": ["Campo1", "Campo2"]
}

# EAN column name
EAN_COLUMN = "EAN"

# Image optimization settings
MAX_SIZE_MB = 1
JPEG_QUALITY = 85


def optimize_jpeg_image(file_path):
    """
    Optimize JPEG image if it's larger than MAX_SIZE_MB.
    
    Args:
        file_path (str): Path to the JPEG file
        
    Returns:
        bool: True if the file was optimized, False otherwise
    """
    # Check if file exists
    if not os.path.isfile(file_path):
        return False
        
    # Get file size in MB
    file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
    
    # Only optimize if file size is greater than MAX_SIZE_MB
    if file_size_mb <= MAX_SIZE_MB:
        return False
    
    try:
        # Run jpegoptim with specified quality
        command = ["jpegoptim", "--max=" + str(JPEG_QUALITY), "--strip-all", file_path]
        process = subprocess.run(command, capture_output=True, text=True)
        
        # Check if the command was successful
        if process.returncode == 0:
            # Get new file size
            new_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            reduction = (file_size_mb - new_size_mb) / file_size_mb * 100
            return (True, file_size_mb, new_size_mb, reduction)
        else:
            print(f"Errore durante l'ottimizzazione di {file_path}: {process.stderr}")
            return False
    except Exception as e:
        print(f"Errore durante l'ottimizzazione di {file_path}: {e}")
        return False


def optimize_images_in_folder(folder_path):
    """
    Optimize all JPEG images in the folder that are larger than MAX_SIZE_MB.
    
    Args:
        folder_path (str): Path to the folder containing the images
        
    Returns:
        tuple: (total_images, optimized_images, total_reduction_mb)
    """
    if not os.path.isdir(folder_path):
        print(f"Errore: La cartella '{folder_path}' non esiste.")
        return (0, 0, 0)
    
    print(f"Ottimizzazione immagini nella cartella {folder_path}...")
    
    total_images = 0
    optimized_images = 0
    total_reduction_mb = 0
    
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(FILE_EXTENSION):
            total_images += 1
            file_path = os.path.join(folder_path, filename)
            
            optimization_result = optimize_jpeg_image(file_path)
            if optimization_result and optimization_result is not False:
                optimized = True
                original_size, new_size, reduction = optimization_result[1], optimization_result[2], optimization_result[3]
                optimized_images += 1
                reduction_mb = original_size - new_size
                total_reduction_mb += reduction_mb
                print(f"Ottimizzato: {filename} - Da {original_size:.2f}MB a {new_size:.2f}MB ({reduction:.1f}% riduzione)")
    
    print(f"\nRiepilogo ottimizzazione:")
    print(f"Immagini totali: {total_images}")
    print(f"Immagini ottimizzate: {optimized_images}")
    print(f"Spazio risparmiato: {total_reduction_mb:.2f}MB")
    
    return (total_images, optimized_images, total_reduction_mb)


def parse_command_line_args():
    """
    Parse command line arguments for season and brand name.
    Returns:
        tuple: (season, brand_name)
    """
    if len(sys.argv) < 3:
        print("Errore: specificare --season --brand-name.")
        print("Esempio: python script.py pe25 liujo")
        sys.exit(1)

    season = sys.argv[1]
    brand_name = sys.argv[2]

    # Validate brand name
    if brand_name not in BRAND_COLUMN_MAPPINGS:
        print(f"Errore: brand '{brand_name}' non riconosciuto.")
        print(f"Brand disponibili: {', '.join(BRAND_COLUMN_MAPPINGS.keys())}")
        sys.exit(1)

    return season, brand_name


def setup_paths(season, brand_name):
    """
    Set up the necessary paths for the script.
    
    Args:
        season (str): Season identifier (e.g., 'pe25')
        brand_name (str): Brand name
        
    Returns:
        tuple: (photo_folder, excel_file, reports_folder)
    """
    base_dir = f"./{season}"
    photo_folder = f"{base_dir}/{DEFAULT_PHOTOS_SUBDIR}/{brand_name}/"
    excel_file = f"{base_dir}/{DEFAULT_EXCELS_SUBDIR}/{brand_name}.xlsx"
    reports_folder = f"{base_dir}/{DEFAULT_REPORTS_SUBDIR}"
    
    return photo_folder, excel_file, reports_folder


def make_report_dir(reports_folder):
    """
    Create the reports directory if it doesn't exist.
    
    Args:
        reports_folder (str): Path to the reports folder
    """
    if not os.path.exists(reports_folder):
        os.makedirs(reports_folder)
        print(f"Cartella reports creata: {reports_folder}")


def rename_files(files, ean, indice_file, photo_folder, file_rinominati):
    """
    Rename a batch of files with the given EAN code.
    
    Args:
        files (list): List of files to rename
        ean (str): EAN code to use in the new filenames
        indice_file (dict): Index of files and their lowercase versions
        photo_folder (str): Folder containing the photos
        file_rinominati (set): Set of already renamed files
        
    Returns:
        int: Number of successfully renamed files
    """
    print(f"Rinomino {len(files)} file con EAN {ean}...")
    conteggio_rinominati = 0
    
    # Rename the found files
    for i, vecchio_nome in enumerate(files):
        estensione = os.path.splitext(vecchio_nome)[1]
        nuovo_nome = f"{ean}-{i}{estensione}"
        vecchio_percorso = os.path.join(photo_folder, vecchio_nome)
        nuovo_percorso = os.path.join(photo_folder, nuovo_nome)
        
        try:
            os.rename(vecchio_percorso, nuovo_percorso)
            file_rinominati.add(vecchio_nome)
            conteggio_rinominati += 1
            # Remove the renamed file from the index
            del indice_file[vecchio_nome]
        except Exception as e:
            print(f"Errore durante il rinomino del file {vecchio_nome}: {e}")
    
    return conteggio_rinominati


def print_results(conteggio_rinominati, file_jpg, elapsed_time):
    """
    Print the final results and statistics to the screen.
    
    Args:
        conteggio_rinominati (int): Number of renamed files
        file_jpg (list): List of all JPG files
        elapsed_time (float): Elapsed time in seconds
    """
    print(f"\nOperazione completata in {elapsed_time:.2f} secondi.")
    print(f"Totale file rinominati: {conteggio_rinominati}")
    print(f"File non rinominati: {len(file_jpg) - conteggio_rinominati}")


def make_report(conteggio_rinominati, ean_non_trovati, reports_folder, brand_name, columns_to_match):
    """
    Create a final report of EAN codes that weren't found.
    
    Args:
        conteggio_rinominati (int): Number of renamed files
        ean_non_trovati (list): List of dictionaries with information about EANs not found
        reports_folder (str): Path to the reports folder
        brand_name (str): Brand name
        columns_to_match (list): Columns to match from the Excel file
    """
    # Get the current date and time for the report
    now = datetime.now()
    datetime_str = now.strftime("%Y-%m-%d_%H-%M-%S")
    
    # Create a report of EAN codes not found
    if ean_non_trovati:
        report_file = f"{reports_folder}/report_ean_non_trovati_{brand_name}_{datetime_str}.csv"
        print(f"\nCreazione report dei codici EAN non trovati: {report_file}")
        
        # Determine the CSV file headers
        fieldnames = ['riga_excel', 'ean', 'data'] + columns_to_match
        
        with open(report_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(ean_non_trovati)
            
        print(f"Totale codici EAN senza corrispondenza: {len(ean_non_trovati)}")
        print(f"Il report è stato salvato in: {report_file}")
    else:
        print("\nTutti i codici EAN hanno trovato corrispondenza con almeno un file.")


def process_sheet(sheet_name, excel_file, indice_file, columns_to_match, file_rinominati, 
                 photo_folder, ean_non_trovati, codeno, codesi):
    """
    Process a single Excel sheet.
    
    Args:
        sheet_name (str): Name of the Excel sheet
        excel_file (str): Path to the Excel file
        indice_file (dict): Index of files and their lowercase versions
        columns_to_match (list): Columns to match from the Excel file
        file_rinominati (set): Set of already renamed files
        photo_folder (str): Folder containing the photos
        ean_non_trovati (list): List to store information about EANs not found
        codeno (list): List to store EANs without corresponding files
        codesi (list): List to store EANs with corresponding files
        
    Returns:
        int: Number of renamed files
    """
    renamed = 0
    
    # Only read the columns we need to save memory
    needed_columns = columns_to_match + [EAN_COLUMN]
    df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str, usecols=needed_columns)
    print(f"Foglio Excel letto: {len(df)} righe trovate")
    
    # Verify that each column to match is present in the dataframe
    for colonna in columns_to_match:
        if colonna not in df.columns:
            print(f"Errore: la colonna '{colonna}' non è presente nel file Excel.")
            sys.exit(1)

    # Verify that the EAN column is present in the dataframe
    if EAN_COLUMN not in df.columns:
        print(f"Errore: la colonna '{EAN_COLUMN}' non è presente nel file Excel.")
        sys.exit(1)
    
    # For each row in the dataframe
    for idx, riga in df.iterrows():
        if idx % 100 == 0:
            print(f"- Elaborazione riga {idx}...")
        
        # Extract values from columns of interest
        valori = {}
        for colonna in columns_to_match:
            if colonna in riga and pd.notna(riga[colonna]):
                valori[colonna] = str(riga[colonna]).lower().strip()
        
        # Get the EAN value
        if EAN_COLUMN in riga and pd.notna(riga[EAN_COLUMN]):
            ean = str(riga[EAN_COLUMN]).strip()
        else:
            continue  # Skip this row if there's no EAN
        
        # Find corresponding files
        file_corrispondenti = []
        for nome_file, nome_file_lower in list(indice_file.items()):
            if nome_file in file_rinominati:
                continue  # Skip files already renamed
                
            # Check if all values are in the filename
            tutti_presenti = True
            for valore in valori.values():
                if valore and valore not in nome_file_lower:
                    tutti_presenti = False
                    break
            
            if tutti_presenti:
                file_corrispondenti.append(nome_file)
        
        # If there are no corresponding files, record the EAN not found
        if not file_corrispondenti:
            # Create a dictionary with the relevant row data
            info_riga = {
                'riga_excel': idx + 1,  # +1 because indices start at 0
                'ean': ean,
                'data': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            # Add the values of the columns of interest
            for colonna in columns_to_match:
                if colonna in riga and pd.notna(riga[colonna]):
                    info_riga[colonna] = riga[colonna]
            
            ean_non_trovati.append(info_riga)
            codeno.append(ean)
        else:
            codesi.append(ean)
            renamed += rename_files(
                files=file_corrispondenti, 
                ean=ean, 
                indice_file=indice_file, 
                photo_folder=photo_folder, 
                file_rinominati=file_rinominati
            )
    
    return renamed


def rinomina_foto_in_batch(season, brand_name):
    """
    Rename photos in batch based on EAN codes from an Excel file.
    
    Args:
        season (str): Season identifier (e.g., 'pe25')
        brand_name (str): Brand name
    """
    # Setup paths and variables
    photo_folder, excel_file, reports_folder = setup_paths(season, brand_name)
    columns_to_match = BRAND_COLUMN_MAPPINGS[brand_name]
    print(f"Colonne da utilizzare: {columns_to_match}")
    
    # Initialize tracking variables
    file_rinominati = set()
    ean_non_trovati = []
    codeno = []
    codesi = []
    conteggio_rinominati = 0
    
    # Create the reports folder if it doesn't exist
    make_report_dir(reports_folder)
    
    # Start the timer
    start_time = time.time()

    # Optimize images before renaming
    print("\nVerifica e ottimizzazione immagini di grandi dimensioni...")
    optimize_images_in_folder(photo_folder)
    
    # Read all jpg files from the folder at once
    print(f"\nLettura dei file dalla cartella {photo_folder}...")
    if not os.path.isdir(photo_folder):
        print(f"Errore: La cartella '{photo_folder}' non esiste.")
        return
    
    tutti_file = os.listdir(photo_folder)
    file_jpg = [file for file in tutti_file if file.lower().endswith(FILE_EXTENSION)]
    print(f"Trovati {len(file_jpg)} file JPG")
    
    # Create an index for quick access to files
    indice_file = {file: file.lower() for file in file_jpg}

    # Start reading the Excel file
    # Read the file one sheet at a time
    print(f"Lettura del file Excel {excel_file}...")
    try:
        excel = pd.ExcelFile(excel_file)
        for sheet in excel.sheet_names:
            print(f"Elaborazione foglio {sheet}...")
            conteggio_rinominati += process_sheet(
                sheet_name=sheet, 
                excel_file=excel_file, 
                indice_file=indice_file, 
                columns_to_match=columns_to_match, 
                file_rinominati=file_rinominati, 
                photo_folder=photo_folder, 
                ean_non_trovati=ean_non_trovati,
                codeno=codeno,
                codesi=codesi
            )
    except FileNotFoundError:
        print(f"Errore: Il file Excel '{excel_file}' non esiste.")
        return
    except Exception as e:
        print(f"Errore durante la lettura del file Excel: {e}")
        return

    # End the timer
    elapsed_time = time.time() - start_time
    
    # Print the final report
    print_results(conteggio_rinominati=conteggio_rinominati, file_jpg=file_jpg, elapsed_time=elapsed_time)

    # Create a report of EAN codes not found
    make_report(conteggio_rinominati, ean_non_trovati, reports_folder, brand_name, columns_to_match)


def main():
    """Main function to execute the script."""
    season, brand_name = parse_command_line_args()
    rinomina_foto_in_batch(season, brand_name)


if __name__ == "__main__":
    main()
