import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import importlib.util
import logging
from datetime import datetime

def main():
    root = tk.Tk()
    PhotoRenamerGUI(root)
    root.mainloop()

class PhotoRenamerGUI:
    def __init__(self, master):
        # Configure logging
        self.setup_logging()
        # Main window setup
        self.master = master
        master.title("Multi Sheet Awesome Photo Renamer")
        master.geometry("700x800")

        # Import the Multi Sheet Awesome Photo Renamer module
        try:
            self.renamer_module = self.import_renaming_module()
        except Exception as e:
            messagebox.showerror("Import Error", f"Could not import renaming module: {e}")
            sys.exit(1)

        # Create main frame
        self.create_ui_components()

    # Setup logging
    def setup_logging(self):
        """Set up logging configuration"""
        # Create logs directory if it doesn't exist
        log_dir = os.path.join(os.path.dirname(__file__), 'logs')
        os.makedirs(log_dir, exist_ok=True)

        # Create log filename with timestamp
        log_filename = os.path.join(log_dir, f'photo_renamer_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')

        # Configure logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s: %(message)s',
            handlers=[
                logging.FileHandler(log_filename),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    # Import the renaming module Multi Sheet Awesome Photo Renamer
    def import_renaming_module(self):
        try:
            script_path = os.path.join(os.path.dirname(__file__), 'multi_sheet_photo_renamer.py')
            spec = importlib.util.spec_from_file_location("multi_sheet_photo_renamer", script_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            self.logger.info("Successfully imported renaming module")
            return module
        except Exception as e:
            self.logger.error(f"Failed to import renaming module: {e}")
            raise

    # Create input fields and labels
    def create_ui_components(self):
        # Main frame
        self.main_frame = ttk.Frame(self.master, padding="10 10 10 10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

        # Variables
        self.season_var = tk.StringVar()
        self.brand_var = tk.StringVar()
        self.photo_folder_var = tk.StringVar()
        self.excel_file_var = tk.StringVar()
        self.optimize_var = tk.BooleanVar(value=True)

        # Create input fields and labels
        self.create_input_fields()

        # Create log area
        self.create_log_area()

    def create_input_fields(self):
        """Create input fields for the GUI"""
        # Season input
        ttk.Label(self.main_frame, text="Season (e.g., pe25):").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.main_frame, textvariable=self.season_var, width=30).grid(row=0, column=1, sticky=tk.W, pady=5)

        # Brand selection
        ttk.Label(self.main_frame, text="Brand:").grid(row=1, column=0, sticky=tk.W, pady=5)
        brand_dropdown = ttk.Combobox(
            self.main_frame, 
            textvariable=self.brand_var, 
            values=list(self.renamer_module.BRAND_COLUMN_MAPPINGS.keys()),
            width=30,
            state="readonly"
        )
        brand_dropdown.grid(row=1, column=1, sticky=tk.W, pady=5)

        # Photo folder selection
        ttk.Label(self.main_frame, text="Photo Folder:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.main_frame, textvariable=self.photo_folder_var, width=40).grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=self.browse_photo_folder).grid(row=2, column=2, sticky=tk.W, pady=5)

        # Excel file selection
        ttk.Label(self.main_frame, text="Excel File:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.main_frame, textvariable=self.excel_file_var, width=40).grid(row=3, column=1, sticky=tk.W, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=self.browse_excel_file).grid(row=3, column=2, sticky=tk.W, pady=5)

        # Columns selection
        ttk.Label(self.main_frame, text="Columns to Match:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.columns_listbox = tk.Listbox(self.main_frame, selectmode=tk.MULTIPLE, width=40, height=6)
        self.columns_listbox.grid(row=5, column=1, sticky=tk.W, pady=5)
        ttk.Button(self.main_frame, text="Load Columns", command=self.load_excel_columns).grid(row=5, column=2, sticky=tk.W, pady=5)

        # Optimize images checkbox
        ttk.Checkbutton(self.main_frame, text="Optimize Images", variable=self.optimize_var).grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5)

        # Run button
        ttk.Button(self.main_frame, text="Rename Photos", command=self.print_selected_fields).grid(row=7, column=0, columnspan=3, pady=10)

    def create_log_area(self):
        self.log_text = tk.Text(self.main_frame, width=80, height=15, wrap=tk.WORD)
        self.log_text.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        log_scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scrollbar.grid(row=8, column=3, sticky=(tk.N, tk.S))
        self.log_text.configure(yscroll=log_scrollbar.set)

    # Load folder containing photos
    def browse_photo_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.photo_folder_var.set(folder_selected)
            self.logger.info(f"Photo folder selected: {folder_selected}")

    # Load Excel file
    def browse_excel_file(self):
        """Open file selection dialog for Excel file"""
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_selected:
            self.excel_file_var.set(file_selected)
            self.load_excel_columns()
            self.logger.info(f"Excel file selected: {file_selected}")

    # Load columns from Excel file
    def load_excel_columns(self):
        excel_file = self.excel_file_var.get()
        if not excel_file:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return

        try:
            df = pd.read_excel(excel_file, nrows=0)
            
            # Clear existing items
            self.columns_listbox.delete(0, tk.END)
            
            # Add columns to listbox
            for column in df.columns:
                self.columns_listbox.insert(tk.END, column)
            
            self.logger.info(f"Loaded columns from {excel_file}")
        except Exception as e:
            self.logger.error(f"Could not read Excel file: {e}")
            messagebox.showerror("Error", f"Could not read Excel file: {e}")

    def print_selected_fields(self):
        season = self.season_var.get().strip()
        brand = self.brand_var.get()
        photo_folder = self.photo_folder_var.get()
        excel_file = self.excel_file_var.get()

        self.logger.info(f"Season: {season}")
        self.logger.info(f"Brand: {brand}")
        self.logger.info(f"Photo Folder: {photo_folder}")
        self.logger.info(f"Excel File: {excel_file}")

# Main function
if __name__ == "__main__":
    main()