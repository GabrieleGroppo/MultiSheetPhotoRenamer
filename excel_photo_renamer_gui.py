import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import importlib.util

# Import the original script functionality
def import_renaming_module():
    """Dynamically import the original renaming script"""
    script_path = os.path.join(os.path.dirname(__file__), 'multi_sheet_photo_renamer.py')
    spec = importlib.util.spec_from_file_location("multi_sheet_photo_renamer", script_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

class PhotoRenamerGUI:
    def __init__(self, master):
        self.master = master
        master.title("Multi Sheet Awesome Photo Renamer")
        master.geometry("600x700")

        # Import the original module
        self.renamer_module = import_renaming_module()

        # Create main frame
        self.main_frame = ttk.Frame(master, padding="10 10 10 10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

        # Season input
        ttk.Label(self.main_frame, text="Season (e.g., pe25):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.season_var = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.season_var, width=20).grid(row=0, column=1, sticky=tk.W, pady=5)

        # Brand selection
        ttk.Label(self.main_frame, text="Brand:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.brand_var = tk.StringVar()
        self.brand_dropdown = ttk.Combobox(
            self.main_frame, 
            textvariable=self.brand_var, 
            values=list(self.renamer_module.BRAND_COLUMN_MAPPINGS.keys()),
            width=20,
            state="readonly"
        )
        self.brand_dropdown.grid(row=1, column=1, sticky=tk.W, pady=5)

        # Photo folder selection
        ttk.Label(self.main_frame, text="Photo Folder:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.photo_folder_var = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.photo_folder_var, width=40).grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=self.browse_photo_folder).grid(row=2, column=2, sticky=tk.W, pady=5)

        # Excel file selection
        ttk.Label(self.main_frame, text="Excel File:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.excel_file_var = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.excel_file_var, width=40).grid(row=3, column=1, sticky=tk.W, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=self.browse_excel_file).grid(row=3, column=2, sticky=tk.W, pady=5)

        # Output folder selection
        ttk.Label(self.main_frame, text="Output Folder:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.output_folder_var = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.output_folder_var, width=40).grid(row=4, column=1, sticky=tk.W, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=self.browse_output_folder).grid(row=4, column=2, sticky=tk.W, pady=5)

        # Column selection
        ttk.Label(self.main_frame, text="Columns to Match:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.columns_listbox = tk.Listbox(self.main_frame, selectmode=tk.MULTIPLE, width=40, height=6)
        self.columns_listbox.grid(row=5, column=1, sticky=tk.W, pady=5)
        ttk.Button(self.main_frame, text="Load Columns", command=self.load_excel_columns).grid(row=5, column=2, sticky=tk.W, pady=5)

        # Optimize images checkbox
        self.optimize_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.main_frame, text="Optimize Images", variable=self.optimize_var).grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5)

        # Run button
        ttk.Button(self.main_frame, text="Rename Photos", command=self.rename_photos).grid(row=7, column=0, columnspan=3, pady=10)

        # Log area
        self.log_text = tk.Text(self.main_frame, width=70, height=10, wrap=tk.WORD)
        self.log_text.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        log_scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scrollbar.grid(row=8, column=3, sticky=(tk.N, tk.S))
        self.log_text.configure(yscroll=log_scrollbar.set)

    def browse_photo_folder(self):
        """Open folder selection dialog for photos"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.photo_folder_var.set(folder_selected)

    def browse_excel_file(self):
        """Open file selection dialog for Excel file"""
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_selected:
            self.excel_file_var.set(file_selected)
            # Automatically try to load columns when Excel file is selected
            self.load_excel_columns()

    def browse_output_folder(self):
        """Open folder selection dialog for output"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_var.set(folder_selected)

    def load_excel_columns(self):
        """Load columns from the selected Excel file"""
        excel_file = self.excel_file_var.get()
        if not excel_file:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return

        try:
            # Read the first sheet to get column names
            df = pd.read_excel(excel_file, nrows=0)
            
            # Clear existing items
            self.columns_listbox.delete(0, tk.END)
            
            # Add columns to listbox
            for column in df.columns:
                self.columns_listbox.insert(tk.END, column)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel file: {str(e)}")

    def rename_photos(self):
        """Perform photo renaming with GUI inputs"""
        # Validate inputs
        season = self.season_var.get().strip()
        brand = self.brand_var.get()
        photo_folder = self.photo_folder_var.get()
        excel_file = self.excel_file_var.get()
        output_folder = self.output_folder_var.get()

        # Validate required fields
        if not all([season, brand, photo_folder, excel_file]):
            messagebox.showwarning("Warning", "Please fill in all required fields.")
            return

        # Get selected columns
        selected_columns_indices = self.columns_listbox.curselection()
        if not selected_columns_indices:
            # If no columns selected, use default for the brand
            selected_columns = self.renamer_module.BRAND_COLUMN_MAPPINGS.get(brand, [])
        else:
            # Get selected column names
            selected_columns = [self.columns_listbox.get(i) for i in selected_columns_indices]

        # Redirect stdout to log area
        import io
        import sys
        log_capture = io.StringIO()
        sys.stdout = log_capture

        try:
            # Create output folder if it doesn't exist
            if output_folder and not os.path.exists(output_folder):
                os.makedirs(output_folder)

            # Custom renaming function that uses GUI inputs
            def custom_rename_func():
                # Create a base directory if not exists
                base_dir = f"./{season}"
                if not os.path.exists(base_dir):
                    os.makedirs(base_dir)

                # Prepare paths
                output_photos_dir = os.path.join(base_dir, self.renamer_module.DEFAULT_PHOTOS_SUBDIR, brand)
                reports_dir = os.path.join(base_dir, self.renamer_module.DEFAULT_REPORTS_SUBDIR)

                # Copy photos to output directory if specified
                if output_folder:
                    import shutil
                    if not os.path.exists(output_photos_dir):
                        os.makedirs(output_photos_dir)
                    for filename in os.listdir(photo_folder):
                        if filename.lower().endswith(self.renamer_module.FILE_EXTENSION):
                            shutil.copy(
                                os.path.join(photo_folder, filename), 
                                os.path.join(output_photos_dir, filename)
                            )
                    photo_folder_to_use = output_photos_dir
                else:
                    photo_folder_to_use = photo_folder

                # Prepare Excel file
                excel_output_dir = os.path.join(base_dir, self.renamer_module.DEFAULT_EXCELS_SUBDIR)
                if not os.path.exists(excel_output_dir):
                    os.makedirs(excel_output_dir)
                excel_output_path = os.path.join(excel_output_dir, f"{brand}.xlsx")
                shutil.copy(excel_file, excel_output_path)

                # Run image optimization if checkbox is checked
                if self.optimize_var.get():
                    self.renamer_module.optimize_images_in_folder(photo_folder_to_use)

                # Call the original renaming function with modified inputs
                return self.renamer_module.rinomina_foto_in_batch(
                    season=season, 
                    brand_name=brand, 
                    photo_folder=photo_folder_to_use,
                    excel_file=excel_output_path,
                    columns_to_match=selected_columns,
                    reports_folder=reports_dir
                )

            # Run the renaming process
            custom_rename_func()

            # Show success message
            messagebox.showinfo("Success", "Photo renaming completed successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            # Restore stdout and show log
            sys.stdout = sys.__stdout__
            log_output = log_capture.getvalue()
            self.log_text.delete(1.0, tk.END)
            self.log_text.insert(tk.END, log_output)

def main():
    """Main function to launch the GUI"""
    root = tk.Tk()
    gui = PhotoRenamerGUI(root)
    root.mainloop()

def run_with_cli_args():
    """Run the script with command-line arguments if provided"""
    if len(sys.argv) > 2:
        # If CLI arguments are provided, use the original script's main function
        import importlib.util
        spec = importlib.util.spec_from_file_location("multi_sheet_photo_renamer", 'multi_sheet_photo_renamer.py')
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        module.main()
    else:
        # Otherwise, launch the GUI
        main()

if __name__ == "__main__":
    run_with_cli_args()
